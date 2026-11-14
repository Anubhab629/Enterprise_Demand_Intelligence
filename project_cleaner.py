"""Controlled project maintenance utility.

This module is intentionally separate from the application pipeline.
It performs only explicitly-scoped cleanup of generated runtime artifacts
and approved empty project directories.

It is NOT imported by ``main.py`` and does not modify application source
code, datasets, Power BI integration logic, or the environment controller.
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import stat
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import env


# ---------------------------------------------------------------------------
# Cleanup policy
# ---------------------------------------------------------------------------

KEEP_RECENT_RUNS = 5
MIN_KEEP_RECENT_RUNS = 3
MAX_KEEP_RECENT_RUNS = 5

# Only these generated locations are eligible for run-artifact retention.
# Retention is evaluated independently for every location.
MANAGED_RUN_LOCATIONS = (
    Path("data/output/runs"),
    Path("powerbi/data"),
    Path("powerbi"),  # filtered below to Demand_Intelligence_Report_run_*
)

# Empty-directory cleanup is deliberately allowlisted.  In particular, we
# do not recursively remove arbitrary empty directories inside retained runs.
APPROVED_EMPTY_DIRECTORIES = (
    Path("data/companies"),
    Path("data/input"),
    Path("data/reference"),
    Path("data/uploads"),
    Path("backups"),
    Path("_archive"),
    Path("documentation"),
    Path("logs"),
    Path("outputs"),
)

# Explicitly protected paths.  The cleaner never traverses these as cleanup
# targets and never deletes them.
PROTECTED_RELATIVE_PATHS = (
    Path(".git"),
    Path(".env_backups"),
    Path("src"),
    Path("data/raw"),
    Path("data/processed"),
    Path("main.py"),
    Path("config.py"),
    Path("env.py"),
    Path("powerbi_integration.py"),
    Path(".gitignore"),
)

RUN_ID_RE = re.compile(r"run_(\d{8}T\d{6}_\d{6}Z_[0-9a-fA-F]+)$")
REPORT_RUN_RE = re.compile(
    r"Demand_Intelligence_Report_run_(\d{8}T\d{6}_\d{6}Z_[0-9a-fA-F]+)$"
)


@dataclass(frozen=True)
class CleanupCandidate:
    """A specific generated path selected for deletion."""

    location: Path
    path: Path
    run_id: str


@dataclass(frozen=True)
class CleanupPlan:
    """Read-only cleanup plan calculated before any deletion."""

    artifact_candidates: tuple[CleanupCandidate, ...]
    empty_directories: tuple[Path, ...]


# ---------------------------------------------------------------------------
# Path safety
# ---------------------------------------------------------------------------


def project_root() -> Path:
    """Return the project root from the existing environment controller."""
    return env.get_project_root().resolve()


def absolute_path(relative_path: Path) -> Path:
    """Resolve a project-relative path without allowing path escape."""
    root = project_root()
    resolved = (root / relative_path).resolve()
    if resolved != root and root not in resolved.parents:
        raise RuntimeError(f"Unsafe path outside project root: {relative_path}")
    return resolved


def is_protected(path: Path) -> bool:
    """Return True if *path* is or is inside an explicitly protected path."""
    root = project_root()
    resolved = path.resolve()
    for relative in PROTECTED_RELATIVE_PATHS:
        protected = (root / relative).resolve()
        if resolved == protected or protected in resolved.parents:
            return True
    return False

# OneDrive Cloud Files uses distinct confirmed tags for directory and file
# placeholders. Both were observed and verified in this project's filesystem.
CLOUD_FILES_REPARSE_TAGS = frozenset({
    0x9000E01A,  # directory placeholder
    0x9000401A,  # file placeholder
})


def get_reparse_tag(path: Path) -> int | None:
    """Return a Windows reparse tag; None means the path is not a reparse point.

    .NET is used for the attribute check because ordinary Win32 attribute
    queries returned incomplete metadata for this project's OneDrive folders.
    fsutil is used only to read the tag of an actual reparse point.
    """
    import base64
    import subprocess

    if os.name != "nt":
        raise RuntimeError("OneDrive reparse-point inspection requires Windows.")

    if path.is_symlink():
        raise RuntimeError(f"Symbolic links are not allowed: {path}")

    absolute = str(path.absolute())
    escaped = absolute.replace("'", "''")
    script = (
        f"$p = '{escaped}'; "
        "try { [int][IO.File]::GetAttributes($p) } "
        "catch { exit 20 }"
    )
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")

    try:
        attributes_result = subprocess.run(
            [
                "powershell.exe", "-NoProfile", "-NonInteractive",
                "-EncodedCommand", encoded,
            ],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(
            f"Unable to inspect filesystem attributes: {path}"
        ) from exc

    if attributes_result.returncode != 0:
        raise RuntimeError(f"Unable to inspect filesystem attributes: {path}")

    try:
        attributes = int(attributes_result.stdout.strip().splitlines()[-1])
    except (ValueError, IndexError) as exc:
        raise RuntimeError(
            f"Invalid filesystem attribute response for: {path}"
        ) from exc

    if not attributes & 0x400:
        return None

    try:
        tag_result = subprocess.run(
            ["fsutil.exe", "reparsepoint", "query", absolute],
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(
            f"Unable to inspect reparse-point tag: {path}"
        ) from exc

    if tag_result.returncode != 0:
        raise RuntimeError(
            f"Unable to inspect reparse-point tag: {path}: "
            f"{tag_result.stderr.strip()}"
        )

    import re

    match = re.search(
        r"Reparse Tag Value\s*:\s*(0x[0-9a-fA-F]+)",
        tag_result.stdout + "\n" + tag_result.stderr,
        flags=re.IGNORECASE,
    )
    if not match:
        raise RuntimeError(f"Could not determine reparse-point tag: {path}")

    return int(match.group(1), 16)


def is_symlink_or_reparse_point(path: Path) -> bool:
    """Detect any symlink or reparse point; fail closed on inspection errors."""
    try:
        return path.is_symlink() or get_reparse_tag(path) is not None
    except (OSError, RuntimeError):
        return True


def is_unsafe_reparse_point(path: Path) -> bool:
    """Reject symlinks and every reparse tag except the confirmed OneDrive tag."""
    if path.is_symlink():
        return True

    try:
        tag = get_reparse_tag(path)
    except (OSError, RuntimeError):
        return True

    return tag is not None and tag not in CLOUD_FILES_REPARSE_TAGS


def _batch_validate_reparse_points(root: Path) -> bool:
    """Inspect reparse tags for an entire candidate tree in one PowerShell run.

    Only the confirmed OneDrive Cloud Files tag is accepted. Any other tag,
    symlink, inaccessible item, or enumeration error rejects the whole tree.
    """
    import base64
    import subprocess

    if os.name != "nt":
        raise RuntimeError("OneDrive tree inspection requires Windows.")

    escaped = str(root.absolute()).replace("'", "''")
    script = r"""
$ErrorActionPreference = 'Stop'
$root = '__ROOT__'
try {
    $items = @(Get-Item -LiteralPath $root -Force -ErrorAction Stop)
    $items += @(Get-ChildItem -LiteralPath $root -Force -Recurse -ErrorAction Stop)
    foreach ($item in $items) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            $q = & fsutil.exe reparsepoint query $item.FullName 2>&1 | Out-String
            if ($LASTEXITCODE -ne 0) { Write-Output ('UNSAFE|' + $item.FullName); exit 41 }
            if ($q -notmatch '(?i)Reparse Tag Value\s*:\s*0x9000(e01a|401a)') {
                Write-Output ('UNSAFE|' + $item.FullName); exit 42
            }
        }
        if (($item.Attributes -band [IO.FileAttributes]::Directory) -eq 0 -and
            ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) {
            # Regular files are allowed; unsupported filesystem objects are
            # independently rejected by the Python structural walk below.
        }
    }
    Write-Output 'TREE_SAFE'
    exit 0
} catch {
    Write-Output ('INSPECTION_ERROR|' + $_.Exception.Message)
    exit 43
}
""".replace('__ROOT__', escaped)
    encoded = base64.b64encode(script.encode('utf-16le')).decode('ascii')
    try:
        result = subprocess.run(
            ['powershell.exe', '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
            capture_output=True, text=True, timeout=180, check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(f"Batch filesystem inspection failed: {root}") from exc
    if result.returncode != 0 or 'TREE_SAFE' not in result.stdout:
        detail = (result.stdout.strip() + ' ' + result.stderr.strip()).strip()
        raise RuntimeError(f"Unsafe or inaccessible item in generated tree {root}: {detail}")
    return True


def is_safe_generated_tree(path: Path) -> bool:
    """Validate a generated tree with one batched reparse-point inspection."""
    try:
        if path.is_symlink() or is_protected(path):
            return False
        root_info = path.lstat()
        if not stat.S_ISDIR(root_info.st_mode):
            return False

        # One PowerShell process per candidate tree, not one per file/folder.
        _batch_validate_reparse_points(path)

        stack = [path]
        while stack:
            current = stack.pop()
            if current.is_symlink() or is_protected(current):
                return False
            with os.scandir(current) as entries:
                for entry in entries:
                    child = Path(entry.path)
                    if entry.is_symlink() or is_protected(child):
                        return False
                    info = entry.stat(follow_symlinks=False)
                    if stat.S_ISDIR(info.st_mode):
                        stack.append(child)
                    elif not stat.S_ISREG(info.st_mode):
                        # Cloud Files placeholders can surface as reparse
                        # objects; the batch PowerShell check has validated
                        # their tag, so permit only those explicitly tagged.
                        if not (getattr(info, 'st_file_attributes', 0) & 0x400):
                            return False
        return True
    except (PermissionError, OSError, RuntimeError):
        return False


# ---------------------------------------------------------------------------
# Run discovery
# ---------------------------------------------------------------------------


def run_id_from_name(name: str) -> str | None:
    """Extract a canonical run ID from a generated run directory name."""
    match = RUN_ID_RE.fullmatch(name)
    if match:
        return match.group(1)

    match = REPORT_RUN_RE.fullmatch(name)
    if match:
        return match.group(1)

    return None


def discover_run_directories(location: Path) -> list[tuple[str, Path]]:
    """Discover recognized generated run directories in one managed location."""
    root = absolute_path(location)

    if not root.exists():
        return []

    if is_unsafe_reparse_point(root):
        raise RuntimeError(
            f"Managed cleanup location contains an unsupported reparse point: {root}"
        )

    if not root.is_dir():
        raise RuntimeError(
            f"Managed cleanup location is not a directory: {root}"
        )

    if is_protected(root):
        raise RuntimeError(
            f"Managed cleanup location is protected: {root}"
        )

    discovered: list[tuple[str, Path]] = []

    for child in root.iterdir():
        if is_unsafe_reparse_point(child):
            continue

        if not child.is_dir():
            continue

        run_id = run_id_from_name(child.name)

        if run_id is None:
            continue

        if is_protected(child):
            continue

        discovered.append((run_id, child))

    discovered.sort(key=lambda item: item[0], reverse=True)

    return discovered


def discover_all_managed_runs() -> dict[Path, list[tuple[str, Path]]]:
    """Inspect every explicitly managed generated-run location."""
    results: dict[Path, list[tuple[str, Path]]] = {}

    for location in MANAGED_RUN_LOCATIONS:
        root = absolute_path(location)

        if not root.exists():
            results[location] = []
            continue

        if is_unsafe_reparse_point(root):
            raise RuntimeError(
                f"Managed cleanup location contains an unsupported reparse point: {root}"
            )

        if not root.is_dir():
            raise RuntimeError(
                f"Managed cleanup location is not a directory: {root}"
            )

        if is_protected(root):
            raise RuntimeError(
                f"Managed cleanup location is protected: {root}"
            )

        discovered: list[tuple[str, Path]] = []

        for child in root.iterdir():
            if is_unsafe_reparse_point(child):
                continue

            if not child.is_dir():
                continue

            run_id = run_id_from_name(child.name)

            if run_id is None:
                continue

            if location == Path("powerbi"):
                if not child.name.startswith(
                    "Demand_Intelligence_Report_run_"
                ):
                    continue

            if is_protected(child):
                continue

            discovered.append((run_id, child))

        discovered.sort(key=lambda item: item[0], reverse=True)
        results[location] = discovered

    return results


# ---------------------------------------------------------------------------
# Plan generation
# ---------------------------------------------------------------------------


def validate_keep_count(keep: int) -> int:
    if not MIN_KEEP_RECENT_RUNS <= keep <= MAX_KEEP_RECENT_RUNS:
        raise ValueError(
            f"--keep must be between {MIN_KEEP_RECENT_RUNS} and "
            f"{MAX_KEEP_RECENT_RUNS}; received {keep}."
        )
    return keep


def calculate_artifact_candidates(
    inventories: dict[Path, list[tuple[str, Path]]],
    keep: int,
) -> tuple[CleanupCandidate, ...]:
    """
    Select older generated run directories only when their entire
    directory trees pass the safety inspection.
    """
    candidates: list[CleanupCandidate] = []

    for location, entries in inventories.items():
        older_entries = entries[keep:]
        for index, (run_id, path) in enumerate(older_entries, start=1):
            if run_id_from_name(path.name) != run_id:
                continue

            print(
                f"\nValidating older run {index}/{len(older_entries)} "
                f"in [{location}]: {path.name}",
                flush=True,
            )
            if not is_safe_generated_tree(path):
                print(f"SKIPPED UNSAFE RUN: {path}", flush=True)
                continue

            print(f"  Validation passed: {path.name}", flush=True)
            candidates.append(
                CleanupCandidate(
                    location=location,
                    path=path,
                    run_id=run_id,
                )
            )

    return tuple(candidates)


def discover_approved_empty_directories() -> tuple[Path, ...]:
    """Return only approved directories verified to be genuinely empty."""
    empty: list[Path] = []

    for relative in APPROVED_EMPTY_DIRECTORIES:
        path = absolute_path(relative)

        if not path.exists():
            continue

        if not path.is_dir():
            raise RuntimeError(
                f"Approved empty-directory target is not a directory: {path}"
            )

        if is_protected(path):
            raise RuntimeError(
                f"Approved empty-directory target is protected: {path}"
            )

        # Never treat symlinks or cloud-managed directories as empty.
        if path.is_symlink():
            continue

        try:
            attributes = getattr(path.lstat(), "st_file_attributes", 0)
            if attributes & 0x400:  # Windows ReparsePoint
                continue

            # Only include directories with absolutely no entries.
            with __import__("os").scandir(path) as entries:
                if next(entries, None) is None:
                    empty.append(path)

        except (PermissionError, OSError):
            # If emptiness cannot be verified, do not list it.
            continue

    return tuple(empty)


def build_plan(keep: int) -> CleanupPlan:
    """Perform inspection and calculate a deletion plan without deleting."""
    keep = validate_keep_count(keep)
    inventories = discover_all_managed_runs()
    candidates = calculate_artifact_candidates(inventories, keep)
    empty_directories = discover_approved_empty_directories()
    return CleanupPlan(
        artifact_candidates=candidates,
        empty_directories=empty_directories,
    )


# ---------------------------------------------------------------------------
# Reporting and confirmation
# ---------------------------------------------------------------------------


def display_inventory(
    inventories: dict[Path, list[tuple[str, Path]]],
    keep: int,
) -> None:
    root = project_root()

    print("\n" + "=" * 78)
    print("PROJECT CLEANER - INSPECTION")
    print("=" * 78)
    print(f"Project root : {root}")
    print(f"Keep recent  : {keep} runs per managed location")

    for location, entries in inventories.items():
        print(f"\n[{location}]")

        if not entries:
            print("  No recognized generated runs found.")
            continue

        for index, (run_id, path) in enumerate(entries, start=1):
            timestamp = run_id[:15]
            formatted_date = (
                f"{timestamp[6:8]}-{timestamp[4:6]}-{timestamp[:4]} "
                f"{timestamp[9:11]}:{timestamp[11:13]}:{timestamp[13:15]}"
            )

            marker = "KEEP" if index <= keep else "OLDER RUN"

            print(
                f"  {marker:<12} [{formatted_date}] {path}"
            )


def display_plan(plan: CleanupPlan) -> None:
    print("\n" + "=" * 78)
    print("CLEANUP PLAN")
    print("=" * 78)

    print("\nOld generated run directories eligible for deletion:")
    if plan.artifact_candidates:
        for candidate in plan.artifact_candidates:
            timestamp = candidate.run_id[:15]
            formatted_date = (
                f"{timestamp[6:8]}-{timestamp[4:6]}-{timestamp[:4]} "
                f"{timestamp[9:11]}:{timestamp[11:13]}:{timestamp[13:15]}"
            )

            print(
                f"  DELETE RUN DIRECTORY  [{formatted_date}] {candidate.path}"
            )
    else:
        print("  None")

    print("\nApproved empty directories:")
    if plan.empty_directories:
        for path in plan.empty_directories:
            print(f"  EMPTY   {path}")
    else:
        print("  None")


def confirm(prompt: str) -> bool:
    """Return True only for an explicit yes response."""
    answer = input(f"{prompt} [y/N]: ").strip().lower()
    return answer in {"y", "yes"}


# ---------------------------------------------------------------------------
# Deletion and verification
# ---------------------------------------------------------------------------


def safe_remove_directory(path: Path) -> None:
    """Delete one previously inspected directory after re-validating safety."""
    root = project_root()
    resolved = path.absolute()

    if not resolved.exists():
        raise RuntimeError(
            f"Cleanup target disappeared before deletion: {resolved}"
        )

    if not resolved.is_dir():
        raise RuntimeError(
            f"Cleanup target is no longer a directory: {resolved}"
        )

    if is_protected(resolved):
        raise RuntimeError(
            f"Refusing to delete protected path: {resolved}"
        )

    if resolved == root or root not in resolved.parents:
        raise RuntimeError(
            f"Refusing to delete path outside project root: {resolved}"
        )

    # Re-check emptiness immediately before deletion.
    if any(resolved.iterdir()):
        raise RuntimeError(
            f"Refusing to delete directory that is no longer empty: {resolved}"
        )

    # Windows may deny deletion of a directory carrying the Read-only
    # attribute. Remove only that attribute; do not alter ownership,
    # permissions, or other filesystem attributes.
    if sys.platform == "win32":
        import ctypes

        FILE_ATTRIBUTE_READONLY = 0x0001
        INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF

        attributes = ctypes.windll.kernel32.GetFileAttributesW(str(resolved))

        if attributes == INVALID_FILE_ATTRIBUTES:
            raise RuntimeError(
                f"Unable to inspect Windows attributes before deletion: {resolved}"
            )

        if attributes & FILE_ATTRIBUTE_READONLY:
            result = ctypes.windll.kernel32.SetFileAttributesW(
                str(resolved),
                attributes & ~FILE_ATTRIBUTE_READONLY,
            )

            if not result:
                raise RuntimeError(
                    f"Unable to remove Read-only attribute from: {resolved}"
                )

    resolved.rmdir()

    if resolved.exists():
        raise RuntimeError(
            f"Directory deletion could not be verified: {resolved}"
        )

def safe_remove_generated_run_directory(candidate: CleanupCandidate) -> None:
    """Revalidate and remove exactly one approved generated run directory."""
    root = project_root().absolute()

    if candidate.location not in MANAGED_RUN_LOCATIONS:
        raise RuntimeError(
            f"Refusing deletion from unmanaged location: {candidate.location}"
        )

    managed_root = (root / candidate.location).absolute()
    target = candidate.path.absolute()

    # The target must be one direct child of its exact managed directory.
    if target.parent != managed_root:
        raise RuntimeError(
            f"Refusing deletion outside the managed directory boundary: {target}"
        )

    if root not in target.parents or target == root:
        raise RuntimeError(f"Refusing deletion outside project root: {target}")

    parsed_run_id = run_id_from_name(target.name)

    if parsed_run_id is None or parsed_run_id != candidate.run_id:
        raise RuntimeError(
            f"Generated-run name or ID no longer matches the cleanup plan: {target}"
        )

    if is_protected(target):
        raise RuntimeError(f"Refusing to delete protected path: {target}")

    if not target.exists() or not target.is_dir():
        raise RuntimeError(f"Cleanup target is missing or is not a directory: {target}")

    # Re-run the complete structural and OneDrive reparse-tag validation
    # immediately before invoking the single-operation deletion routine.
    if not is_safe_generated_tree(target):
        raise RuntimeError(
            f"Generated run failed final safety validation; not deleting: {target}"
        )

    print(f"Deleting generated run directory: {target}", flush=True)
    _remove_generated_tree(target)

    if target.exists():
        raise RuntimeError(
            f"Generated-run deletion could not be verified: {target}"
        )

    print(f"  Deletion verified: {target.name}", flush=True)


def _remove_generated_tree(path: Path) -> None:
    """Delete one fully validated generated run tree in a single operation.

    The caller performs a complete safety inspection immediately before this
    function. PowerShell is used once for the whole directory so OneDrive
    Cloud Files entries are handled by the Windows filesystem provider rather
    than launching a process for every child item.
    """
    import base64
    import subprocess

    if os.name != "nt":
        raise RuntimeError("Generated-run deletion through OneDrive requires Windows.")
    if path.is_symlink():
        raise RuntimeError(f"Refusing to remove a symbolic link: {path}")

    escaped = str(path.absolute()).replace("'", "''")
    script = (
        "$ErrorActionPreference = 'Stop'; "
        f"$p = '{escaped}'; "
        "try { "
        "$item = Get-Item -LiteralPath $p -Force -ErrorAction Stop; "
        "if (-not $item.PSIsContainer) { throw 'Cleanup target is not a directory.' }; "
        "Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction Stop; "
        "if (Test-Path -LiteralPath $p) { exit 31 }; "
        "exit 0 "
        "} catch { Write-Error $_; exit 32 }"
    )
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")

    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded],
            capture_output=True,
            text=True,
            timeout=600,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(f"Generated-run deletion did not complete: {path}") from exc

    if result.returncode != 0 or path.exists():
        details = result.stderr.strip() or result.stdout.strip()
        raise RuntimeError(
            f"Generated-run deletion failed or could not be verified: {path}. {details}"
        )


def verify_artifact_deletions(candidates: Iterable[CleanupCandidate]) -> None:
    """Verify every approved artifact candidate was removed."""
    failures = [str(c.path) for c in candidates if c.path.exists()]
    if failures:
        raise RuntimeError(
            "Deletion verification failed; these candidates still exist:\n"
            + "\n".join(f"  {path}" for path in failures)
        )


def verify_empty_directory_deletions(paths: Iterable[Path]) -> None:
    """Verify every approved empty-directory target was removed."""
    failures = [str(path) for path in paths if path.exists()]
    if failures:
        raise RuntimeError(
            "Empty-directory deletion verification failed:\n"
            + "\n".join(f"  {path}" for path in failures)
        )


def verify_retained_runs(
    inventories: dict[Path, list[tuple[str, Path]]],
    keep: int,
) -> None:
    """Verify the retained run artifacts still exist after cleanup."""
    failures: list[str] = []
    for _location, entries in inventories.items():
        for _run_id, path in entries[:keep]:
            if not path.exists():
                failures.append(str(path))
    if failures:
        raise RuntimeError(
            "Retention verification failed; retained artifacts are missing:\n"
            + "\n".join(f"  {path}" for path in failures)
        )


def snapshot_protected_paths() -> tuple[Path, ...]:
    """Record only protected paths that existed before cleanup."""
    root = project_root()
    return tuple(
        (root / relative).resolve()
        for relative in PROTECTED_RELATIVE_PATHS
        if (root / relative).exists()
    )


def verify_protected_paths(expected_existing: Iterable[Path]) -> None:
    """Verify protected paths that existed before cleanup still exist."""
    missing = [str(path) for path in expected_existing if not path.exists()]
    if missing:
        raise RuntimeError(
            "Protected project paths are missing after cleanup:\n"
            + "\n".join(f"  {path}" for path in missing)
        )


def delete_artifacts(candidates: tuple[CleanupCandidate, ...]) -> None:
    """Delete approved old generated run directories."""
    for candidate in candidates:
        safe_remove_generated_run_directory(candidate)


def delete_empty_directories(paths: tuple[Path, ...]) -> None:
    for path in paths:
        # Re-check emptiness immediately before deletion. This prevents an
        # empty-directory candidate from deleting a directory that received
        # content after inspection.
        if not path.exists():
            continue
        if any(path.iterdir()):
            raise RuntimeError(
                f"Refusing to delete directory that is no longer empty: {path}"
            )
        print(f"Deleting empty directory: {path}")
        safe_remove_directory(path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Safely inspect and clean explicitly managed generated project "
            "artifacts."
        )
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Inspect and display the cleanup plan without deleting anything.",
    )
    parser.add_argument(
        "--keep",
        type=int,
        default=KEEP_RECENT_RUNS,
        metavar="N",
        help=(
            f"Number of recent runs to keep independently in each managed "
            f"location ({MIN_KEEP_RECENT_RUNS}-{MAX_KEEP_RECENT_RUNS}; default: "
            f"{KEEP_RECENT_RUNS})."
        ),
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        keep = validate_keep_count(args.keep)
    except ValueError as exc:
        print(f"ERROR: {exc}")
        return 2

    print("\nStarting controlled project maintenance...")
    if not env.authenticate():
        print("Authentication failed. Cleanup aborted.")
        return 1

    try:
        inventories = discover_all_managed_runs()
        display_inventory(inventories, keep)

        protected_snapshot = snapshot_protected_paths()

        plan = CleanupPlan(
            artifact_candidates=calculate_artifact_candidates(inventories, keep),
            empty_directories=discover_approved_empty_directories(),
        )
        display_plan(plan)

        if args.dry_run:
            print("\nDRY RUN: no files or directories were deleted.")
            return 0

        if not plan.artifact_candidates and not plan.empty_directories:
            print("\nNothing is eligible for cleanup.")
            return 0

        artifacts_approved = False
        empty_dirs_approved = False

        if plan.artifact_candidates:
            artifacts_approved = confirm(
                "Delete the complete old generated run directories listed above?"
            )

        if plan.empty_directories:
            empty_dirs_approved = confirm(
                "Delete the approved empty directories listed above?"
            )

        if not artifacts_approved and not empty_dirs_approved:
            print("\nNo cleanup action approved. Nothing was deleted.")
            return 0

        if artifacts_approved:
            delete_artifacts(plan.artifact_candidates)
            verify_artifact_deletions(plan.artifact_candidates)

        if empty_dirs_approved:
            delete_empty_directories(plan.empty_directories)
            verify_empty_directory_deletions(plan.empty_directories)

        # Retention and protected-area verification are based on the exact
        # pre-cleanup inventory, so the cleaner can prove that approved
        # retained artifacts and protected paths survived.
        verify_retained_runs(inventories, keep)
        verify_protected_paths(protected_snapshot)

        print("\nCleanup completed and verification passed.")
        return 0

    except KeyboardInterrupt:
        print("\nCleanup interrupted. No further deletion performed.")
        return 130
    except Exception as exc:  # pragma: no cover - defensive CLI boundary
        print(f"\nERROR: Cleanup aborted: {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
