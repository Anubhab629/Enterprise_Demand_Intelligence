"""Controlled project maintenance utility.

This module is intentionally separate from the application pipeline.
It performs only explicitly-scoped cleanup of generated runtime artifacts
and approved empty project directories.

It is NOT imported by ``main.py`` and does not modify application source
code, datasets, Power BI integration logic, or the environment controller.
"""

from __future__ import annotations

import argparse
import re
import shutil
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
    Path("powerbi/documentation"),
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
    """Discover only recognized generated run directories in one location."""
    root = absolute_path(location)
    if not root.exists():
        return []
    if not root.is_dir():
        raise RuntimeError(f"Managed cleanup location is not a directory: {root}")
    if is_protected(root):
        raise RuntimeError(f"Managed cleanup location is protected: {root}")

    discovered: list[tuple[str, Path]] = []
    for child in root.iterdir():
        if not child.is_dir():
            continue

        run_id = run_id_from_name(child.name)
        if run_id is None:
            continue

        if is_protected(child):
            raise RuntimeError(f"Generated candidate resolves into protected path: {child}")

        discovered.append((run_id, child))

    # Run IDs begin with a UTC timestamp, so descending lexical order is
    # deterministic and newest-first for recognized IDs.
    discovered.sort(key=lambda item: item[0], reverse=True)
    return discovered


def discover_all_managed_runs() -> dict[Path, list[tuple[str, Path]]]:
    """Inspect every managed location independently."""
    results: dict[Path, list[tuple[str, Path]]] = {}

    for location in MANAGED_RUN_LOCATIONS:
        if location == Path("powerbi"):
            # Only report projects matching the explicit Power BI run-project
            # naming contract. Other Power BI content is never a cleanup target.
            root = absolute_path(location)
            if not root.exists():
                results[location] = []
                continue
            discovered: list[tuple[str, Path]] = []
            for child in root.iterdir():
                if not child.is_dir():
                    continue
                run_id = run_id_from_name(child.name)
                if run_id is None:
                    continue
                if not child.name.startswith("Demand_Intelligence_Report_run_"):
                    continue
                if is_protected(child):
                    raise RuntimeError(
                        f"Generated Power BI candidate resolves into protected path: {child}"
                    )
                discovered.append((run_id, child))
            discovered.sort(key=lambda item: item[0], reverse=True)
            results[location] = discovered
        else:
            results[location] = discover_run_directories(location)

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
    """Select only artifacts older than the independent retention boundary."""
    candidates: list[CleanupCandidate] = []

    for location, entries in inventories.items():
        for run_id, path in entries[keep:]:
            candidates.append(
                CleanupCandidate(location=location, path=path, run_id=run_id)
            )

    return tuple(candidates)


def discover_approved_empty_directories() -> tuple[Path, ...]:
    """Return approved project directories that are currently empty."""
    empty: list[Path] = []
    for relative in APPROVED_EMPTY_DIRECTORIES:
        path = absolute_path(relative)
        if not path.exists():
            continue
        if not path.is_dir():
            raise RuntimeError(f"Approved empty-directory target is not a directory: {path}")
        if is_protected(path):
            raise RuntimeError(f"Approved empty-directory target is protected: {path}")
        if not any(path.iterdir()):
            empty.append(path)
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
            marker = "KEEP" if index <= keep else "DELETE CANDIDATE"
            print(f"  {marker:<16} {path}")


def display_plan(plan: CleanupPlan) -> None:
    print("\n" + "=" * 78)
    print("CLEANUP PLAN")
    print("=" * 78)

    print("\nGenerated artifact candidates:")
    if plan.artifact_candidates:
        for candidate in plan.artifact_candidates:
            print(f"  DELETE  {candidate.path}")
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
    for candidate in candidates:
        print(f"Deleting: {candidate.path}")
        safe_remove_directory(candidate.path)


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
                "Delete the generated artifact candidates listed above?"
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
