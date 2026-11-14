
"""Read-only, dataset-independent access to EDI pipeline runs and data."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
RUNS_DIRECTORY = PROJECT_ROOT / "data" / "output" / "runs"

PIPELINE_STAGES = (
    "cleaning",
    "preprocessing",
    "analysis",
    "forecasting",
    "recommendation",
    "powerbi",
)

# Explicit allowlist: the dashboard cannot open arbitrary run artifacts.
ARTIFACTS: dict[str, tuple[str, ...]] = {
    "cleaning": (
        "processed/processed_dataset.json",
        "processed/processed_dataset.xlsx",
        "processed/cleaning_report.json",
    ),
    "preprocessing": (
        "output/analytical_dataset.json",
        "output/analytical_dataset.xlsx",
        "output/preprocessing_report.json",
    ),
    "analysis": (
        "output/analysis_results.json",
        "output/analysis_results.xlsx",
        "output/analysis_report.json",
    ),
    "forecasting": (
        "output/forecast_results.json",
        "output/forecast_results.xlsx",
        "output/forecast_report.json",
    ),
    "recommendation": (
        "output/recommendation_results.xlsx",
        "output/recommendation_report.json",
    ),
}

ALLOWED_ARTIFACTS = {
    path for paths in ARTIFACTS.values() for path in paths
}


def read_json(path: Path) -> Any:
    """Read JSON without assuming its root structure."""
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def sha256_file(path: Path) -> str:
    """Calculate a file's SHA-256 hash without loading it all into memory."""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _read_manifest(path: Path) -> dict[str, Any] | None:
    try:
        payload = read_json(path)
    except (OSError, json.JSONDecodeError):
        return None

    return payload if isinstance(payload, dict) else None


def _is_successful_run(
    run_directory: Path,
    manifest: dict[str, Any],
) -> bool:
    """Require matching identity, every stage passed, and success log marker."""
    if manifest.get("run_id") != run_directory.name:
        return False

    stages = manifest.get("stages")
    if not isinstance(stages, dict):
        return False

    for stage in PIPELINE_STAGES:
        state = stages.get(stage)
        if not isinstance(state, dict) or state.get("status") != "PASSED":
            return False

    log_path = run_directory / "execution.log"
    try:
        with log_path.open("rb") as handle:
            handle.seek(0, 2)
            size = handle.tell()
            handle.seek(max(0, size - 128_000))
            tail = handle.read().decode("utf-8", errors="replace")
    except OSError:
        return False

    return "PIPELINE: SUCCESS" in tail


def get_run_history() -> list[dict[str, Any]]:
    """Return readable run manifests, newest first, including unsuccessful runs."""
    if not RUNS_DIRECTORY.is_dir():
        return []

    history: list[dict[str, Any]] = []

    for directory in RUNS_DIRECTORY.iterdir():
        if not directory.is_dir():
            continue

        manifest_path = directory / "run_manifest.json"
        if not manifest_path.is_file():
            continue

        manifest = _read_manifest(manifest_path)
        if manifest is None or manifest.get("run_id") != directory.name:
            continue

        try:
            modified_at = manifest_path.stat().st_mtime
        except OSError:
            continue

        history.append(
            {
                "run_id": directory.name,
                "run_directory": directory,
                "manifest_path": manifest_path,
                "manifest": manifest,
                "modified_at": modified_at,
                "successful": _is_successful_run(directory, manifest),
            }
        )

    return sorted(history, key=lambda item: item["modified_at"], reverse=True)


def get_successful_runs() -> list[dict[str, Any]]:
    """Return verified successful runs, newest first."""
    return [run for run in get_run_history() if run["successful"]]


def get_latest_successful_run() -> dict[str, Any] | None:
    runs = get_successful_runs()
    return runs[0] if runs else None


def get_artifact_path(
    run: dict[str, Any],
    relative_path: str,
) -> Path | None:
    """Resolve only allowlisted artifacts inside the selected run."""
    if relative_path not in ALLOWED_ARTIFACTS:
        return None

    run_directory = Path(run["run_directory"]).resolve()
    candidate = (run_directory / relative_path).resolve()

    try:
        candidate.relative_to(run_directory)
    except ValueError:
        return None

    return candidate if candidate.is_file() else None


def get_artifact_catalog(run: dict[str, Any]) -> list[dict[str, Any]]:
    """Describe expected artifacts without assuming they all exist."""
    catalog = []

    for stage, relative_paths in ARTIFACTS.items():
        for relative_path in relative_paths:
            path = get_artifact_path(run, relative_path)
            catalog.append(
                {
                    "stage": stage,
                    "relative_path": relative_path,
                    "exists": path is not None,
                    "size_bytes": path.stat().st_size if path else None,
                }
            )

    return catalog


def load_run_tables(
    run: dict[str, Any],
    relative_path: str,
) -> dict[str, pd.DataFrame]:
    """Load a known dataset artifact using EDI's existing loader."""
    path = get_artifact_path(run, relative_path)
    if path is None:
        raise FileNotFoundError(
            f"Artifact is unavailable in run {run['run_id']}: {relative_path}"
        )

    from src.data_loader import load_file

    return load_file(path)


def load_original_dataset(run: dict[str, Any]) -> dict[str, pd.DataFrame]:
    """Load the original input only when it still matches the recorded hash."""
    manifest = run["manifest"]
    raw_path = manifest.get("dataset_path")
    expected_hash = manifest.get("dataset_sha256")

    if not isinstance(raw_path, str) or not raw_path.strip():
        raise FileNotFoundError("The run manifest does not identify an input dataset.")

    if not isinstance(expected_hash, str) or not expected_hash.strip():
        raise ValueError("The run manifest has no input dataset hash.")

    path = Path(raw_path).expanduser().resolve()

    if not path.is_file():
        raise FileNotFoundError("The original input dataset is no longer available.")

    # Reuse EDI's supported-format and file-size validation.
    from src.data_loader import load_file, validate_file

    path = validate_file(path)

    if sha256_file(path).lower() != expected_hash.lower():
        raise ValueError(
            "The input file no longer matches this run's recorded dataset hash."
        )

    return load_file(path)


def load_run_json(
    run: dict[str, Any],
    relative_path: str,
) -> Any:
    if not relative_path.endswith(".json"):
        raise ValueError("Expected a JSON artifact path.")

    path = get_artifact_path(run, relative_path)
    if path is None:
        raise FileNotFoundError(
            f"Artifact is unavailable in run {run['run_id']}: {relative_path}"
        )

    return read_json(path)


def load_run_workbook(
    run: dict[str, Any],
    relative_path: str,
) -> dict[str, pd.DataFrame]:
    if not relative_path.endswith((".xlsx", ".xls")):
        raise ValueError("Expected an Excel artifact path.")

    path = get_artifact_path(run, relative_path)
    if path is None:
        raise FileNotFoundError(
            f"Artifact is unavailable in run {run['run_id']}: {relative_path}"
        )

    return pd.read_excel(path, sheet_name=None)
