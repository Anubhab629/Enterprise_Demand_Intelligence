"""
Enterprise Demand Intelligence Platform
GENERIC DATA CLEANING AND STANDARDIZATION PIPELINE

Purpose
-------
Create a trustworthy processed representation of the active dataset without
modifying the raw source.

Design
------
The cleaning layer does not contain a business schema. It consumes structural
evidence discovered by data_loader.py and data_profiler.py when that evidence
is available. It never requires particular table names, column names, entity
counts, business values, targets, or domain concepts.

Safe transformations
--------------------
- trim surrounding whitespace in textual values
- convert truly empty textual values to missing
- preserve genuine missing values
- preserve identifiers and identifier-like values as text when structural
  evidence identifies them as identifiers
- standardize columns already identified as temporal by structural evidence
- standardize columns already identified as numeric or boolean by structural
  evidence
- remove only exact duplicate rows
- record every transformation
- never overwrite the raw source
- fail rather than silently destroy information

Important
---------
Cleaning is intentionally conservative. It does not invent values, impute
missing values, remove statistical outliers, infer a demand target, apply
domain-specific rules, or rename business fields.
"""

from __future__ import annotations

import json
import math
import os
import sys
from pathlib import Path
from typing import Any, Mapping

import pandas as pd


# ============================================================
# PROJECT PATH
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


# ============================================================
# CONFIGURATION
# ============================================================

import config


PROCESSED_DIRECTORY = (
    PROJECT_ROOT / "data" / "processed"
)

PROCESSED_DIRECTORY.mkdir(
    parents=True,
    exist_ok=True,
)

PROCESSED_JSON_PATH = (
    PROCESSED_DIRECTORY / "processed_dataset.json"
)

PROCESSED_EXCEL_PATH = (
    PROCESSED_DIRECTORY / "processed_dataset.xlsx"
)

CLEANING_REPORT_PATH = (
    PROCESSED_DIRECTORY / "cleaning_report.json"
)


# ============================================================
# GENERIC VALUE HELPERS
# ============================================================

def _is_missing(value: Any) -> bool:
    """Return True only when a scalar value is genuinely missing."""

    if value is None:
        return True

    try:
        result = pd.isna(value)
    except Exception:
        return False

    return bool(result) if isinstance(result, bool) else False


def _normalize_text_value(value: Any) -> Any:
    """
    Trim surrounding whitespace from text.

    A string that becomes empty is converted to pandas missing value.
    Other values are preserved unchanged.
    """

    if isinstance(value, str):
        cleaned = value.strip()
        return pd.NA if cleaned == "" else cleaned

    return value


def _canonical_scalar(value: Any) -> Any:
    """
    Produce a stable scalar representation for structural comparisons.
    """

    if _is_missing(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.isoformat()

    if isinstance(value, (float,)):
        if not math.isfinite(value):
            return str(value)

    try:
        if hasattr(value, "item"):
            value = value.item()
    except Exception:
        pass

    if isinstance(value, (str, int, float, bool)):
        return value

    return str(value)


def _safe_json_value(value: Any) -> Any:
    """Convert pandas/numpy/date values to JSON-safe values."""

    if isinstance(value, dict):
        return {
            str(key): _safe_json_value(item)
            for key, item in value.items()
        }

    if isinstance(value, (list, tuple, set)):
        return [
            _safe_json_value(item)
            for item in value
        ]

    if value is None:
        return None

    if _is_missing(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.isoformat()

    try:
        if hasattr(value, "item"):
            return _safe_json_value(value.item())
    except Exception:
        pass

    if isinstance(value, float):
        if not math.isfinite(value):
            return None

    if isinstance(value, (str, int, float, bool)):
        return value

    return str(value)


# ============================================================
# DISCOVERY RESULT HANDLING
# ============================================================

def _get_discovery_result() -> Mapping[str, Any]:
    """
    Load the active dataset exclusively through data_loader.py.

    The loader remains responsible for supported formats, file-size limits,
    dataset-path resolution, table discovery, and structural inference.
    """

    from src.data_loader import discover_dataset

    result = discover_dataset()

    if not isinstance(result, Mapping):
        raise TypeError(
            "data_loader.discover_dataset() must return a mapping."
        )

    return result


def _extract_tables(
    discovery: Mapping[str, Any],
) -> dict[str, pd.DataFrame]:
    """
    Extract discovered tables from the loader result.

    The normal loader contract exposes a 'tables' mapping. Values may already
    be DataFrames or may be record-like structures.
    """

    raw_tables = discovery.get("tables")

    if not isinstance(raw_tables, Mapping):
        raise RuntimeError(
            "The dataset loader did not expose a discovered 'tables' mapping."
        )

    tables: dict[str, pd.DataFrame] = {}

    for table_name, table_value in raw_tables.items():
        name = str(table_name)

        if isinstance(table_value, pd.DataFrame):
            tables[name] = table_value.copy()
            continue

        if isinstance(table_value, Mapping):
            records = table_value.get("records")

            if isinstance(records, list):
                tables[name] = pd.DataFrame(records)
                continue

        if isinstance(table_value, list):
            tables[name] = pd.DataFrame(table_value)
            continue

        raise TypeError(
            f"Discovered table {name!r} is not a supported table representation."
        )

    if not tables:
        raise RuntimeError(
            "No non-empty table mapping was returned by data_loader.py."
        )

    return tables


def _extract_dataset_path(
    discovery: Mapping[str, Any],
) -> str | None:
    """Extract a dataset path when the loader reports one."""

    dataset_info = discovery.get("dataset")

    if isinstance(dataset_info, Mapping):
        for key in ("path", "dataset_path", "source_path"):
            value = dataset_info.get(key)
            if value:
                return str(value)

    for key in ("path", "dataset_path", "source_path"):
        value = discovery.get(key)
        if value:
            return str(value)

    active = getattr(config, "ACTIVE_DATASET_PATH", None)

    if active:
        return str(active)

    return None


# ============================================================
# STRUCTURAL EVIDENCE
# ============================================================

def _extract_profiles(
    discovery: Mapping[str, Any],
) -> Mapping[str, Any]:
    profiles = discovery.get("profiles")
    return profiles if isinstance(profiles, Mapping) else {}


def _extract_table_profile(
    profiles: Mapping[str, Any],
    table_name: str,
) -> Mapping[str, Any]:
    value = profiles.get(table_name)
    return value if isinstance(value, Mapping) else {}


def _profile_columns(
    table_profile: Mapping[str, Any],
) -> list[Mapping[str, Any]]:
    columns = table_profile.get("columns")

    if isinstance(columns, list):
        return [
            column
            for column in columns
            if isinstance(column, Mapping)
        ]

    if isinstance(columns, Mapping):
        return [
            {
                "name": column_name,
                **(
                    column_profile
                    if isinstance(column_profile, Mapping)
                    else {}
                ),
            }
            for column_name, column_profile in columns.items()
        ]

    return []


def _column_profile(
    table_profile: Mapping[str, Any],
    column_name: Any,
) -> Mapping[str, Any]:
    columns = _profile_columns(table_profile)

    for profile in columns:
        if profile.get("name") == column_name:
            return profile

        if str(profile.get("name")) == str(column_name):
            return profile

    return {}


def _evidence_tokens(
    column_profile: Mapping[str, Any],
) -> set[str]:
    """
    Collect structural type labels from discovered metadata.

    This function intentionally does not inspect column names.
    """

    tokens: set[str] = set()

    for key in (
        "inferred_type",
        "type",
        "data_type",
        "structural_type",
        "role",
    ):
        value = column_profile.get(key)

        if isinstance(value, str):
            tokens.add(value.strip().lower())

    for key in (
        "is_identifier",
        "identifier_candidate",
        "appears_identifier",
    ):
        value = column_profile.get(key)

        if value is True:
            tokens.add("identifier")

    return tokens


def _column_kind(
    table_profile: Mapping[str, Any],
    column_name: Any,
    series: pd.Series,
) -> str:
    """
    Resolve a conservative structural kind from profiler evidence.

    Existing pandas dtype is used only as a fallback. Column names are never
    inspected to determine semantic type.
    """

    profile = _column_profile(
        table_profile,
        column_name,
    )

    tokens = _evidence_tokens(profile)

    identifier_tokens = {
        "identifier",
        "id",
        "identifier_candidate",
        "key",
    }

    temporal_tokens = {
        "temporal",
        "datetime",
        "date",
        "timestamp",
    }

    numeric_tokens = {
        "numeric",
        "number",
        "float",
        "integer",
        "int",
        "decimal",
    }

    boolean_tokens = {
        "boolean",
        "bool",
    }

    if tokens & identifier_tokens:
        return "identifier"

    if tokens & temporal_tokens:
        return "temporal"

    if tokens & boolean_tokens:
        return "boolean"

    if tokens & numeric_tokens:
        return "numeric"

    # Structural dtype fallback. This does not use business vocabulary.
    if pd.api.types.is_bool_dtype(series.dtype):
        return "boolean"

    if pd.api.types.is_numeric_dtype(series.dtype):
        return "numeric"

    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        return "temporal"

    if isinstance(series.dtype, pd.DatetimeTZDtype):
        return "temporal"

    return "other"


# ============================================================
# SAFE STRING NORMALIZATION
# ============================================================

def _normalize_text_columns(
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    """Normalize whitespace without changing non-text values."""

    df = dataframe.copy()
    changes = 0

    for column in df.columns:
        series = df[column]

        if not (
            pd.api.types.is_object_dtype(series)
            or pd.api.types.is_string_dtype(series)
        ):
            continue

        before = series.copy()
        after = series.map(_normalize_text_value)

        changed = (
            before.astype("string").fillna("<NA>")
            != after.astype("string").fillna("<NA>")
        )

        changes += int(changed.sum())
        df[column] = after

    return df, changes


# ============================================================
# STRUCTURAL TYPE STANDARDIZATION
# ============================================================

def _normalize_identifier(
    series: pd.Series,
) -> tuple[pd.Series, int]:
    """
    Preserve identifier content while using a stable text dtype.

    No identifier vocabulary is used.
    """

    before = series.copy()

    def convert(value: Any) -> Any:
        if _is_missing(value):
            return pd.NA

        if isinstance(value, str):
            return value.strip() if value.strip() else pd.NA

        return str(value)

    after = series.map(convert).astype("string")

    changed = (
        before.astype("string").fillna("<NA>")
        != after.astype("string").fillna("<NA>")
    )

    return after, int(changed.sum())


def _normalize_temporal(
    series: pd.Series,
    table_name: str,
    column_name: str,
) -> tuple[pd.Series, int]:
    """
    Convert a structurally identified temporal column.

    Conversion is transactional: if any non-null source value cannot be
    converted, the function raises instead of silently losing that value.
    """

    source = series.copy()

    if pd.api.types.is_datetime64_any_dtype(source.dtype):
        return source, 0

    if isinstance(source.dtype, pd.DatetimeTZDtype):
        return source, 0

    normalized = source.map(_normalize_text_value)

    source_non_null = int(normalized.notna().sum())

    if source_non_null == 0:
        return (
            pd.to_datetime(normalized, errors="coerce"),
            0,
        )

    converted = pd.to_datetime(
        normalized,
        errors="coerce",
        format="mixed",
    )

    converted_non_null = int(converted.notna().sum())

    if converted_non_null != source_non_null:
        failed = (
            normalized.loc[
                normalized.notna() & converted.isna()
            ]
            .astype(str)
            .head(10)
            .tolist()
        )

        raise ValueError(
            "Temporal standardization would destroy source values in "
            f"{table_name}.{column_name}. "
            f"Source non-null values={source_non_null:,}; "
            f"converted={converted_non_null:,}; "
            f"examples={failed}"
        )

    changed = (
        source.astype("string").fillna("<NA>")
        != converted.astype("string").fillna("<NA>")
    )

    return converted, int(changed.sum())


def _normalize_numeric(
    series: pd.Series,
    table_name: str,
    column_name: str,
) -> tuple[pd.Series, int]:
    """
    Standardize a structurally numeric column.

    Existing numeric values are preserved. Text values are converted only
    when every non-null source value converts successfully.
    """

    source = series.copy()

    if pd.api.types.is_numeric_dtype(source.dtype):
        return source, 0

    non_null = source.dropna()

    if non_null.empty:
        return source, 0

    converted_non_null = pd.to_numeric(
        non_null,
        errors="coerce",
    )

    if not converted_non_null.notna().all():
        failed = (
            non_null.loc[converted_non_null.isna()]
            .astype(str)
            .head(10)
            .tolist()
        )

        raise ValueError(
            "Numeric standardization would destroy source values in "
            f"{table_name}.{column_name}. "
            f"Examples={failed}"
        )

    converted = pd.to_numeric(
        source,
        errors="coerce",
    )

    if int(converted.notna().sum()) != int(source.notna().sum()):
        raise ValueError(
            "Numeric standardization changed the number of non-null values in "
            f"{table_name}.{column_name}."
        )

    changed = (
        source.astype("string").fillna("<NA>")
        != converted.astype("string").fillna("<NA>")
    )

    return converted, int(changed.sum())


def _normalize_boolean(
    series: pd.Series,
    table_name: str,
    column_name: str,
) -> tuple[pd.Series, int]:
    """
    Standardize a column only when the profiler has already identified it as
    boolean.

    Unknown values are never converted to NULL.
    """

    source = series.copy()

    if pd.api.types.is_bool_dtype(source.dtype):
        return source, 0

    true_values = {True, "true", "True", "TRUE", 1}
    false_values = {False, "false", "False", "FALSE", 0}

    result: list[Any] = []

    for value in source:
        if _is_missing(value):
            result.append(pd.NA)
        elif value in true_values:
            result.append(True)
        elif value in false_values:
            result.append(False)
        else:
            raise ValueError(
                "Boolean standardization encountered an unrecognized non-null "
                f"value in {table_name}.{column_name}: {value!r}"
            )

    converted = pd.Series(
        result,
        index=source.index,
        dtype="boolean",
    )

    changed = (
        source.astype("string").fillna("<NA>")
        != converted.astype("string").fillna("<NA>")
    )

    return converted, int(changed.sum())


# ============================================================
# EXACT DUPLICATES
# ============================================================

def _remove_exact_duplicates(
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    """Remove only completely identical rows."""

    duplicate_count = int(
        dataframe.duplicated(keep="first").sum()
    )

    if duplicate_count == 0:
        return dataframe.reset_index(drop=True), 0

    return (
        dataframe.drop_duplicates(
            keep="first"
        ).reset_index(drop=True),
        duplicate_count,
    )


# ============================================================
# TABLE CLEANING
# ============================================================

def _clean_table(
    table_name: str,
    dataframe: pd.DataFrame,
    table_profile: Mapping[str, Any],
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Clean one table using only structural evidence."""

    df = dataframe.copy()

    original_rows = len(df)
    original_columns = list(df.columns)

    # Column labels are part of the source schema and are preserved exactly.
    # Only cell values may be normalized by the cleaning rules below.
    if len(set(original_columns)) != len(original_columns):
        raise ValueError(
            f"Duplicate column labels detected in "
            f"{table_name!r}."
        )

    df, text_changes = _normalize_text_columns(df)

    type_changes = {
        "identifier": 0,
        "temporal": 0,
        "numeric": 0,
        "boolean": 0,
    }

    converted_columns: dict[Any, str] = {}

    for column in df.columns:
        kind = _column_kind(
            table_profile,
            column,
            df[column],
        )

        if kind == "identifier":
            df[column], count = _normalize_identifier(df[column])
            type_changes["identifier"] += count
            converted_columns[column] = "identifier"

        elif kind == "temporal":
            df[column], count = _normalize_temporal(
                df[column],
                table_name,
                str(column),
            )
            type_changes["temporal"] += count
            converted_columns[column] = "temporal"

        elif kind == "numeric":
            df[column], count = _normalize_numeric(
                df[column],
                table_name,
                str(column),
            )
            type_changes["numeric"] += count
            converted_columns[column] = "numeric"

        elif kind == "boolean":
            df[column], count = _normalize_boolean(
                df[column],
                table_name,
                str(column),
            )
            type_changes["boolean"] += count
            converted_columns[column] = "boolean"

    df, duplicates_removed = _remove_exact_duplicates(df)

    return df, {
        "table": table_name,
        "original_rows": original_rows,
        "final_rows": len(df),
        "columns": len(df.columns),
        "text_value_changes": text_changes,
        "type_standardization_changes": type_changes,
        "converted_columns": converted_columns,
        "exact_duplicates_removed": duplicates_removed,
        "null_values_after_cleaning": int(df.isna().sum().sum()),
    }


# ============================================================
# STRUCTURAL SAFETY CHECKS
# ============================================================

def _check_data_integrity(
    original_tables: Mapping[str, pd.DataFrame],
    cleaned_tables: Mapping[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Verify that cleaning did not silently alter the basic structure.

    Row counts may decrease only because exact duplicates were removed.
    Column counts, column order, and column labels must remain stable.
    """

    findings: list[dict[str, Any]] = []

    for table_name, original in original_tables.items():
        cleaned = cleaned_tables[table_name]

        if list(original.columns) != list(
            cleaned.columns
        ):
            findings.append({
                "severity": "ERROR",
                "table": table_name,
                "issue": "Column schema changed unexpectedly.",
            })

        if len(cleaned) > len(original):
            findings.append({
                "severity": "ERROR",
                "table": table_name,
                "issue": "Cleaning increased row count.",
            })

    return findings


def _check_finite_numeric_values(
    tables: Mapping[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Check only a technical property: numeric values must be finite.

    No business-specific lower/upper bounds are applied.
    """

    findings: list[dict[str, Any]] = []

    for table_name, df in tables.items():
        for column in df.columns:
            series = df[column]

            if not pd.api.types.is_numeric_dtype(series.dtype):
                continue

            numeric = pd.to_numeric(
                series,
                errors="coerce",
            )

            invalid_mask = (
                series.notna()
                & numeric.isna()
            )

            if invalid_mask.any():
                findings.append({
                    "severity": "ERROR",
                    "table": table_name,
                    "column": column,
                    "issue": "Non-convertible numeric values detected.",
                    "count": int(invalid_mask.sum()),
                })

                continue

            try:
                finite_mask = (
                    numeric.notna()
                    & ~numeric.map(math.isfinite)
                )
            except TypeError:
                continue

            if finite_mask.any():
                findings.append({
                    "severity": "ERROR",
                    "table": table_name,
                    "column": str(column),
                    "issue": "Non-finite numeric values detected.",
                    "count": int(finite_mask.sum()),
                })

    return findings


# ============================================================
# OPTIONAL VALIDATION INTEGRATION
# ============================================================

def _run_available_validation(
    tables: dict[str, pd.DataFrame],
    profiler_report: Mapping[str, Any] | None = None,
    relationship_candidates: Any = None,
) -> dict[str, Any]:
    """
    Run the project's explicit validation contract against the processed
    tables.

    Validation remains read-only and uses structural evidence supplied by
    the earlier discovery stage.
    """

    try:
        from src.validation import validate_dataset
    except ImportError as exc:
        return {
            "status": "NOT_RUN",
            "reason": (
                "Could not import validation.validate_dataset: "
                f"{exc}"
            ),
        }

    try:
        result = validate_dataset(
            tables=tables,
            primary_key_map=None,
            foreign_key_map=None,
            profiler_report=profiler_report,
            relationship_candidates=relationship_candidates,
        )

        if not isinstance(result, Mapping):
            return {
                "status": "COMPLETED",
                "result": _safe_json_value(result),
            }

        status = str(
            result.get("status", "")
        ).upper()

        if status == "INVALID":
            return {
                "status": "FAILED",
                "result": result,
                "reason": (
                    "The processed dataset failed validation."
                ),
            }

        return {
            "status": "PASSED",
            "result": result,
        }

    except Exception as exc:
        return {
            "status": "FAILED",
            "reason": str(exc),
        }


def _validation_result_is_valid(
    result: Mapping[str, Any],
) -> bool:
    """
    Interpret common validation report status shapes without assuming
    particular business sections.
    """

    status = str(
        result.get("overall_status", result.get("status", ""))
    ).upper()

    if status in {
        "INVALID",
        "FAILED",
        "FAIL",
    }:
        return False

    errors = result.get("errors")

    if isinstance(errors, int):
        return errors == 0

    if isinstance(errors, list):
        return len(errors) == 0

    return True


# ============================================================
# OUTPUT WRITERS
# ============================================================

def _write_json(
    tables: Mapping[str, pd.DataFrame],
    path: Path,
) -> None:
    payload: dict[str, list[dict[str, Any]]] = {}

    for table_name, dataframe in tables.items():
        records = dataframe.to_dict(
            orient="records"
        )

        payload[str(table_name)] = [
            {
                str(key): _safe_json_value(value)
                for key, value in record.items()
            }
            for record in records
        ]

    temporary = path.with_suffix(
        path.suffix + ".tmp"
    )

    with temporary.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            payload,
            handle,
            ensure_ascii=False,
            indent=2,
        )

    temporary.replace(path)


def _safe_sheet_name(
    table_name: str,
    used: set[str],
) -> str:
    """
    Produce a valid Excel sheet label without using business vocabulary.

    Excel limits sheet labels to 31 characters.
    """

    base = (
        str(table_name)
        .replace("/", "_")
        .replace("\\", "_")
        .replace("?", "_")
        .replace("*", "_")
        .replace("[", "_")
        .replace("]", "_")
        .replace(":", "_")
        .strip()
    )

    if not base:
        base = "Table"

    base = base[:31]

    candidate = base
    counter = 1

    while candidate in used:
        suffix = f"_{counter}"
        candidate = (
            base[: 31 - len(suffix)]
            + suffix
        )
        counter += 1

    used.add(candidate)
    return candidate


def _write_excel(
    tables: Mapping[str, pd.DataFrame],
    path: Path,
) -> None:
    temporary = path.with_suffix(
        path.suffix + ".tmp.xlsx"
    )

    used_sheets: set[str] = set()

    with pd.ExcelWriter(
        temporary,
        engine="openpyxl",
    ) as writer:
        for table_name, dataframe in tables.items():
            dataframe.to_excel(
                writer,
                sheet_name=_safe_sheet_name(
                    str(table_name),
                    used_sheets,
                ),
                index=False,
            )

    temporary.replace(path)


# ============================================================
# CLEANING REPORT
# ============================================================

def _build_report(
    source_path: str | None,
    statistics: list[dict[str, Any]],
    integrity_findings: list[dict[str, Any]],
    technical_findings: list[dict[str, Any]],
    validation_result: Mapping[str, Any],
) -> dict[str, Any]:
    return {
        "status": (
            "FAILED"
            if any(
                finding.get("severity") == "ERROR"
                for finding in (
                    integrity_findings
                    + technical_findings
                )
            )
            or validation_result.get("status") == "FAILED"
            else "COMPLETED"
        ),
        "source": source_path,
        "tables_processed": len(statistics),
        "total_original_rows": sum(
            item["original_rows"]
            for item in statistics
        ),
        "total_final_rows": sum(
            item["final_rows"]
            for item in statistics
        ),
        "total_exact_duplicates_removed": sum(
            item["exact_duplicates_removed"]
            for item in statistics
        ),
        "tables": statistics,
        "integrity_findings": integrity_findings,
        "technical_findings": technical_findings,
        "validation": validation_result,
        "transformations": {
            "string_whitespace_trimmed": True,
            "empty_text_to_missing": True,
            "temporal_standardization": True,
            "numeric_standardization": True,
            "boolean_standardization": True,
            "identifier_preservation": True,
            "exact_duplicate_removal_only": True,
            "missing_value_imputation": False,
            "outlier_removal": False,
            "business_rule_mutation": False,
            "target_inference": False,
            "business_column_renaming": False,
        },
    }


def _write_report(
    report: Mapping[str, Any],
    path: Path,
) -> None:
    temporary = path.with_suffix(
        path.suffix + ".tmp"
    )

    with temporary.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            _safe_json_value(dict(report)),
            handle,
            ensure_ascii=False,
            indent=2,
        )

    temporary.replace(path)


# ============================================================
# MAIN PIPELINE
# ============================================================

def run_cleaning_pipeline() -> dict[str, Any]:
    print("=" * 72)
    print("ENTERPRISE DEMAND INTELLIGENCE PLATFORM")
    print("GENERIC DATA CLEANING AND STANDARDIZATION")
    print("=" * 72)

    print()
    print("Loading dataset through data_loader.py...")

    discovery = _get_discovery_result()
    raw_tables = _extract_tables(discovery)

    source_path = _extract_dataset_path(discovery)
    profiles = _extract_profiles(discovery)

    print(
        f"Dataset loading: PASSED ({len(raw_tables)} tables)"
    )

    if source_path:
        print(f"Dataset source: {source_path}")

    print()
    print("Running conservative cleaning...")

    processed_tables: dict[str, pd.DataFrame] = {}
    statistics: list[dict[str, Any]] = []

    for table_name, dataframe in raw_tables.items():
        print()
        print(f"Cleaning: {table_name}")

        table_profile = _extract_table_profile(
            profiles,
            table_name,
        )

        cleaned, stats = _clean_table(
            table_name,
            dataframe,
            table_profile,
        )

        processed_tables[table_name] = cleaned
        statistics.append(stats)

        print(
            f"  Rows: {stats['original_rows']:,} "
            f"-> {stats['final_rows']:,}"
        )
        print(
            f"  Exact duplicates removed: "
            f"{stats['exact_duplicates_removed']:,}"
        )
        print(
            f"  NULL values after cleaning: "
            f"{stats['null_values_after_cleaning']:,}"
        )

    print()
    print("Checking structural integrity...")

    integrity_findings = _check_data_integrity(
        raw_tables,
        processed_tables,
    )

    technical_findings = _check_finite_numeric_values(
        processed_tables
    )

    if integrity_findings or technical_findings:
        print()
        print("Cleaning safety checks: FAILED.")

        for finding in (
            integrity_findings
            + technical_findings
        ):
            print(
                f"  [{finding.get('severity')}] "
                f"{finding.get('table')}"
                f"{'.' + finding.get('column') if finding.get('column') else ''}: "
                f"{finding.get('issue')}"
            )

        raise RuntimeError(
            "Cleaning safety checks failed. "
            "Processed output will not be written."
        )

    print("Cleaning safety checks: PASSED.")

    print()
    print("Running available dataset validation...")

    validation_result = _run_available_validation(
        processed_tables,
        profiler_report=discovery,
        relationship_candidates=discovery.get(
            "relationships",
            [],
        ),
    )

    print(
        f"Validation integration status: "
        f"{validation_result.get('status')}"
    )

    if validation_result.get("status") == "FAILED":
        raise RuntimeError(
            "The processed dataset failed the available validation layer: "
            + str(validation_result.get("reason", "validation failed"))
        )

    report = _build_report(
        source_path,
        statistics,
        integrity_findings,
        technical_findings,
        validation_result,
    )

    print()
    print("Writing processed JSON...")

    _write_json(
        processed_tables,
        PROCESSED_JSON_PATH,
    )

    print(
        f"Processed JSON: {PROCESSED_JSON_PATH}"
    )

    print()
    print("Writing processed Excel...")

    _write_excel(
        processed_tables,
        PROCESSED_EXCEL_PATH,
    )

    print(
        f"Processed Excel: {PROCESSED_EXCEL_PATH}"
    )

    _write_report(
        report,
        CLEANING_REPORT_PATH,
    )

    print()
    print("=" * 72)
    print("DATA CLEANING COMPLETED SUCCESSFULLY")
    print("=" * 72)

    print()
    print(
        f"Tables processed: {len(statistics)}"
    )
    print(
        f"Original rows: "
        f"{report['total_original_rows']:,}"
    )
    print(
        f"Final rows: "
        f"{report['total_final_rows']:,}"
    )
    print(
        f"Exact duplicates removed: "
        f"{report['total_exact_duplicates_removed']:,}"
    )

    print()
    print("Raw data protection:")
    print("  Raw source was not modified.")
    print("  No missing values were fabricated.")
    print("  No business values were invented.")
    print("  No outliers were removed.")
    print("  No target was inferred.")
    print("  No business-specific column names were required.")

    print()
    print("Cleaning report:")
    print(f"  {CLEANING_REPORT_PATH}")

    return report


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    try:
        run_cleaning_pipeline()

    except Exception as error:
        print()
        print("=" * 72)
        print("DATA CLEANING PIPELINE FAILED")
        print("=" * 72)
        print()
        print(f"ERROR: {error}")
        print()
        print("The raw dataset was not modified.")
        raise
