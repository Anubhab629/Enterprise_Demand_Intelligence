"""
Enterprise Demand Intelligence Platform
GENERIC DATA VALIDATION LAYER

Purpose
-------
Validate a dynamically discovered dataset without requiring a
demo/reference schema and without modifying input DataFrames.

Design principles
-----------------
- Dataset-driven rather than reference-dataset-driven.
- No business-specific table, column, entity, or target assumptions.
- Uses structural properties and optional discovered metadata.
- Distinguishes errors from warnings and informational findings.
- Never modifies input DataFrames.
- Supports direct execution through the VS Code Run button.
- Supports optional explicit key/relationship policies when supplied
  externally, without requiring them for generic validation.

Validation areas
----------------
1. Dataset/table structure
2. Missing and blank values
3. Structural key evidence
4. Explicit primary-key validation
5. Explicit/discovered relationship evidence
6. Numeric finiteness
7. Temporal validity
8. Duplicate rows
9. Validation summary and human-readable reporting
"""

from __future__ import annotations

from pathlib import Path
import json
import math
import sys
from typing import Any, Mapping, Sequence

import numpy as np
import pandas as pd


# ============================================================
# PROJECT ROOT
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


# ============================================================
# GENERIC CONFIGURATION
# ============================================================

try:
    import config
except ImportError as exc:
    raise ImportError(
        "Unable to import project config.py. "
        f"Original error: {exc}"
    ) from exc


REQUIRE_NON_NULL_VALUES = bool(
    getattr(config, "REQUIRE_NON_NULL_VALUES", False)
)

REQUIRE_NON_BLANK_VALUES = bool(
    getattr(config, "REQUIRE_NON_BLANK_VALUES", False)
)

REQUIRE_UNIQUE_PRIMARY_KEYS = bool(
    getattr(config, "REQUIRE_UNIQUE_PRIMARY_KEYS", False)
)

REQUIRE_VALID_FOREIGN_KEYS = bool(
    getattr(config, "REQUIRE_VALID_FOREIGN_KEYS", False)
)

REQUIRE_FINITE_NUMERIC_VALUES = bool(
    getattr(config, "REQUIRE_FINITE_NUMERIC_VALUES", True)
)

MAX_INVALID_VALUES_REPORTED = int(
    getattr(config, "MAX_INVALID_VALUES_REPORTED", 100)
)

VALIDATION_SAMPLE_SIZE = int(
    getattr(config, "SAMPLE_SIZE", 5000)
)


# ============================================================
# TYPE DEFINITIONS
# ============================================================

PrimaryKeyDefinition = Any

ForeignKeyRelationship = tuple[Any, Any]


# ============================================================
# GENERAL HELPERS
# ============================================================

def _safe_int(value: Any, default: int = 0) -> int:
    """Convert a value to int without allowing validation itself to fail."""
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _is_missing_scalar(value: Any) -> bool:
    """Return True for scalar missing/null values."""
    if value is None:
        return True

    try:
        result = pd.isna(value)

        if isinstance(result, (bool, np.bool_)):
            return bool(result)
    except (TypeError, ValueError):
        pass

    return False


def _is_blank_value(value: Any) -> bool:
    """
    Return True only for blank textual values.

    Missing/null values are intentionally handled separately.
    """
    if _is_missing_scalar(value):
        return False

    if isinstance(value, str):
        return value.strip() == ""

    return False


def _canonical_scalar(value: Any) -> Any:
    """
    Create a temporary canonical representation for relationship
    comparisons without modifying source data.
    """
    if _is_missing_scalar(value):
        return None

    if isinstance(value, str):
        return value.strip()

    if isinstance(value, (pd.Timestamp, np.datetime64)):
        try:
            return pd.Timestamp(value).isoformat()
        except (TypeError, ValueError):
            return str(value)

    if isinstance(value, (bool, np.bool_)):
        return bool(value)

    if isinstance(value, (np.integer, int)):
        return int(value)

    if isinstance(value, (np.floating, float)):
        numeric = float(value)

        if math.isfinite(numeric) and numeric.is_integer():
            return int(numeric)

        return numeric

    return str(value).strip()


def _canonical_key(value: Any) -> str | None:
    """
    Canonical string used only for comparing values across tables.
    """
    canonical = _canonical_scalar(value)

    if canonical is None:
        return None

    return str(canonical)


def _normalize_key_definition(
    key_definition: Any,
) -> list[str]:
    """Normalize one key definition into a list of column names."""
    if isinstance(key_definition, str):
        return [key_definition]

    if isinstance(key_definition, (tuple, list)):
        columns = [str(column) for column in key_definition]

        if not columns:
            raise ValueError("Key definition cannot be empty.")

        return columns

    raise TypeError(
        "Key definition must be a string, tuple, or list."
    )


def _normalize_primary_key_map(
    primary_key_map: Any,
) -> dict[str, PrimaryKeyDefinition]:
    """
    Normalize externally supplied primary-key metadata.

    This function only handles the internal metadata protocol. It does
    not infer keys from table or column names.
    """
    if primary_key_map is None:
        return {}

    if isinstance(primary_key_map, Mapping):
        return {
            str(table): definition
            for table, definition in primary_key_map.items()
        }

    if isinstance(primary_key_map, (list, tuple)):
        normalized: dict[str, PrimaryKeyDefinition] = {}

        for item in primary_key_map:
            if isinstance(item, Mapping):
                table_name = item.get("table")
                definition = item.get("columns")

                if definition is None:
                    definition = item.get("primary_key")

                if table_name is not None and definition is not None:
                    normalized[str(table_name)] = definition
                    continue

                if len(item) == 1:
                    table_name, definition = next(iter(item.items()))
                    normalized[str(table_name)] = definition
                    continue

            elif isinstance(item, (tuple, list)) and len(item) == 2:
                normalized[str(item[0])] = item[1]

            elif isinstance(item, (tuple, list)) and len(item) > 2:
                normalized[str(item[0])] = list(item[1:])

        return normalized

    raise TypeError(
        "Primary-key metadata must be a mapping, list, tuple, or None."
    )


def _normalize_foreign_key_map(
    foreign_key_map: Any,
) -> dict[str, dict[str, ForeignKeyRelationship]]:
    """
    Normalize externally supplied relationship metadata.

    Canonical internal form:

        {
            child_table: {
                child_column: (parent_table, parent_column)
            }
        }
    """
    if foreign_key_map is None:
        return {}

    if isinstance(foreign_key_map, Mapping):
        normalized: dict[str, dict[str, ForeignKeyRelationship]] = {}

        for child_table, relationships in foreign_key_map.items():
            if not isinstance(relationships, Mapping):
                continue

            normalized[str(child_table)] = {}

            for child_column, relationship in relationships.items():
                if (
                    isinstance(relationship, (tuple, list))
                    and len(relationship) == 2
                ):
                    normalized[str(child_table)][str(child_column)] = (
                        str(relationship[0]),
                        str(relationship[1]),
                    )

        return normalized

    if isinstance(foreign_key_map, (list, tuple)):
        normalized = {}

        for item in foreign_key_map:
            if not isinstance(item, Mapping):
                continue

            child_table = item.get("table")
            if child_table is None:
                child_table = item.get("child_table")

            child_column = item.get("column")
            if child_column is None:
                child_column = item.get("foreign_key")

            parent_table = item.get("referenced_table")
            if parent_table is None:
                parent_table = item.get("parent_table")

            parent_column = item.get("referenced_key")
            if parent_column is None:
                parent_column = item.get("parent_key")

            if (
                child_table is None
                or child_column is None
                or parent_table is None
                or parent_column is None
            ):
                continue

            normalized.setdefault(str(child_table), {})[
                str(child_column)
            ] = (str(parent_table), str(parent_column))

        return normalized

    raise TypeError(
        "Foreign-key metadata must be a mapping, list, tuple, or None."
    )


# ============================================================
# STRUCTURAL VALIDATION
# ============================================================

def validate_dataset_structure(
    tables: Mapping[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Validate basic table structure.

    No fixed table count or table name is required.
    """
    if not isinstance(tables, Mapping):
        return [{
            "scope": "dataset",
            "status": "ERROR",
            "valid": False,
            "table": None,
            "row_count": 0,
            "column_count": 0,
            "duplicate_columns": [],
            "message": "Dataset tables must be provided as a mapping.",
        }]

    if not tables:
        return [{
            "scope": "dataset",
            "status": "ERROR",
            "valid": False,
            "table": None,
            "row_count": 0,
            "column_count": 0,
            "duplicate_columns": [],
            "message": "No tables were detected.",
        }]

    results: list[dict[str, Any]] = []

    for table_name, df in tables.items():
        result: dict[str, Any] = {
            "scope": "table",
            "table": str(table_name),
            "status": "ERROR",
            "valid": False,
            "row_count": 0,
            "column_count": 0,
            "duplicate_columns": [],
            "empty": False,
        }

        if not isinstance(df, pd.DataFrame):
            result["message"] = "Table is not a pandas DataFrame."
            results.append(result)
            continue

        result["row_count"] = len(df)
        result["column_count"] = len(df.columns)
        result["empty"] = len(df) == 0

        duplicate_columns = (
            df.columns[df.columns.duplicated()].tolist()
        )
        result["duplicate_columns"] = duplicate_columns

        if duplicate_columns:
            result["message"] = "Duplicate column names detected."
            results.append(result)
            continue

        if len(df.columns) == 0:
            result["message"] = "Table contains no columns."
            results.append(result)
            continue

        if len(df) == 0:
            result["status"] = "WARNING"
            result["valid"] = True
            result["message"] = (
                "Table is structurally valid but contains no records."
            )
            results.append(result)
            continue

        result["status"] = "PASS"
        result["valid"] = True
        result["message"] = "Table structure is valid."
        results.append(result)

    return results


# ============================================================
# MISSING / BLANK VALIDATION
# ============================================================

def validate_null_values(
    tables: Mapping[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Inspect NULL and blank values for every discovered table/column.

    Results are column-level so downstream modules can identify the
    exact affected fields.
    """
    results: list[dict[str, Any]] = []

    for table_name, df in tables.items():
        if not isinstance(df, pd.DataFrame):
            continue

        for column in df.columns:
            series = df[column]

            null_count = _safe_int(series.isna().sum())

            try:
                blank_count = _safe_int(
                    series.map(_is_blank_value).sum()
                )
            except (TypeError, ValueError):
                blank_count = 0

            null_valid = (
                null_count == 0
                if REQUIRE_NON_NULL_VALUES
                else True
            )

            blank_valid = (
                blank_count == 0
                if REQUIRE_NON_BLANK_VALUES
                else True
            )

            valid = null_valid and blank_valid

            if valid:
                status = "PASS"
            else:
                status = "ERROR"

            results.append({
                "table": str(table_name),
                "column": str(column),
                "null_count": null_count,
                "blank_count": blank_count,
                "status": status,
                "valid": valid,
            })

    return results


# ============================================================
# DUPLICATE ROW VALIDATION
# ============================================================

def validate_duplicate_rows(
    tables: Mapping[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Report exact duplicate rows.

    Duplicate rows are reported as data-quality information rather than
    universally treated as an error, because duplicate observations may
    be legitimate in some datasets.
    """
    results: list[dict[str, Any]] = []

    for table_name, df in tables.items():
        if not isinstance(df, pd.DataFrame):
            continue

        duplicate_mask = df.duplicated(keep=False)
        duplicate_row_count = _safe_int(duplicate_mask.sum())

        results.append({
            "table": str(table_name),
            "duplicate_row_count": duplicate_row_count,
            "has_duplicates": duplicate_row_count > 0,
            "status": (
                "WARNING"
                if duplicate_row_count > 0
                else "PASS"
            ),
            "valid": True,
        })

    return results


# ============================================================
# PRIMARY KEY VALIDATION
# ============================================================

def validate_primary_keys(
    tables: Mapping[str, pd.DataFrame],
    primary_key_map: Any = None,
) -> list[dict[str, Any]]:
    """
    Validate only explicitly supplied primary-key definitions.

    Absence of a configured key is not an error. Generic key candidates
    are handled separately by structural-key validation.
    """
    normalized = _normalize_primary_key_map(primary_key_map)
    results: list[dict[str, Any]] = []

    for table_name, definition in normalized.items():
        df = tables.get(table_name)

        result: dict[str, Any] = {
            "table": table_name,
            "primary_key": [],
            "exists": False,
            "null_count": 0,
            "blank_count": 0,
            "duplicate_row_count": 0,
            "missing_columns": [],
            "status": "ERROR",
            "valid": False,
        }

        try:
            columns = _normalize_key_definition(definition)
        except (TypeError, ValueError) as exc:
            result["message"] = str(exc)
            results.append(result)
            continue

        result["primary_key"] = columns

        if df is None:
            result["message"] = "Configured table is not present."
            results.append(result)
            continue

        missing = [column for column in columns if column not in df.columns]

        if missing:
            result["missing_columns"] = missing
            result["message"] = "Configured key column(s) are missing."
            results.append(result)
            continue

        result["exists"] = True

        key_frame = df[columns]

        result["null_count"] = _safe_int(
            key_frame.isna().any(axis=1).sum()
        )

        result["blank_count"] = _safe_int(
            key_frame.map(_is_blank_value).any(axis=1).sum()
        )

        result["duplicate_row_count"] = _safe_int(
            key_frame.duplicated(keep=False).sum()
        )

        null_valid = (
            result["null_count"] == 0
            if REQUIRE_NON_NULL_VALUES
            else True
        )

        blank_valid = (
            result["blank_count"] == 0
            if REQUIRE_NON_BLANK_VALUES
            else True
        )

        duplicate_valid = (
            result["duplicate_row_count"] == 0
            if REQUIRE_UNIQUE_PRIMARY_KEYS
            else True
        )

        result["valid"] = (
            null_valid
            and blank_valid
            and duplicate_valid
        )

        result["status"] = (
            "PASS"
            if result["valid"]
            else "ERROR"
        )

        if result["valid"]:
            result["message"] = "Configured primary key is valid."
        else:
            result["message"] = "Configured primary key has violations."

        results.append(result)

    return results


# ============================================================
# STRUCTURAL KEY EVIDENCE
# ============================================================

def validate_structural_key_candidates(
    tables: Mapping[str, pd.DataFrame],
    profiler_report: Mapping[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """
    Validate structural key candidates against the full supplied
    DataFrames.

    Profiler candidates are treated as evidence only. Uniqueness,
    nulls, and duplicates are checked against the full table before
    reporting the result.
    """
    results: list[dict[str, Any]] = []

    profiler_tables: dict[str, Any] = {}

    if isinstance(profiler_report, Mapping):
        discovery_profiles = profiler_report.get(
            "discovery_profiles"
        )

        if isinstance(discovery_profiles, Mapping):
            profiler_tables = dict(discovery_profiles)
        else:
            detailed_tables = profiler_report.get("tables")

            if isinstance(detailed_tables, Mapping):
                profiler_tables = dict(detailed_tables)

            elif isinstance(detailed_tables, Sequence):
                for table_profile in detailed_tables:
                    if not isinstance(table_profile, Mapping):
                        continue

                    table_name = table_profile.get("table")

                    if table_name is not None:
                        profiler_tables[str(table_name)] = (
                            table_profile
                        )

    for table_name, df in tables.items():
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue

        table_profile = profiler_tables.get(
            str(table_name)
        )

        supplied_candidates: list[Any] = []

        if isinstance(table_profile, Mapping):
            supplied_candidates = (
                table_profile.get("identifier_candidates")
                or table_profile.get("candidate_keys")
                or table_profile.get("structural_key_candidates")
                or table_profile.get("key_candidates")
                or []
            )

        if supplied_candidates:
            for candidate in supplied_candidates:
                if isinstance(candidate, str):
                    columns = [candidate]

                elif isinstance(candidate, Mapping):
                    columns = candidate.get("columns")

                    if columns is None:
                        columns = candidate.get("column")

                    if isinstance(columns, str):
                        columns = [columns]

                else:
                    columns = candidate

                if not isinstance(columns, (list, tuple)):
                    continue

                columns = list(columns)

                if not columns:
                    continue

                missing_columns = [
                    column
                    for column in columns
                    if column not in df.columns
                ]

                if missing_columns:
                    continue

                key_frame = df.loc[:, columns]

                null_count = _safe_int(
                    key_frame.isna().any(axis=1).sum()
                )

                try:
                    blank_count = _safe_int(
                        key_frame.map(
                            _is_blank_value
                        ).any(axis=1).sum()
                    )
                except (TypeError, ValueError):
                    blank_count = 0

                duplicate_count = _safe_int(
                    key_frame.duplicated(
                        keep=False
                    ).sum()
                )

                appears_unique = (
                    len(df) > 0
                    and null_count == 0
                    and blank_count == 0
                    and duplicate_count == 0
                )

                results.append({
                    "table": str(table_name),
                    "columns": columns,
                    "candidate_source": "profiler",
                    "null_count": null_count,
                    "blank_count": blank_count,
                    "duplicate_row_count": duplicate_count,
                    "appears_unique": appears_unique,
                    "candidate_evidence": {
                        "row_count": int(len(df)),
                        "validated_against_full_table": True,
                        "null_count": null_count,
                        "blank_count": blank_count,
                        "duplicate_row_count": duplicate_count,
                    },
                    "status": (
                        "PASS"
                        if appears_unique
                        else "WARNING"
                    ),
                    "valid": True,
                })

            continue

        for column in df.columns:
            series = df[column]

            if series.isna().any():
                continue

            try:
                if series.map(
                    _is_blank_value
                ).any():
                    continue
            except (TypeError, ValueError):
                continue

            unique_count = _safe_int(
                series.nunique(
                    dropna=False
                )
            )

            if unique_count == len(df) and len(df) > 0:
                results.append({
                    "table": str(table_name),
                    "columns": [column],
                    "candidate_source": "structural_fallback",
                    "null_count": 0,
                    "blank_count": 0,
                    "duplicate_row_count": 0,
                    "appears_unique": True,
                    "candidate_evidence": {
                        "row_count": int(len(df)),
                        "validated_against_full_table": True,
                        "null_count": 0,
                        "blank_count": 0,
                        "duplicate_row_count": 0,
                    },
                    "status": "INFO",
                    "valid": True,
                })

    return results


# ============================================================
# FOREIGN KEY / RELATIONSHIP VALIDATION
# ============================================================

def _validate_one_relationship(
    child_table: str,
    child_column: str,
    parent_table: str,
    parent_column: str,
    tables: Mapping[str, pd.DataFrame],
) -> dict[str, Any]:
    """Validate one explicitly supplied relationship."""
    result: dict[str, Any] = {
        "table": child_table,
        "foreign_key": child_column,
        "referenced_table": parent_table,
        "referenced_key": parent_column,
        "null_count": 0,
        "blank_count": 0,
        "invalid_row_count": 0,
        "unique_invalid_value_count": 0,
        "invalid_values": [],
        "status": "ERROR",
        "valid": False,
    }

    child_df = tables.get(child_table)
    parent_df = tables.get(parent_table)

    if not isinstance(child_df, pd.DataFrame):
        result["message"] = "Child table is not present."
        return result

    if not isinstance(parent_df, pd.DataFrame):
        result["message"] = "Referenced table is not present."
        return result

    if child_column not in child_df.columns:
        result["message"] = "Foreign-key column is missing."
        return result

    if parent_column not in parent_df.columns:
        result["message"] = "Referenced key column is missing."
        return result

    child_series = child_df[child_column]
    parent_series = parent_df[parent_column]

    result["null_count"] = _safe_int(child_series.isna().sum())

    try:
        result["blank_count"] = _safe_int(
            child_series.map(_is_blank_value).sum()
        )
    except (TypeError, ValueError):
        result["blank_count"] = 0

    parent_values = {
        canonical
        for canonical in (
            _canonical_key(value)
            for value in parent_series.tolist()
        )
        if canonical is not None
    }

    invalid_values: list[str] = []
    invalid_row_count = 0

    for value in child_series.tolist():
        if _is_missing_scalar(value) or _is_blank_value(value):
            continue

        canonical = _canonical_key(value)

        if canonical not in parent_values:
            invalid_row_count += 1

            if canonical is not None:
                invalid_values.append(canonical)

    unique_invalid_values = sorted(set(invalid_values))

    result["invalid_row_count"] = invalid_row_count
    result["unique_invalid_value_count"] = len(unique_invalid_values)
    result["invalid_values"] = unique_invalid_values[
        :MAX_INVALID_VALUES_REPORTED
    ]

    null_valid = (
        result["null_count"] == 0
        if REQUIRE_NON_NULL_VALUES
        else True
    )

    blank_valid = (
        result["blank_count"] == 0
        if REQUIRE_NON_BLANK_VALUES
        else True
    )

    reference_valid = (
        result["invalid_row_count"] == 0
        if REQUIRE_VALID_FOREIGN_KEYS
        else True
    )

    result["valid"] = (
        null_valid
        and blank_valid
        and reference_valid
    )

    result["status"] = (
        "PASS"
        if result["valid"]
        else "ERROR"
    )

    result["message"] = (
        "Relationship is valid."
        if result["valid"]
        else "Relationship has referential-integrity violations."
    )

    return result


def validate_foreign_keys(
    tables: Mapping[str, pd.DataFrame],
    foreign_key_map: Any = None,
    relationship_candidates: Sequence[Any] | None = None,
) -> list[dict[str, Any]]:
    """
    Validate explicitly supplied relationships.

    Discovered relationship candidates can also be supplied. Candidate
    relationships are reported as evidence and are never automatically
    treated as mandatory foreign keys.
    """
    normalized = _normalize_foreign_key_map(foreign_key_map)
    results: list[dict[str, Any]] = []

    for child_table, relationships in normalized.items():
        for child_column, relationship in relationships.items():
            parent_table, parent_column = relationship

            results.append(
                _validate_one_relationship(
                    child_table,
                    child_column,
                    parent_table,
                    parent_column,
                    tables,
                )
            )

    # Discovered candidates are informational unless explicitly promoted
    # into the foreign_key_map.
    if relationship_candidates:
        for candidate in relationship_candidates:
            if not isinstance(candidate, Mapping):
                continue

            child_table = (
                candidate.get("child_table")
                or candidate.get("table")
            )
            child_column = (
                candidate.get("child_column")
                or candidate.get("foreign_key")
                or candidate.get("column")
            )
            parent_table = (
                candidate.get("parent_table")
                or candidate.get("referenced_table")
            )
            parent_column = (
                candidate.get("parent_column")
                or candidate.get("referenced_key")
            )

            if not all(
                value is not None
                for value in (
                    child_table,
                    child_column,
                    parent_table,
                    parent_column,
                )
            ):
                continue

            results.append({
                "table": str(child_table),
                "foreign_key": str(child_column),
                "referenced_table": str(parent_table),
                "referenced_key": str(parent_column),
                "candidate": True,
                "status": "INFO",
                "valid": True,
                "message": (
                    "Relationship was structurally discovered; "
                    "it is not treated as mandatory."
                ),
            })

    return results


# ============================================================
# NUMERIC VALIDATION
# ============================================================

def validate_numeric_values(
    tables: Mapping[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Validate numeric columns for non-finite values.

    No business interpretation is applied to negative/zero values.
    """
    results: list[dict[str, Any]] = []

    if not REQUIRE_FINITE_NUMERIC_VALUES:
        return results

    for table_name, df in tables.items():
        if not isinstance(df, pd.DataFrame):
            continue

        numeric_columns = df.select_dtypes(
            include=["number"]
        ).columns

        for column in numeric_columns:
            series = df[column]

            try:
                values = series.to_numpy(dtype=float)
                non_finite_mask = ~np.isfinite(values)
                violation_count = _safe_int(non_finite_mask.sum())
            except (TypeError, ValueError):
                violation_count = 0

            if violation_count > 0:
                results.append({
                    "table": str(table_name),
                    "column": str(column),
                    "rule": "FINITE_NUMERIC",
                    "violation_count": violation_count,
                    "status": "ERROR",
                    "valid": False,
                    "message": "Non-finite numeric values detected.",
                })

    return results


# ============================================================
# TEMPORAL VALIDATION
# ============================================================

def validate_temporal_values(
    tables: Mapping[str, pd.DataFrame],
    profiler_report: Mapping[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """
    Validate columns already identified as temporal by structural
    profiling.

    Temporal meaning is never inferred from column names.
    """
    results: list[dict[str, Any]] = []

    discovered_profiles: dict[str, Any] = {}

    if isinstance(profiler_report, Mapping):
        canonical_profiles = profiler_report.get(
            "discovery_profiles"
        )

        if isinstance(canonical_profiles, Mapping):
            discovered_profiles = dict(
                canonical_profiles
            )

        else:
            detailed_profiles = profiler_report.get(
                "tables"
            )

            if isinstance(detailed_profiles, Mapping):
                discovered_profiles = dict(
                    detailed_profiles
                )

            elif isinstance(detailed_profiles, Sequence):
                for table_profile in detailed_profiles:
                    if not isinstance(
                        table_profile,
                        Mapping,
                    ):
                        continue

                    table_name = table_profile.get(
                        "table"
                    )

                    if table_name is not None:
                        discovered_profiles[
                            str(table_name)
                        ] = table_profile

    for table_name, df in tables.items():
        if not isinstance(df, pd.DataFrame):
            continue

        table_profile = discovered_profiles.get(
            str(table_name),
            {},
        )

        profiled_temporal_columns: list[Any] = []

        if isinstance(table_profile, Mapping):
            profiled_temporal_columns = (
                table_profile.get(
                    "temporal_columns"
                )
                or table_profile.get(
                    "date_like_columns"
                )
                or []
            )

            if not profiled_temporal_columns:
                columns_profile = table_profile.get(
                    "columns"
                )

                if isinstance(columns_profile, Sequence):
                    for column_profile in columns_profile:
                        if not isinstance(
                            column_profile,
                            Mapping,
                        ):
                            continue

                        if column_profile.get(
                            "inferred_type"
                        ) == "temporal":
                            column_name = (
                                column_profile.get(
                                    "name"
                                )
                            )

                            if column_name is not None:
                                profiled_temporal_columns.append(
                                    column_name
                                )

        for column in df.columns:
            series = df[column]

            is_temporal_dtype = (
                pd.api.types.is_datetime64_any_dtype(
                    series
                )
                or isinstance(
                    series.dtype,
                    pd.DatetimeTZDtype,
                )
            )

            if not is_temporal_dtype:
                if column not in profiled_temporal_columns:
                    continue

            try:
                converted = pd.to_datetime(
                    series,
                    errors="coerce",
                    format="mixed",
                )

                original_present = ~series.isna()
                invalid_mask = (
                    original_present
                    & converted.isna()
                )

                invalid_count = _safe_int(
                    invalid_mask.sum()
                )

            except (
                TypeError,
                ValueError,
                OverflowError,
            ):
                invalid_count = 0

            if invalid_count > 0:
                results.append({
                    "table": str(table_name),
                    "column": column,
                    "rule": "TEMPORAL_PARSE",
                    "violation_count": invalid_count,
                    "status": "ERROR",
                    "valid": False,
                    "message": (
                        "Invalid temporal values detected."
                    ),
                })

    return results


# ============================================================
# OPTIONAL TYPE-CONSISTENCY VALIDATION
# ============================================================

def validate_type_consistency(
    tables: Mapping[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Report columns with mixed Python scalar types.

    This is informational because heterogeneous values can be valid
    depending on the source format and pandas representation.
    """
    results: list[dict[str, Any]] = []

    for table_name, df in tables.items():
        if not isinstance(df, pd.DataFrame):
            continue

        for column in df.columns:
            series = df[column].dropna()

            if len(series) == 0:
                continue

            if not (
                pd.api.types.is_object_dtype(series.dtype)
                or pd.api.types.is_string_dtype(series.dtype)
            ):
                continue

            type_names = sorted({
                type(value).__name__
                for value in series
            })

            if len(type_names) > 1:
                results.append({
                    "table": str(table_name),
                    "column": str(column),
                    "observed_types": type_names,
                    "status": "WARNING",
                    "valid": True,
                    "message": (
                        "Multiple scalar Python types were observed."
                    ),
                })

    return results


# ============================================================
# REPORT AGGREGATION
# ============================================================

def _count_status(
    results: Sequence[Mapping[str, Any]],
    status: str,
) -> int:
    return sum(
        1
        for result in results
        if str(result.get("status", "")).upper() == status
    )


def _has_errors(
    results: Sequence[Mapping[str, Any]],
) -> bool:
    return any(
        str(result.get("status", "")).upper() == "ERROR"
        or result.get("valid") is False
        for result in results
    )


def _has_warnings(
    results: Sequence[Mapping[str, Any]],
) -> bool:
    return any(
        str(result.get("status", "")).upper() == "WARNING"
        for result in results
    )


# ============================================================
# COMPLETE DATASET VALIDATION
# ============================================================

def validate_dataset(
    tables: Mapping[str, pd.DataFrame],
    primary_key_map: Any = None,
    foreign_key_map: Any = None,
    profiler_report: Mapping[str, Any] | None = None,
    relationship_candidates: Sequence[Any] | None = None,
) -> dict[str, Any]:
    """
    Run the complete generic validation suite.

    primary_key_map and foreign_key_map are optional. If omitted, the
    validator does not invent mandatory schema rules.

    profiler_report and relationship_candidates are optional structural
    evidence from earlier pipeline stages.
    """
    structure = validate_dataset_structure(tables)

    null_values = validate_null_values(tables)

    duplicate_rows = validate_duplicate_rows(tables)

    primary_keys = validate_primary_keys(
        tables,
        primary_key_map,
    )

    structural_keys = validate_structural_key_candidates(
        tables,
        profiler_report,
    )

    foreign_keys = validate_foreign_keys(
        tables,
        foreign_key_map,
        relationship_candidates,
    )

    numeric_values = validate_numeric_values(tables)

    temporal_values = validate_temporal_values(
        tables,
        profiler_report,
    )

    type_consistency = validate_type_consistency(tables)

    sections = {
        "structure": structure,
        "null_values": null_values,
        "duplicate_rows": duplicate_rows,
        "primary_keys": primary_keys,
        "structural_key_candidates": structural_keys,
        "foreign_keys": foreign_keys,
        "numeric_values": numeric_values,
        "temporal_values": temporal_values,
        "type_consistency": type_consistency,
    }

    all_results = [
        result
        for section in sections.values()
        for result in section
    ]

    has_errors = _has_errors(all_results)
    has_warnings = _has_warnings(all_results)

    if has_errors:
        overall_status = "INVALID"
    elif has_warnings:
        overall_status = "VALID_WITH_WARNINGS"
    else:
        overall_status = "VALID"

    return {
        "valid": not has_errors,
        "status": overall_status,
        "summary": {
            "table_count": len(tables)
            if isinstance(tables, Mapping)
            else 0,
            "total_findings": len(all_results),
            "errors": _count_status(all_results, "ERROR"),
            "warnings": _count_status(all_results, "WARNING"),
            "info": _count_status(all_results, "INFO"),
            "passes": _count_status(all_results, "PASS"),
            "structure_valid": not _has_errors(structure),
            "null_values_valid": not _has_errors(null_values),
            "primary_keys_valid": not _has_errors(primary_keys),
            "foreign_keys_valid": not _has_errors(
                [
                    result
                    for result in foreign_keys
                    if not result.get("candidate", False)
                ]
            ),
            "numeric_values_valid": not _has_errors(numeric_values),
            "temporal_values_valid": not _has_errors(temporal_values),
        },
        **sections,
    }


# ============================================================
# REPORT PRINTING
# ============================================================

def _print_section(
    title: str,
    results: Sequence[Mapping[str, Any]],
) -> None:
    print()
    print(title)

    if not results:
        print("  No findings.")
        return

    for result in results:
        status = str(
            result.get("status", "INFO")
        ).upper()

        table = result.get("table")
        column = result.get("column")
        columns = result.get("columns")

        location = str(table) if table is not None else "dataset"

        if column is not None:
            location += f".{column}"
        elif columns:
            if isinstance(columns, (list, tuple)):
                location += "." + ", ".join(
                    str(value) for value in columns
                )
            else:
                location += f".{columns}"

        message = result.get("message", "")

        print(
            f"  [{status}] {location}"
            + (f" - {message}" if message else "")
        )

        if result.get("null_count", 0):
            print(
                f"       NULL values: {result['null_count']}"
            )

        if result.get("blank_count", 0):
            print(
                f"       Blank values: {result['blank_count']}"
            )

        if result.get("duplicate_row_count", 0):
            print(
                "       Duplicate rows: "
                f"{result['duplicate_row_count']}"
            )

        if result.get("invalid_row_count", 0):
            print(
                "       Invalid rows: "
                f"{result['invalid_row_count']}"
            )

        if result.get("violation_count", 0):
            print(
                "       Violations: "
                f"{result['violation_count']}"
            )


def print_validation_report(
    report: Mapping[str, Any],
) -> None:
    """Print a compact human-readable validation report."""
    print()
    print("=" * 72)
    print("GENERIC DATASET VALIDATION REPORT")
    print("=" * 72)

    print()
    print(
        "Overall status : "
        f"{report.get('status', 'UNKNOWN')}"
    )

    summary = report.get("summary", {})

    print(
        "Table count    : "
        f"{summary.get('table_count', 0)}"
    )

    print(
        "Errors         : "
        f"{summary.get('errors', 0)}"
    )

    print(
        "Warnings       : "
        f"{summary.get('warnings', 0)}"
    )

    print(
        "Information    : "
        f"{summary.get('info', 0)}"
    )

    _print_section(
        "STRUCTURE VALIDATION",
        report.get("structure", []),
    )

    _print_section(
        "NULL / BLANK VALIDATION",
        report.get("null_values", []),
    )

    _print_section(
        "DUPLICATE ROW VALIDATION",
        report.get("duplicate_rows", []),
    )

    _print_section(
        "PRIMARY KEY VALIDATION",
        report.get("primary_keys", []),
    )

    _print_section(
        "STRUCTURAL KEY EVIDENCE",
        report.get("structural_key_candidates", []),
    )

    _print_section(
        "RELATIONSHIP VALIDATION",
        report.get("foreign_keys", []),
    )

    _print_section(
        "NUMERIC VALIDATION",
        report.get("numeric_values", []),
    )

    _print_section(
        "TEMPORAL VALIDATION",
        report.get("temporal_values", []),
    )

    _print_section(
        "TYPE CONSISTENCY",
        report.get("type_consistency", []),
    )

    print()
    print("=" * 72)


# ============================================================
# REPORT EXPORT
# ============================================================

def save_validation_report(
    report: Mapping[str, Any],
    output_path: str | Path | None = None,
) -> Path:
    """
    Save the structured validation report as JSON.

    The output path is generic and never depends on a dataset name.
    """
    if output_path is None:
        output_path = (
            Path(getattr(
                config,
                "PROCESSED_DATA_DIRECTORY",
                PROJECT_ROOT / "data" / "processed",
            ))
            / "validation_report.json"
        )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            report,
            handle,
            indent=2,
            ensure_ascii=False,
            default=str,
        )

    return output_path


# ============================================================
# DATA LOADER INTEGRATION
# ============================================================

def _load_tables_from_data_loader() -> dict[str, pd.DataFrame]:
    """
    Discover and load the active dataset through data_loader.py.

    The current data_loader API exposes ``discover_dataset()`` rather than
    a separate ``load_dataset()`` function.  The returned discovery object
    contains the parsed tables plus structural metadata.
    """
    try:
        from src.data_loader import discover_dataset
    except ImportError:
        from data_loader import discover_dataset

    result = discover_dataset()

    if not isinstance(result, Mapping):
        raise TypeError(
            "data_loader.discover_dataset() did not return a mapping."
        )

    data = result.get("tables")
    if not isinstance(data, Mapping):
        raise TypeError(
            "data_loader.discover_dataset() returned no table mapping."
        )

    tables: dict[str, pd.DataFrame] = {}

    for table_name, table_data in data.items():
        if isinstance(table_data, pd.DataFrame):
            tables[str(table_name)] = table_data

        elif isinstance(table_data, list):
            tables[str(table_name)] = pd.DataFrame(table_data)

        elif isinstance(table_data, Mapping):
            try:
                tables[str(table_name)] = pd.DataFrame(table_data)
            except (TypeError, ValueError):
                try:
                    tables[str(table_name)] = pd.DataFrame(
                        list(table_data.values())
                    )
                except (TypeError, ValueError):
                    continue

    return tables


# ============================================================
# OPTIONAL PROFILER METADATA
# ============================================================

def _build_profiler_evidence(
    tables: Mapping[str, pd.DataFrame],
) -> dict[str, Any]:
    """
    Obtain the canonical structural evidence produced by
    data_loader.py.

    The standalone validator uses the same discovery contract as the
    rest of the pipeline instead of reconstructing a separate partial
    profiler.
    """
    try:
        from src.data_loader import discover_dataset
    except ImportError:
        from data_loader import discover_dataset

    discovery = discover_dataset()

    if not isinstance(discovery, Mapping):
        raise TypeError(
            "data_loader.discover_dataset() did not return "
            "a mapping."
        )

    return {
        "dataset": discovery.get(
            "dataset",
            {},
        ),
        "summary": discovery.get(
            "summary",
            {},
        ),
        "discovery_profiles": discovery.get(
            "profiles",
            {},
        ),
        "tables": discovery.get(
            "profiles",
            {},
        ),
        "discovery_roles": discovery.get(
            "roles",
            {},
        ),
        "discovery_relationships": discovery.get(
            "relationships",
            [],
        ),
        "discovery_reuse_candidates": discovery.get(
            "reuse_candidates",
            [],
        ),
        "quality": discovery.get(
            "quality",
            {},
        ),
        "ml_risks": discovery.get(
            "ml_risks",
            {},
        ),
    }


# ============================================================
# STANDALONE MODULE TEST
# ============================================================

def main() -> None:
    """
    Standalone validation run.

    Uses one canonical data_loader discovery result so the same
    tables and structural evidence are validated without loading
    the dataset twice.
    """
    print("=" * 72)
    print("ENTERPRISE DEMAND INTELLIGENCE PLATFORM")
    print("GENERIC DATA VALIDATION")
    print("=" * 72)

    print()
    print("Loading dataset through data_loader.py...")

    try:
        from src.data_loader import discover_dataset
    except ImportError:
        from data_loader import discover_dataset

    discovery = discover_dataset()

    if not isinstance(discovery, Mapping):
        raise TypeError(
            "data_loader.discover_dataset() did not return "
            "a mapping."
        )

    tables = discovery.get("tables")

    if not isinstance(tables, Mapping):
        raise TypeError(
            "data_loader.discover_dataset() returned no table mapping."
        )

    print(
        f"Dataset loading: PASSED "
        f"({len(tables)} tables)"
    )

    if not tables:
        raise RuntimeError(
            "No usable DataFrame tables were produced by "
            "data_loader.py."
        )

    profiler_evidence = {
        "dataset": discovery.get(
            "dataset",
            {},
        ),
        "summary": discovery.get(
            "summary",
            {},
        ),
        "discovery_profiles": discovery.get(
            "profiles",
            {},
        ),
        "tables": discovery.get(
            "profiles",
            {},
        ),
        "discovery_roles": discovery.get(
            "roles",
            {},
        ),
        "discovery_relationships": discovery.get(
            "relationships",
            [],
        ),
        "discovery_reuse_candidates": discovery.get(
            "reuse_candidates",
            [],
        ),
        "quality": discovery.get(
            "quality",
            {},
        ),
        "ml_risks": discovery.get(
            "ml_risks",
            {},
        ),
    }

    relationship_candidates = discovery.get(
        "relationships",
        [],
    )

    print()
    print("Running structural validation...")

    report = validate_dataset(
        tables,
        profiler_report=profiler_evidence,
        relationship_candidates=relationship_candidates,
    )

    print_validation_report(report)

    try:
        report_path = save_validation_report(report)

        print()
        print(
            "Validation report saved to:"
            f"\n  {report_path}"
        )
    except (OSError, TypeError, ValueError) as exc:
        print()
        print(
            "Validation report could not be saved:"
            f" {exc}"
        )

    print()

    if report["status"] == "VALID":
        print("Validation completed successfully.")
    elif report["status"] == "VALID_WITH_WARNINGS":
        print(
            "Validation completed with warnings. "
            "No blocking structural errors were detected."
        )
    else:
        print(
            "Validation completed with data-quality errors."
        )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    main()

