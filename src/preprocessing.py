"""
Enterprise Demand Intelligence Platform
GENERIC ANALYTICAL DATA PREPARATION

Purpose
-------
Prepare the cleaned dataset for downstream analysis, forecasting,
recommendation generation, and Power BI without requiring a fixed business
schema.

Core rule
---------
This module is built for the Enterprise Demand Intelligence problem, not for
any particular demo dataset. It discovers what the cleaned dataset actually
contains and adapts to it.

The module never requires:
- a particular dataset filename
- a particular table name
- a particular column name
- a particular entity count
- a particular date range
- a particular business domain
- a particular demand target

Safe responsibilities
---------------------
- discover and load the processed dataset
- preserve every discovered table
- preserve every source column
- add reusable structural/calendar features when supported by discovered
  data types
- keep transformations row-level
- avoid unnecessary aggregation
- avoid business-specific semantic guesses
- write analytical output separately
- create a machine-readable preprocessing report

This module deliberately does NOT:
- infer a demand target
- invent business aliases
- apply business-specific thresholds
- classify columns from their names
- aggregate by assumed business entities
- impute missing values
- remove outliers
- rename source business columns
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any, Mapping

import numpy as np
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


PROCESSED_DIRECTORY = Path(
    getattr(
        config,
        "PROCESSED_DATA_DIRECTORY",
        PROJECT_ROOT / "data" / "processed",
    )
).expanduser().resolve()

OUTPUT_DIRECTORY = Path(
    getattr(
        config,
        "OUTPUT_DATA_DIRECTORY",
        PROJECT_ROOT / "data" / "output",
    )
).expanduser().resolve()

PROCESSED_DIRECTORY.mkdir(
    parents=True,
    exist_ok=True,
)

OUTPUT_DIRECTORY.mkdir(
    parents=True,
    exist_ok=True,
)

PROCESSED_JSON_PATH = (
    PROCESSED_DIRECTORY / "processed_dataset.json"
)

PROCESSED_EXCEL_PATH = (
    PROCESSED_DIRECTORY / "processed_dataset.xlsx"
)

ANALYTICAL_JSON_PATH = (
    OUTPUT_DIRECTORY / "analytical_dataset.json"
)

ANALYTICAL_EXCEL_PATH = (
    OUTPUT_DIRECTORY / "analytical_dataset.xlsx"
)

PREPROCESSING_REPORT_PATH = (
    OUTPUT_DIRECTORY / "preprocessing_report.json"
)


# ============================================================
# DISPLAY
# ============================================================

SEPARATOR = "=" * 72
SUB_SEPARATOR = "-" * 72


def print_header(title: str) -> None:
    print()
    print(SEPARATOR)
    print(title)
    print(SEPARATOR)


def print_section(title: str) -> None:
    print()
    print(SUB_SEPARATOR)
    print(title)
    print(SUB_SEPARATOR)


# ============================================================
# SOURCE RESOLUTION
# ============================================================

def resolve_processed_source() -> Path:
    """
    Resolve the processed dataset produced by cleaning.py.

    Resolution is intentionally generic. The cleaner's generic output is
    preferred. No demo dataset filename is referenced.
    """

    candidates = [
        PROCESSED_JSON_PATH,
        PROCESSED_EXCEL_PATH,
    ]

    configured_directory = getattr(
        config,
        "PROCESSED_DATA_DIRECTORY",
        None,
    )

    if configured_directory:
        directory = Path(configured_directory)

        candidates.extend([
            directory / PROCESSED_JSON_PATH.name,
            directory / PROCESSED_EXCEL_PATH.name,
        ])

    seen: set[Path] = set()

    for candidate in candidates:
        candidate = candidate.resolve()

        if candidate in seen:
            continue

        seen.add(candidate)

        if candidate.exists() and candidate.is_file():
            return candidate

    message = (
        "No generic processed dataset was found.\n"
        "Expected the output created by cleaning.py at one of:\n"
        + "\n".join(
            f"  {candidate}"
            for candidate in seen
        )
    )

    raise FileNotFoundError(message)


# ============================================================
# JSON HELPERS
# ============================================================

def _is_missing(value: Any) -> bool:
    if value is None:
        return True

    try:
        result = pd.isna(value)
    except Exception:
        return False

    return bool(result) if isinstance(result, bool) else False


def _json_safe(value: Any) -> Any:
    if _is_missing(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.isoformat()

    try:
        if hasattr(value, "item"):
            return _json_safe(value.item())
    except Exception:
        pass

    if isinstance(value, (str, int, float, bool)):
        return value

    if isinstance(value, (list, tuple)):
        return [_json_safe(item) for item in value]

    if isinstance(value, dict):
        return {
            str(key): _json_safe(item)
            for key, item in value.items()
        }

    return str(value)


# ============================================================
# DATA LOADING
# ============================================================

def _load_json(path: Path) -> dict[str, pd.DataFrame]:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    if not isinstance(payload, Mapping):
        raise RuntimeError(
            "Processed JSON root must be a mapping of discovered tables."
        )

    tables: dict[str, pd.DataFrame] = {}

    for table_name, records in payload.items():
        if isinstance(records, list):
            tables[str(table_name)] = pd.DataFrame(records)
        elif isinstance(records, Mapping):
            nested = records.get("records")

            if isinstance(nested, list):
                tables[str(table_name)] = pd.DataFrame(nested)
            else:
                raise RuntimeError(
                    f"Table {table_name!r} does not contain record data."
                )
        else:
            raise RuntimeError(
                f"Table {table_name!r} has an unsupported JSON structure."
            )

    return tables


def _load_excel(path: Path) -> dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(path)

    if not workbook.sheet_names:
        raise RuntimeError(
            "Processed workbook contains no worksheets."
        )

    tables: dict[str, pd.DataFrame] = {}

    for sheet_name in workbook.sheet_names:
        tables[str(sheet_name)] = pd.read_excel(
            path,
            sheet_name=sheet_name,
        )

    return tables


def load_processed_dataset() -> tuple[Path, dict[str, pd.DataFrame]]:
    """
    Load the generic processed output produced by cleaning.py.

    The JSON representation is preferred because it preserves the discovered
    table structure without relying on spreadsheet sheet-name limitations.
    """

    source = resolve_processed_source()

    if source.suffix.lower() == ".json":
        tables = _load_json(source)
    elif source.suffix.lower() in {".xlsx", ".xls"}:
        tables = _load_excel(source)
    else:
        raise RuntimeError(
            f"Unsupported processed dataset format: {source.suffix}"
        )

    if not tables:
        raise RuntimeError(
            "The processed dataset contains no discovered tables."
        )

    return source, tables


# ============================================================
# STRUCTURAL TYPE DISCOVERY
# ============================================================

def _structural_column_kind(
    series: pd.Series,
) -> str:
    """
    Determine a structural data type from dtype and observed values.

    Column names are never used.
    """

    if pd.api.types.is_bool_dtype(series.dtype):
        return "boolean"

    if pd.api.types.is_numeric_dtype(series.dtype):
        return "numeric"

    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        return "temporal"

    if isinstance(series.dtype, pd.DatetimeTZDtype):
        return "temporal"

    if (
        pd.api.types.is_object_dtype(series.dtype)
        or pd.api.types.is_string_dtype(series.dtype)
    ):
        non_missing = series.dropna()

        if non_missing.empty:
            return "empty"

        try:
            parsed = pd.to_datetime(
                non_missing.astype(str).str.strip(),
                errors="coerce",
                format="mixed",
            )

            if float(parsed.notna().mean()) >= 0.95:
                return "temporal"

        except (TypeError, ValueError):
            pass

        return "text"

    return "other"


def _discover_column_kinds(
    dataframe: pd.DataFrame,
) -> dict[Any, str]:
    return {
        column: _structural_column_kind(
            dataframe[column]
        )
        for column in dataframe.columns
    }


# ============================================================
# SAFE FEATURE NAME GENERATION
# ============================================================

def _technical_name_fragment(
    column_name: Any,
) -> str:
    """
    Convert an actual source column label into a safe technical feature-name
    fragment.

    This does not map business vocabulary to another vocabulary; it only
    sanitizes the label already present in the dataset.
    """

    text = str(column_name).strip()

    text = re.sub(
        r"[^A-Za-z0-9]+",
        "_",
        text,
    )

    text = re.sub(
        r"_+",
        "_",
        text,
    ).strip("_")

    return text or "field"


def _unique_feature_name(
    dataframe: pd.DataFrame,
    base_name: str,
) -> str:
    """
    Return a collision-safe technical column name.
    """

    if base_name not in dataframe.columns:
        return base_name

    counter = 2

    while f"{base_name}_{counter}" in dataframe.columns:
        counter += 1

    return f"{base_name}_{counter}"


def _add_feature(
    dataframe: pd.DataFrame,
    name: str,
    values: pd.Series,
) -> str:
    """
    Add a feature without overwriting an existing source/derived column.
    """

    final_name = _unique_feature_name(
        dataframe,
        name,
    )

    dataframe[final_name] = values
    return final_name


# ============================================================
# TEMPORAL FEATURES
# ============================================================

def _add_temporal_features(
    dataframe: pd.DataFrame,
    column_name: Any,
) -> list[str]:
    """
    Add reusable calendar features for every structurally temporal column.

    The source temporal column is preserved.

    If a dataset has several temporal columns, each receives its own derived
    calendar features. No particular temporal field is assumed.
    """

    series = dataframe[column_name]

    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        temporal = series
    elif isinstance(series.dtype, pd.DatetimeTZDtype):
        temporal = series
    else:
        temporal = pd.to_datetime(
            series,
            errors="coerce",
        )

        source_non_null = int(series.notna().sum())
        converted_non_null = int(temporal.notna().sum())

        if source_non_null != converted_non_null:
            raise ValueError(
                f"Temporal feature generation would lose values in "
                f"{column_name!r}: "
                f"{source_non_null:,} source non-null values versus "
                f"{converted_non_null:,} parsed values."
            )

    fragment = _technical_name_fragment(column_name)

    added: list[str] = []

    features: list[tuple[str, pd.Series]] = [
        (
            f"{fragment}__year",
            temporal.dt.year.astype("Int64"),
        ),
        (
            f"{fragment}__month",
            temporal.dt.month.astype("Int64"),
        ),
        (
            f"{fragment}__quarter",
            temporal.dt.quarter.astype("Int64"),
        ),
        (
            f"{fragment}__week",
            temporal.dt.isocalendar().week.astype("Int64"),
        ),
        (
            f"{fragment}__day_of_week",
            temporal.dt.dayofweek.astype("Int64"),
        ),
        (
            f"{fragment}__day_of_month",
            temporal.dt.day.astype("Int64"),
        ),
        (
            f"{fragment}__day_of_year",
            temporal.dt.dayofyear.astype("Int64"),
        ),
        (
            f"{fragment}__is_month_start",
            temporal.dt.is_month_start.astype("boolean"),
        ),
        (
            f"{fragment}__is_month_end",
            temporal.dt.is_month_end.astype("boolean"),
        ),
        (
            f"{fragment}__is_quarter_start",
            temporal.dt.is_quarter_start.astype("boolean"),
        ),
        (
            f"{fragment}__is_quarter_end",
            temporal.dt.is_quarter_end.astype("boolean"),
        ),
    ]

    for base_name, values in features:
        added.append(
            _add_feature(
                dataframe,
                base_name,
                values,
            )
        )

    return added


# ============================================================
# GENERIC NUMERIC FEATURES
# ============================================================

def _add_numeric_row_features(
    dataframe: pd.DataFrame,
    numeric_columns: list[Any],
) -> list[str]:
    """
    Add a structurally safe row-level numeric completeness feature.

    The feature counts how many discovered numeric fields contain
    non-missing values. No numeric values from different fields are
    combined because their units or meanings are unknown.
    """

    if not numeric_columns:
        return []

    numeric = dataframe[numeric_columns].apply(
        pd.to_numeric,
        errors="coerce",
    )

    return [
        _add_feature(
            dataframe,
            "__numeric_non_null_count",
            numeric.notna().sum(axis=1).astype("Int64"),
        )
    ]


# ============================================================
# GENERIC MISSINGNESS FEATURES
# ============================================================

def _add_missingness_features(
    dataframe: pd.DataFrame,
) -> list[str]:
    """
    Add row-level structural missingness measures.

    These features describe data completeness. They do not impute or alter
    the source values.
    """

    if dataframe.empty or len(dataframe.columns) == 0:
        return []

    added: list[str] = []

    added.append(
        _add_feature(
            dataframe,
            "__missing_value_count",
            dataframe.isna().sum(axis=1).astype("Int64"),
        )
    )

    added.append(
        _add_feature(
            dataframe,
            "__non_missing_value_count",
            dataframe.notna().sum(axis=1).astype("Int64"),
        )
    )

    denominator = len(dataframe.columns)

    added.append(
        _add_feature(
            dataframe,
            "__missing_value_ratio",
            dataframe.isna().sum(axis=1) / denominator,
        )
    )

    return added


# ============================================================
# TABLE PREPARATION
# ============================================================

def prepare_table(
    table_name: str,
    dataframe: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """
    Prepare one discovered table.

    No table-name vocabulary is used to select transformations.
    """

    df = dataframe.copy()

    source_columns = list(df.columns)
    source_rows = len(df)

    column_kinds = _discover_column_kinds(df)

    temporal_columns = [
        column
        for column, kind in column_kinds.items()
        if kind == "temporal"
    ]

    numeric_columns = [
        column
        for column, kind in column_kinds.items()
        if kind == "numeric"
    ]

    boolean_columns = [
        column
        for column, kind in column_kinds.items()
        if kind == "boolean"
    ]

    text_columns = [
        column
        for column, kind in column_kinds.items()
        if kind == "text"
    ]

    added_features: list[str] = []

    # Feature lineage is recorded alongside the prepared-table statistics so
    # downstream analytical stages can distinguish source fields from
    # technical features created by this stage.
    feature_provenance: dict[str, dict[str, Any]] = {
        str(column): {
            "origin": "source",
            "feature_type": "source",
            "source_columns": [str(column)],
            "demand_candidate_eligible": True,
        }
        for column in source_columns
    }

    # --------------------------------------------------------
    # Calendar features for every discovered temporal field.
    # --------------------------------------------------------

    temporal_features: list[str] = []

    for column in temporal_columns:
        temporal_features.extend(
            _add_temporal_features(
                df,
                column,
            )
        )

    added_features.extend(temporal_features)

    for feature in temporal_features:
        feature_provenance[str(feature)] = {
            "origin": "derived",
            "feature_type": "derived_temporal",
            "source_columns": [str(column) for column in temporal_columns],
            "demand_candidate_eligible": False,
        }

    # --------------------------------------------------------
    # Generic row-level numerical structure.
    # --------------------------------------------------------

    numeric_features = _add_numeric_row_features(
        df,
        numeric_columns,
    )

    added_features.extend(numeric_features)

    for feature in numeric_features:
        feature_provenance[str(feature)] = {
            "origin": "derived",
            "feature_type": "derived_numeric",
            "source_columns": [str(column) for column in numeric_columns],
            "demand_candidate_eligible": False,
        }

    # --------------------------------------------------------
    # Generic row-level completeness structure.
    # --------------------------------------------------------

    missingness_features = _add_missingness_features(
        df
    )

    added_features.extend(missingness_features)

    for feature in missingness_features:
        feature_provenance[str(feature)] = {
            "origin": "derived",
            "feature_type": "derived_missingness",
            "source_columns": [str(column) for column in source_columns],
            "demand_candidate_eligible": False,
        }

    # --------------------------------------------------------
    # Structural protection.
    # --------------------------------------------------------

    if len(df) != source_rows:
        raise RuntimeError(
            f"Row count changed unexpectedly in table {table_name!r}."
        )

    for column in source_columns:
        if column not in df.columns:
            raise RuntimeError(
                f"Source column {column!r} disappeared from table {table_name!r}."
            )

    statistics = {
        "table": table_name,
        "source_rows": source_rows,
        "final_rows": len(df),
        "source_columns": len(source_columns),
        "final_columns": len(df.columns),
        "added_feature_count": len(added_features),
        "added_features": added_features,
        "structural_column_kinds": column_kinds,
        "temporal_columns": temporal_columns,
        "numeric_columns": numeric_columns,
        "boolean_columns": boolean_columns,
        "text_columns": text_columns,
        "feature_provenance": feature_provenance,
        "provenance_contract_version": 1,
    }

    return df, statistics


# ============================================================
# DATASET PREPARATION
# ============================================================

def prepare_dataset(
    tables: Mapping[str, pd.DataFrame],
) -> tuple[dict[str, pd.DataFrame], list[dict[str, Any]]]:
    """Prepare all discovered tables independently."""

    prepared: dict[str, pd.DataFrame] = {}
    statistics: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        cleaned, stats = prepare_table(
            str(table_name),
            dataframe,
        )

        prepared[str(table_name)] = cleaned
        statistics.append(stats)

    return prepared, statistics


# ============================================================
# STRUCTURAL VALIDATION
# ============================================================

def validate_prepared_dataset(
    original_tables: Mapping[str, pd.DataFrame],
    prepared_tables: Mapping[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Confirm that preprocessing did not remove or alter source structure.

    Derived columns are allowed. Source tables and source columns are not
    removed. Row counts remain unchanged because this module performs no
    aggregation or filtering.
    """

    findings: list[dict[str, Any]] = []

    original_names = set(
        str(name)
        for name in original_tables
    )

    prepared_names = set(
        str(name)
        for name in prepared_tables
    )

    missing_tables = original_names - prepared_names

    unexpected_tables = prepared_names - original_names

    for table_name in sorted(missing_tables):
        findings.append({
            "severity": "ERROR",
            "table": table_name,
            "issue": "Source table disappeared.",
        })

    for table_name in sorted(unexpected_tables):
        findings.append({
            "severity": "ERROR",
            "table": table_name,
            "issue": "Unexpected table appeared.",
        })

    for table_name in sorted(original_names & prepared_names):
        original = original_tables[table_name]
        prepared = prepared_tables[table_name]

        if len(original) != len(prepared):
            findings.append({
                "severity": "ERROR",
                "table": table_name,
                "issue": (
                    "Row count changed during preprocessing."
                ),
                "original_rows": len(original),
                "prepared_rows": len(prepared),
            })

        for column in original.columns:
            if column not in prepared.columns:
                findings.append({
                    "severity": "ERROR",
                    "table": table_name,
                    "column": str(column),
                    "issue": "Source column disappeared.",
                })

    return findings


# ============================================================
# JSON OUTPUT
# ============================================================

def write_json_dataset(
    tables: Mapping[str, pd.DataFrame],
    path: Path,
) -> None:
    """Write analytical tables to a generic JSON representation."""

    payload: dict[str, list[dict[str, Any]]] = {}

    for table_name, dataframe in tables.items():
        payload[str(table_name)] = [
            {
                str(key): _json_safe(value)
                for key, value in record.items()
            }
            for record in dataframe.to_dict(
                orient="records"
            )
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


# ============================================================
# EXCEL OUTPUT
# ============================================================

def _safe_sheet_name(
    table_name: str,
    used: set[str],
) -> str:
    """Create a unique Excel worksheet name without business assumptions."""

    base = str(table_name).strip()

    base = re.sub(
        r"[:\\/?*\[\]]",
        "_",
        base,
    )

    base = base[:31] or "Table"

    candidate = base
    counter = 2

    while candidate in used:
        suffix = f"_{counter}"
        candidate = (
            base[: 31 - len(suffix)]
            + suffix
        )
        counter += 1

    used.add(candidate)
    return candidate


def write_excel_dataset(
    tables: Mapping[str, pd.DataFrame],
    path: Path,
) -> None:
    """Write analytical tables to a separate Excel workbook."""

    temporary = path.with_suffix(
        ".tmp.xlsx"
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
# REPORT
# ============================================================

def build_report(
    source: Path,
    statistics: list[dict[str, Any]],
    findings: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "status": (
            "FAILED"
            if any(
                finding.get("severity") == "ERROR"
                for finding in findings
            )
            else "COMPLETED"
        ),
        "source": str(source),
        "tables_processed": len(statistics),
        "total_source_rows": sum(
            item["source_rows"]
            for item in statistics
        ),
        "total_final_rows": sum(
            item["final_rows"]
            for item in statistics
        ),
        "total_source_columns": sum(
            item["source_columns"]
            for item in statistics
        ),
        "total_final_columns": sum(
            item["final_columns"]
            for item in statistics
        ),
        "total_derived_features": sum(
            item["added_feature_count"]
            for item in statistics
        ),
        "tables": statistics,
        "findings": findings,
        "design_guarantees": {
            "source_columns_preserved": True,
            "source_rows_preserved": True,
            "table_count_not_fixed": True,
            "business_schema_not_required": True,
            "business_aliases_not_used": True,
            "dataset_name_not_required": True,
            "demand_target_not_inferred": True,
            "missing_values_not_imputed": True,
            "outliers_not_removed": True,
            "aggregation_not_performed": True,
            "source_data_not_overwritten": True,
        },
    }


def write_report(
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
            _json_safe(dict(report)),
            handle,
            ensure_ascii=False,
            indent=2,
        )

    temporary.replace(path)


# ============================================================
# MAIN
# ============================================================

def run_preprocessing_pipeline() -> dict[str, Any]:
    print_header(
        "ENTERPRISE DEMAND INTELLIGENCE PLATFORM\n"
        "GENERIC ANALYTICAL DATA PREPARATION"
    )

    print()
    print("Architecture mode:")
    print("  DATASET-DRIVEN")
    print()
    print("Transformation mode:")
    print("  STRUCTURAL AND ROW-LEVEL")
    print()
    print("Business schema required:")
    print("  NO")

    # --------------------------------------------------------
    # STEP 1 — Load processed dataset
    # --------------------------------------------------------

    print_section(
        "STEP 1 — LOADING PROCESSED DATASET"
    )

    source, original_tables = (
        load_processed_dataset()
    )

    print(
        f"Source: {source}"
    )

    print(
        f"Tables discovered: {len(original_tables):,}"
    )

    for table_name, dataframe in original_tables.items():
        print(
            f"  {table_name:<35}"
            f"{len(dataframe):>10,} rows"
            f"{len(dataframe.columns):>8} columns"
        )

    # --------------------------------------------------------
    # STEP 2 — Prepare
    # --------------------------------------------------------

    print_section(
        "STEP 2 — STRUCTURAL ANALYTICAL PREPARATION"
    )

    analytical_tables, statistics = (
        prepare_dataset(
            original_tables
        )
    )

    for item in statistics:
        print()
        print(
            f"Prepared: {item['table']}"
        )
        print(
            f"  Rows: "
            f"{item['source_rows']:,} -> "
            f"{item['final_rows']:,}"
        )
        print(
            f"  Columns: "
            f"{item['source_columns']:,} -> "
            f"{item['final_columns']:,}"
        )
        print(
            f"  Derived features: "
            f"{item['added_feature_count']:,}"
        )

    # --------------------------------------------------------
    # STEP 3 — Validate
    # --------------------------------------------------------

    print_section(
        "STEP 3 — STRUCTURAL INTEGRITY VALIDATION"
    )

    findings = validate_prepared_dataset(
        original_tables,
        analytical_tables,
    )

    if findings:
        for finding in findings:
            print(
                f"[{finding.get('severity')}] "
                f"{finding.get('table')}: "
                f"{finding.get('issue')}"
            )

        raise RuntimeError(
            "Analytical preprocessing structural validation failed."
        )

    print("Table preservation: PASSED.")
    print("Row-count preservation: PASSED.")
    print("Source-column preservation: PASSED.")

    # --------------------------------------------------------
    # STEP 4 — Report
    # --------------------------------------------------------

    report = build_report(
        source,
        statistics,
        findings,
    )

    # --------------------------------------------------------
    # STEP 5 — Write outputs
    # --------------------------------------------------------

    print_section(
        "STEP 4 — WRITING ANALYTICAL DATASET"
    )

    print(
        f"Writing JSON: {ANALYTICAL_JSON_PATH}"
    )

    write_json_dataset(
        analytical_tables,
        ANALYTICAL_JSON_PATH,
    )

    print("Analytical JSON: CREATED.")

    print()
    print(
        f"Writing Excel: {ANALYTICAL_EXCEL_PATH}"
    )

    write_excel_dataset(
        analytical_tables,
        ANALYTICAL_EXCEL_PATH,
    )

    print("Analytical Excel: CREATED.")

    write_report(
        report,
        PREPROCESSING_REPORT_PATH,
    )

    # --------------------------------------------------------
    # Final
    # --------------------------------------------------------

    print()
    print(SEPARATOR)
    print(
        "ANALYTICAL DATA PREPARATION "
        "COMPLETED SUCCESSFULLY"
    )
    print(SEPARATOR)

    print()
    print(
        f"Tables processed: "
        f"{report['tables_processed']:,}"
    )

    print(
        f"Total rows: "
        f"{report['total_final_rows']:,}"
    )

    print(
        f"Derived features: "
        f"{report['total_derived_features']:,}"
    )

    print()
    print("Source protection:")
    print("  Raw source modified: NO")
    print("  Processed source modified: NO")
    print("  Source rows removed: NO")
    print("  Source columns removed: NO")
    print("  Aggregation performed: NO")
    print("  Missing values imputed: NO")
    print("  Outliers removed: NO")
    print("  Demand target inferred: NO")
    print("  Business aliases used: NO")

    print()
    print("Outputs:")
    print(
        f"  JSON: {ANALYTICAL_JSON_PATH}"
    )
    print(
        f"  Excel: {ANALYTICAL_EXCEL_PATH}"
    )
    print(
        f"  Report: {PREPROCESSING_REPORT_PATH}"
    )

    return report


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    try:
        run_preprocessing_pipeline()

    except Exception as error:
        print()
        print(SEPARATOR)
        print(
            "ANALYTICAL DATA PREPARATION FAILED"
        )
        print(SEPARATOR)
        print()
        print(
            f"ERROR: {error}"
        )
        print()
        print(
            "Raw and processed source datasets were not overwritten."
        )
        raise
