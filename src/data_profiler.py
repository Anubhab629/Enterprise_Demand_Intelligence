"""
Enterprise Demand Intelligence Platform
GENERIC DATA PROFILER

Purpose
-------
Profile the dataset discovered by data_loader.py.

This module is read-only. It does not:
    - modify source data
    - clean values
    - remove records
    - overwrite source files
    - infer a business schema from names
    - create recommendations

The profiler works from the structures actually discovered in the
dataset. It does not require predefined table names, column names,
entity names, target names, or entity counts.

The module can still profile Excel workbooks directly through
load_workbook_sheets() for compatibility with existing workflows.

VS Code
-------
Running this file directly uses the same automatic dataset resolution
as data_loader.py, so the VS Code Run button works without arguments.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pandas as pd


# ============================================================
# PROJECT ROOT
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


# ============================================================
# PROJECT MODULES
# ============================================================

from config import (
    MAX_DATASET_FILE_SIZE_MB,
    SAMPLE_SIZE,
)

from src import data_loader


# ============================================================
# PROFILING LIMITS
# ============================================================

MAX_UNIQUE_VALUES_TO_DISPLAY = 15


# ============================================================
# GENERIC VALUE HELPERS
# ============================================================

def is_missing(value: Any) -> bool:
    if value is None:
        return True

    if isinstance(value, str):
        return value.strip().lower() in {
            "",
            "nan",
            "nat",
            "none",
            "null",
            "na",
            "n/a",
        }

    try:
        result = pd.isna(value)

        if hasattr(result, "__len__"):
            return False

        return bool(result)

    except (TypeError, ValueError):
        return False


def safe_value(value: Any) -> Any:
    if is_missing(value):
        return None

    if isinstance(value, (str, int, bool)):
        return value

    if isinstance(value, float):
        return value if pd.notna(value) else None

    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except Exception:
            pass

    return str(value)


def evenly_spaced_positions(
    length: int,
    sample_size: int,
) -> list[int]:
    """
    Return deterministic positions distributed across a sequence.

    This avoids profiling only the first records of a large dataset.
    """

    if length <= 0:
        return []

    if sample_size <= 0:
        raise ValueError(
            "sample_size must be greater than zero."
        )

    if length <= sample_size:
        return list(range(length))

    if sample_size == 1:
        return [0]

    step = (length - 1) / (sample_size - 1)

    return sorted({
        round(index * step)
        for index in range(sample_size)
    })


def sample_series(
    series: pd.Series,
    sample_size: int,
) -> pd.Series:
    if sample_size <= 0:
        raise ValueError(
            "sample_size must be greater than zero."
        )

    if len(series) <= sample_size:
        return series

    positions = evenly_spaced_positions(
        len(series),
        sample_size,
    )

    return series.iloc[positions]


# ============================================================
# EXCEL COMPATIBILITY LOADING
# ============================================================

def load_workbook_sheets(
    workbook_path: Path,
) -> dict[str, pd.DataFrame]:
    """
    Load every worksheet from an Excel workbook.

    This function is retained for existing Excel-based workflows.
    It does not impose any expected worksheet or column names.
    """

    workbook_path = Path(
        workbook_path
    ).expanduser().resolve()

    if not workbook_path.exists():
        raise FileNotFoundError(
            f"Excel workbook not found: {workbook_path}"
        )

    if workbook_path.suffix.lower() not in {
        ".xlsx",
        ".xls",
    }:
        raise ValueError(
            "load_workbook_sheets() requires an .xlsx or .xls file."
        )

    engine = (
        "openpyxl"
        if workbook_path.suffix.lower() == ".xlsx"
        else None
    )

    excel_file = pd.ExcelFile(
        workbook_path,
        engine=engine,
    )

    sheets: dict[str, pd.DataFrame] = {}

    for sheet_name in excel_file.sheet_names:
        sheets[str(sheet_name)] = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
        )

    return sheets


# ============================================================
# BASIC PROFILE
# ============================================================

def profile_basic_structure(
    dataframe: pd.DataFrame,
) -> dict[str, Any]:
    memory_bytes = int(
        dataframe.memory_usage(
            deep=True
        ).sum()
    )

    return {
        "rows": int(len(dataframe)),
        "columns": int(len(dataframe.columns)),
        "memory_bytes": memory_bytes,
        "memory_mb": round(
            memory_bytes / (1024 ** 2),
            3,
        ),
    }


# ============================================================
# DATA TYPE PROFILE
# ============================================================

def profile_data_types(
    dataframe: pd.DataFrame,
) -> dict[str, str]:
    return {
        str(column): str(dtype)
        for column, dtype in dataframe.dtypes.items()
    }


# ============================================================
# NULL / BLANK PROFILE
# ============================================================

def profile_missing_values(
    dataframe: pd.DataFrame,
) -> dict[str, dict[str, float | int]]:
    """
    Profile both true missing values and blank strings.

    The profiler does not decide whether either condition should
    be cleaned. It only reports the evidence.
    """

    results: dict[
        str,
        dict[str, float | int]
    ] = {}

    row_count = len(dataframe)

    for column in dataframe.columns:
        series = dataframe[column]

        missing_count = int(
            series.isna().sum()
        )

        blank_count = 0

        if (
            pd.api.types.is_object_dtype(series)
            or pd.api.types.is_string_dtype(series)
        ):
            text = series.dropna().astype(str)

            blank_count = int(
                text.str.strip().eq("").sum()
            )

        total_missing_like = (
            missing_count + blank_count
        )

        if total_missing_like == 0:
            continue

        results[str(column)] = {
            "missing_count": missing_count,
            "blank_string_count": blank_count,
            "missing_or_blank_count": total_missing_like,
            "missing_or_blank_percentage": round(
                (
                    total_missing_like
                    / row_count
                    * 100
                )
                if row_count
                else 0.0,
                2,
            ),
        }

    return results


# ============================================================
# DUPLICATE PROFILE
# ============================================================

def profile_duplicate_rows(
    dataframe: pd.DataFrame,
) -> dict[str, Any]:
    duplicate_mask = dataframe.duplicated(
        keep=False
    )

    return {
        "duplicate_row_count": int(
            duplicate_mask.sum()
        ),
        "duplicate_group_count": int(
            dataframe[
                duplicate_mask
            ].drop_duplicates().shape[0]
        ),
        "has_duplicates": bool(
            duplicate_mask.any()
        ),
    }


# ============================================================
# UNIQUE VALUE PROFILE
# ============================================================

def profile_unique_values(
    dataframe: pd.DataFrame,
) -> dict[str, dict[str, Any]]:
    results = {}

    for column in dataframe.columns:
        series = dataframe[column]

        non_missing = series.dropna()

        unique_count = int(
            non_missing.nunique(
                dropna=True
            )
        )

        row_count = len(series)

        results[str(column)] = {
            "unique_count": unique_count,
            "unique_ratio": round(
                (
                    unique_count
                    / len(non_missing)
                )
                if len(non_missing)
                else 0.0,
                6,
            ),
            "all_values_unique": bool(
                len(non_missing) > 0
                and unique_count == len(non_missing)
            ),
            "constant_non_missing": bool(
                unique_count == 1
            ),
        }

    return results


# ============================================================
# STRUCTURAL KEY CANDIDATE DETECTION
# ============================================================

def identify_candidate_keys(
    dataframe: pd.DataFrame,
) -> list[Any]:
    """
    Identify structurally plausible key columns.

    IMPORTANT:
    No name-based business aliases are used.

    Evidence includes:
        - high uniqueness
        - low missingness
        - data type
        - repeated structured-value patterns
        - integer-like numeric values

    This remains a candidate detector. It does not declare a key.
    """

    candidates: list[str] = []

    for column in dataframe.columns:
        series = dataframe[column]
        non_missing = series.dropna()

        if non_missing.empty:
            continue

        unique_count = int(
            non_missing.nunique()
        )

        unique_ratio = (
            unique_count
            / len(non_missing)
        )

        missing_ratio = (
            series.isna().mean()
        )

        if unique_ratio < 0.98:
            continue

        if missing_ratio > 0.01:
            continue

        inferred_type = data_loader.infer_column_type(
            series
        )

        if inferred_type == "temporal":
            continue

        score = 0.0

        # Strong uniqueness evidence.
        score += min(
            unique_ratio,
            1.0,
        ) * 0.45

        # No missing values.
        if missing_ratio == 0:
            score += 0.15

        if inferred_type == "text":
            text = non_missing.astype(str).str.strip()

            structured_ratio = float(
                text.str.contains(
                    r"[\-_:/]|[A-Za-z].*\d|\d.*[A-Za-z]",
                    regex=True,
                ).mean()
            )

            score += (
                structured_ratio * 0.30
            )

        elif inferred_type == "numeric":
            numeric = pd.to_numeric(
                non_missing,
                errors="coerce",
            )

            if numeric.notna().all():
                integer_ratio = float(
                    numeric.mod(1).eq(0).mean()
                )

                score += (
                    integer_ratio * 0.20
                )

        # Small tables can have accidental uniqueness. Do not
        # treat an arbitrary categorical/text field as a key just
        # because every observed value happens to be different.
        if (
            inferred_type == "categorical"
            and len(non_missing) < 100
        ):
            score *= 0.50

        if score >= 0.75:
            candidates.append(column)

    return candidates


# ============================================================
# KEY QUALITY PROFILE
# ============================================================

def profile_candidate_keys(
    dataframe: pd.DataFrame,
    candidates: list[Any],
) -> dict[Any, dict[str, Any]]:
    results = {}

    for column in candidates:
        series = dataframe[column]

        missing_count = int(
            series.isna().sum()
        )

        non_missing = series.dropna()

        unique_count = int(
            non_missing.nunique(
                dropna=True
            )
        )

        duplicate_value_count = int(
            non_missing.duplicated(
                keep=False
            ).sum()
        )

        row_count = len(series)

        results[column] = {
            "rows": int(row_count),
            "unique_values": unique_count,
            "missing_values": missing_count,
            "duplicate_values": duplicate_value_count,
            "uniqueness_ratio": round(
                (
                    unique_count
                    / len(non_missing)
                )
                if len(non_missing)
                else 0.0,
                6,
            ),
            "appears_unique": bool(
                row_count > 0
                and unique_count == row_count
                and missing_count == 0
            ),
        }

    return results


# ============================================================
# NUMERIC PROFILE
# ============================================================

def profile_numeric_columns(
    dataframe: pd.DataFrame,
) -> dict[str, dict[str, Any]]:
    results = {}

    numeric_columns = dataframe.select_dtypes(
        include="number"
    ).columns

    for column in numeric_columns:
        series = pd.to_numeric(
            dataframe[column],
            errors="coerce",
        ).dropna()

        if series.empty:
            continue

        finite = series[
            series.map(
                lambda value: pd.notna(value)
                and float(value) not in {
                    float("inf"),
                    float("-inf"),
                }
            )
        ]

        if finite.empty:
            continue

        results[str(column)] = {
            "minimum": safe_value(
                finite.min()
            ),
            "maximum": safe_value(
                finite.max()
            ),
            "mean": safe_value(
                finite.mean()
            ),
            "median": safe_value(
                finite.median()
            ),
            "zero_count": int(
                (finite == 0).sum()
            ),
            "negative_count": int(
                (finite < 0).sum()
            ),
            "positive_count": int(
                (finite > 0).sum()
            ),
            "finite_count": int(
                len(finite)
            ),
            "integer_like": bool(
                finite.mod(1).eq(0).all()
            ),
        }

    return results


# ============================================================
# TEXT PROFILE
# ============================================================

def profile_text_columns(
    dataframe: pd.DataFrame,
) -> dict[str, dict[str, Any]]:
    results = {}

    text_columns = dataframe.select_dtypes(
        include=["object", "string", "category"]
    ).columns

    for column in text_columns:
        series = dataframe[column].dropna()

        if series.empty:
            continue

        text_series = series.astype(str)

        blank_count = int(
            text_series.str.strip().eq("").sum()
        )

        whitespace_count = int(
            (
                text_series
                != text_series.str.strip()
            ).sum()
        )

        sample_values = [
            safe_value(value)
            for value in (
                text_series
                .drop_duplicates()
                .head(
                    MAX_UNIQUE_VALUES_TO_DISPLAY
                )
                .tolist()
            )
        ]

        unique_count = int(
            text_series.nunique()
        )

        results[str(column)] = {
            "unique_values": unique_count,
            "unique_ratio": round(
                unique_count
                / len(text_series),
                6,
            ),
            "blank_strings": blank_count,
            "leading_or_trailing_whitespace":
                whitespace_count,
            "sample_values": sample_values,
        }

    return results


# ============================================================
# TEMPORAL PROFILE
# ============================================================

def profile_date_like_columns(
    dataframe: pd.DataFrame,
) -> dict[str, dict[str, Any]]:
    """
    Profile columns that data_loader has structurally identified
    as temporal.

    No column-name vocabulary is used.
    """

    results = {}

    for column in dataframe.columns:
        inferred_type = (
            data_loader.infer_column_type(
                dataframe[column]
            )
        )

        if inferred_type != "temporal":
            continue

        series = dataframe[column].dropna()

        if series.empty:
            continue

        converted = pd.to_datetime(
            series,
            errors="coerce",
            format="mixed",
        )

        valid = converted.dropna()

        result: dict[str, Any] = {
            "values_checked": int(len(series)),
            "valid_datetime_values": int(
                len(valid)
            ),
            "invalid_datetime_values": int(
                converted.isna().sum()
            ),
            "unique_datetime_values": int(
                valid.nunique()
            ),
        }

        if not valid.empty:
            ordered = valid.sort_values()

            result["minimum"] = safe_value(
                ordered.iloc[0]
            )

            result["maximum"] = safe_value(
                ordered.iloc[-1]
            )

            result["monotonic_in_input"] = bool(
                valid.is_monotonic_increasing
            )

            differences = (
                ordered.diff()
                .dropna()
            )

            if not differences.empty:
                result[
                    "median_interval_seconds"
                ] = safe_value(
                    differences.dt.total_seconds()
                    .median()
                )

        results[str(column)] = result

    return results


# ============================================================
# DISTRIBUTION / CARDINALITY PROFILE
# ============================================================

def profile_column_cardinality(
    dataframe: pd.DataFrame,
) -> dict[str, dict[str, Any]]:
    results = {}

    row_count = len(dataframe)

    for column in dataframe.columns:
        series = dataframe[column]

        non_missing = series.dropna()

        unique_count = int(
            non_missing.nunique()
        )

        results[str(column)] = {
            "row_count": int(row_count),
            "non_missing_count": int(
                len(non_missing)
            ),
            "unique_count": unique_count,
            "cardinality_ratio": round(
                (
                    unique_count
                    / len(non_missing)
                )
                if len(non_missing)
                else 0.0,
                6,
            ),
        }

    return results


# ============================================================
# TABLE PROFILE
# ============================================================

def profile_table(
    table_name: str,
    dataframe: pd.DataFrame,
    sample_size: int = SAMPLE_SIZE,
) -> dict[str, Any]:
    """
    Generate a complete read-only profile for one table.
    """

    sampled = data_loader.sample_frame(
        dataframe,
        sample_size,
    )

    candidate_keys = identify_candidate_keys(
        sampled
    )

    temporal_profile = profile_date_like_columns(
        dataframe
    )

    return {
        "table": str(table_name),
        "basic": profile_basic_structure(
            dataframe
        ),
        "sample": {
            "requested_size": int(sample_size),
            "actual_size": int(len(sampled)),
            "sampling_strategy": (
                "deterministic_evenly_spaced"
            ),
        },
        "data_types": profile_data_types(
            dataframe
        ),
        "missing_values": profile_missing_values(
            dataframe
        ),
        "duplicate_rows": profile_duplicate_rows(
            dataframe
        ),
        "unique_values": profile_unique_values(
            dataframe
        ),
        "cardinality": profile_column_cardinality(
            dataframe
        ),
        "candidate_keys": candidate_keys,
        "identifier_candidates": candidate_keys,
        "candidate_key_quality":
            profile_candidate_keys(
                sampled,
                candidate_keys,
            ),
        "numeric_columns":
            profile_numeric_columns(
                dataframe
            ),
        "text_columns":
            profile_text_columns(
                dataframe
            ),
        "date_like_columns":
            temporal_profile,
        "temporal_columns":
            list(temporal_profile.keys()),
    }


# ============================================================
# WORKBOOK / DATASET PROFILE
# ============================================================

def profile_tables(
    tables: dict[str, pd.DataFrame],
    sample_size: int = SAMPLE_SIZE,
) -> list[dict[str, Any]]:
    reports = []

    for table_name, dataframe in tables.items():
        reports.append(
            profile_table(
                table_name,
                dataframe,
                sample_size=sample_size,
            )
        )

    return reports


def profile_workbook(
    workbook_path: Path,
    sample_size: int = SAMPLE_SIZE,
) -> list[dict[str, Any]]:
    """
    Profile an Excel workbook directly.

    Retained for compatibility with existing Excel workflows.
    """

    sheets = load_workbook_sheets(
        workbook_path
    )

    return profile_tables(
        sheets,
        sample_size=sample_size,
    )


def profile_dataset(
    dataset_path: str | Path | None = None,
    sample_size: int = SAMPLE_SIZE,
    max_size_mb: float = MAX_DATASET_FILE_SIZE_MB,
) -> dict[str, Any]:
    """
    Profile the same dynamically resolved dataset used by the loader.

    The loader's structural discovery result remains the canonical
    structural evidence. This module adds detailed read-only profiling
    without replacing that evidence.
    """

    if sample_size <= 0:
        raise ValueError(
            "sample_size must be greater than zero."
        )

    if max_size_mb <= 0:
        raise ValueError(
            "max_size_mb must be greater than zero."
        )

    discovery = data_loader.discover_dataset(
        dataset_path=dataset_path,
        sample_size=sample_size,
        max_size_mb=max_size_mb,
    )

    detailed_profiles = profile_tables(
        discovery["tables"],
        sample_size=sample_size,
    )

    return {
        "dataset": discovery["dataset"],
        "summary": discovery["summary"],
        "tables": detailed_profiles,
        "discovery_profiles": discovery["profiles"],
        "discovery_roles": discovery["roles"],
        "discovery_relationships": discovery["relationships"],
        "discovery_reuse_candidates": discovery[
            "reuse_candidates"
        ],
        "quality": discovery["quality"],
        "ml_risks": discovery["ml_risks"],
    }


# ============================================================
# REPORT SERIALIZATION
# ============================================================

def build_report(
    result: dict[str, Any],
) -> dict[str, Any]:
    return {
        "dataset": result["dataset"],
        "summary": result["summary"],
        "tables": result["tables"],
        "discovery_roles": result[
            "discovery_roles"
        ],
        "discovery_relationships": result[
            "discovery_relationships"
        ],
        "discovery_reuse_candidates": result[
            "discovery_reuse_candidates"
        ],
        "quality": result["quality"],
        "architecture": {
            "read_only": True,
            "business_aliases_used": False,
            "hardcoded_business_table_names": False,
            "hardcoded_business_column_names": False,
            "hardcoded_entity_counts": False,
            "hardcoded_target_column": False,
            "name_based_date_detection": False,
        },
    }


# ============================================================
# REPORT DISPLAY
# ============================================================

def print_table_profile(
    report: dict[str, Any],
) -> None:
    basic = report["basic"]

    print()
    print("-" * 78)
    print(
        f"TABLE: {report['table']}"
    )
    print("-" * 78)

    print(
        f"Rows: {basic['rows']:,}"
    )

    print(
        f"Columns: {basic['columns']}"
    )

    print(
        f"Memory: {basic['memory_mb']} MB"
    )

    duplicate = report[
        "duplicate_rows"
    ]

    print(
        "Duplicate rows: "
        f"{duplicate['duplicate_row_count']:,}"
    )

    missing = report[
        "missing_values"
    ]

    if missing:
        print()
        print("MISSING / BLANK VALUES:")

        for column, values in missing.items():
            print(
                f"  {column}: "
                f"{values['missing_or_blank_count']:,} "
                f"({values['missing_or_blank_percentage']}%)"
            )

    candidates = report[
        "candidate_keys"
    ]

    if candidates:
        print()
        print("STRUCTURAL KEY CANDIDATES:")

        for column in candidates:
            quality = report[
                "candidate_key_quality"
            ][column]

            print(
                f"  {column}: "
                f"unique={quality['unique_values']:,}, "
                f"missing={quality['missing_values']:,}, "
                f"duplicates={quality['duplicate_values']:,}, "
                f"appears_unique="
                f"{quality['appears_unique']}"
            )

    numeric = report[
        "numeric_columns"
    ]

    numeric_issues = {
        column: values
        for column, values in numeric.items()
        if (
            values["negative_count"] > 0
            or values["zero_count"] > 0
        )
    }

    if numeric_issues:
        print()
        print(
            "NUMERIC COLUMNS WITH "
            "ZERO OR NEGATIVE VALUES:"
        )

        for column, values in numeric_issues.items():
            print(
                f"  {column}: "
                f"zeros={values['zero_count']:,}, "
                f"negatives={values['negative_count']:,}, "
                f"min={values['minimum']}"
            )

    text = report[
        "text_columns"
    ]

    text_issues = {
        column: values
        for column, values in text.items()
        if (
            values["blank_strings"] > 0
            or values[
                "leading_or_trailing_whitespace"
            ] > 0
        )
    }

    if text_issues:
        print()
        print("TEXT QUALITY ISSUES:")

        for column, values in text_issues.items():
            print(
                f"  {column}: "
                f"blank={values['blank_strings']:,}, "
                f"whitespace="
                f"{values['leading_or_trailing_whitespace']:,}"
            )

    dates = report[
        "date_like_columns"
    ]

    date_issues = {
        column: values
        for column, values in dates.items()
        if values[
            "invalid_datetime_values"
        ] > 0
    }

    if date_issues:
        print()
        print(
            "TEMPORAL COLUMNS WITH "
            "INVALID VALUES:"
        )

        for column, values in date_issues.items():
            print(
                f"  {column}: "
                f"invalid="
                f"{values['invalid_datetime_values']:,}"
            )


def print_summary(
    result: dict[str, Any],
) -> None:
    dataset = result["dataset"]
    summary = result["summary"]

    print("=" * 78)
    print("GENERIC DATA PROFILING")
    print("=" * 78)
    print(
        f"Dataset : {dataset['path']}"
    )
    print(
        f"Format  : {dataset['format']}"
    )
    print(
        f"Size    : {dataset['size_mb']:.2f} MB"
    )
    print(
        f"Tables  : {summary['table_count']}"
    )
    print(
        f"Rows    : {summary['total_rows']:,}"
    )

    print()
    print(
        "DISCOVERED TABLE PROFILES"
    )

    for report in result["tables"]:
        print_table_profile(
            report
        )

    print()
    print("=" * 78)
    print("PROFILING SUMMARY")
    print("=" * 78)

    duplicate_rows = sum(
        report[
            "duplicate_rows"
        ][
            "duplicate_row_count"
        ]
        for report in result["tables"]
    )

    tables_with_missing = sum(
        bool(
            report["missing_values"]
        )
        for report in result["tables"]
    )

    print(
        f"Tables profiled: "
        f"{len(result['tables']):,}"
    )

    print(
        f"Total rows: "
        f"{summary['total_rows']:,}"
    )

    print(
        f"Duplicate rows: "
        f"{duplicate_rows:,}"
    )

    print(
        f"Tables containing missing/blank values: "
        f"{tables_with_missing:,}"
    )

    print()
    print(
        "Read-only profiling completed."
    )


# ============================================================
# CLI
# ============================================================

def parse_args() -> Any:
    import argparse

    parser = argparse.ArgumentParser(
        description=(
            "Profile a dynamically discovered dataset."
        )
    )

    parser.add_argument(
        "dataset",
        nargs="?",
        help=(
            "Optional dataset path. If omitted, "
            "the same automatic resolution used by "
            "data_loader.py is applied."
        ),
    )

    parser.add_argument(
        "--sample-size",
        type=int,
        default=SAMPLE_SIZE,
    )

    parser.add_argument(
        "--max-size-mb",
        type=float,
        default=MAX_DATASET_FILE_SIZE_MB,
    )

    parser.add_argument(
        "--report",
        help="Optional JSON report output path.",
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.sample_size <= 0:
        raise ValueError(
            "--sample-size must be greater than zero."
        )

    if args.max_size_mb <= 0:
        raise ValueError(
            "--max-size-mb must be greater than zero."
        )

    result = profile_dataset(
        dataset_path=args.dataset,
        sample_size=args.sample_size,
        max_size_mb=args.max_size_mb,
    )

    print_summary(
        result
    )

    if args.report:
        report_path = (
            Path(args.report)
            .expanduser()
            .resolve()
        )

        report_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        report_path.write_text(
            json.dumps(
                build_report(result),
                indent=2,
                ensure_ascii=False,
                default=str,
            ),
            encoding="utf-8",
        )

        print(
            f"Report written to: {report_path}"
        )

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(
            main()
        )
    except Exception as exc:
        print(
            f"ERROR: {exc}",
            file=sys.stderr,
        )
        raise SystemExit(1)
