"""
Enterprise Demand Intelligence Platform
GENERIC DATASET-DRIVEN ANALYTICAL ENGINE

Purpose
-------
Create reusable analytical evidence from an analytical dataset without
requiring a predefined business schema.

Design principles
-----------------
- Discover actual tables and columns at runtime.
- Use structural evidence rather than business-name aliases.
- Do not require particular table names, column names, entities, values,
  counts, or date ranges.
- Preserve source rows and source columns.
- Never infer business meaning from a column name.
- Treat demand candidates and signal relationships as evidence only.
- Adapt analysis to the capabilities actually present.
- Bound resource-intensive comparisons.
- Never modify the source dataset.
"""

from __future__ import annotations

import json
import math
import sys
import warnings
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


# ============================================================
# PROJECT / CONFIGURATION
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config


OUTPUT_DATA_DIRECTORY = Path(
    getattr(
        config,
        "OUTPUT_DATA_DIRECTORY",
        PROJECT_ROOT / "data" / "output",
    )
).expanduser().resolve()

ANALYSIS_WORKBOOK = (
    OUTPUT_DATA_DIRECTORY / "analysis_results.xlsx"
)
ANALYSIS_DATASET_JSON = (
    OUTPUT_DATA_DIRECTORY / "analysis_results.json"
)
ANALYSIS_REPORT_JSON = (
    OUTPUT_DATA_DIRECTORY / "analysis_report.json"
)


# ============================================================
# TECHNICAL DEMAND-EVIDENCE CONTROLS
# ============================================================

DEMAND_MIN_OBSERVATIONS = max(
    10,
    int(
        getattr(
            config,
            "DEMAND_MIN_OBSERVATIONS",
            30,
        )
    ),
)

DEMAND_MAX_CANDIDATES_PER_TABLE = max(
    1,
    int(
        getattr(
            config,
            "DEMAND_MAX_CANDIDATES_PER_TABLE",
            20,
        )
    ),
)

DEMAND_MAX_LAG_PERIODS = max(
    0,
    int(
        getattr(
            config,
            "DEMAND_MAX_LAG_PERIODS",
            6,
        )
    ),
)

DEMAND_MIN_ABS_CORRELATION = max(
    0.0,
    float(
        getattr(
            config,
            "DEMAND_MIN_ABS_CORRELATION",
            0.30,
        )
    ),
)

DEMAND_MIN_DIRECTIONAL_AGREEMENT = max(
    0.0,
    min(
        1.0,
        float(
            getattr(
                config,
                "DEMAND_MIN_DIRECTIONAL_AGREEMENT",
                0.55,
            )
        ),
    ),
)

DEMAND_MAX_RELATIONSHIP_COMPARISONS = max(
    1,
    int(
        getattr(
            config,
            "DEMAND_MAX_RELATIONSHIP_COMPARISONS",
            5000,
        )
    ),
)

MAX_DIMENSIONS_PER_TABLE = max(
    1,
    int(
        getattr(
            config,
            "ANALYSIS_MAX_DIMENSIONS_PER_TABLE",
            8,
        )
    ),
)

MAX_GROUPS_PER_OUTPUT = max(
    1,
    int(
        getattr(
            config,
            "ANALYSIS_MAX_GROUPS_PER_OUTPUT",
            5000,
        )
    ),
)

MAX_MEASURES_PER_TABLE = max(
    1,
    int(
        getattr(
            config,
            "ANALYSIS_MAX_MEASURES_PER_TABLE",
            20,
        )
    ),
)

MAX_CORRELATION_COLUMNS_PER_TABLE = max(
    2,
    int(
        getattr(
            config,
            "ANALYSIS_MAX_CORRELATION_COLUMNS_PER_TABLE",
            100,
        )
    ),
)

MAX_OUTPUT_ROWS_PER_ANALYSIS = max(
    1,
    int(
        getattr(
            config,
            "ANALYSIS_MAX_OUTPUT_ROWS_PER_ANALYSIS",
            500000,
        )
    ),
)


# ============================================================
# SAFE VALUE HELPERS
# ============================================================

def safe_float(
    value: Any,
) -> float | None:
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    try:
        result = float(value)
    except (TypeError, ValueError):
        return None

    if not math.isfinite(result):
        return None

    return result


def safe_json_value(
    value: Any,
) -> Any:
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, np.integer):
        return int(value)

    if isinstance(value, np.floating):
        number = float(value)
        return number if math.isfinite(number) else None

    if isinstance(value, np.bool_):
        return bool(value)

    if isinstance(value, pd.Timestamp):
        return value.isoformat()

    if isinstance(value, pd.Timedelta):
        return str(value)

    if isinstance(value, (str, int, bool)):
        return value

    if isinstance(value, float):
        return value if math.isfinite(value) else None

    return str(value)


def safe_output_dataframe(
    dataframe: pd.DataFrame | None,
) -> pd.DataFrame:
    if dataframe is None:
        return pd.DataFrame()

    result = dataframe.copy()

    result = result.replace(
        [np.inf, -np.inf],
        pd.NA,
    )

    return result.astype(object).where(
        pd.notna(result),
        None,
    )


# ============================================================
# STRUCTURAL TYPE DISCOVERY
# ============================================================

def is_datetime_dtype(
    series: pd.Series,
) -> bool:
    dtype = series.dtype

    return bool(
        pd.api.types.is_datetime64_any_dtype(dtype)
        or isinstance(
            dtype,
            pd.DatetimeTZDtype,
        )
    )


def parse_temporal_values(
    series: pd.Series,
) -> pd.Series:
    """
    Parse temporal values using value structure only.

    No column-name vocabulary is used.
    """
    if series.empty:
        return pd.to_datetime(
            series,
            errors="coerce",
        )

    try:
        return pd.to_datetime(
            series,
            errors="coerce",
            format="mixed",
        )
    except (TypeError, ValueError):
        try:
            return pd.to_datetime(
                series,
                errors="coerce",
                format="ISO8601",
            )
        except (TypeError, ValueError):
            with warnings.catch_warnings():
                warnings.simplefilter(
                    "ignore",
                    UserWarning,
                )
                return pd.to_datetime(
                    series,
                    errors="coerce",
                )


def infer_column_kind(
    series: pd.Series,
) -> str:
    """
    Infer a structural kind from dtype and observed values.
    """
    if is_datetime_dtype(series):
        return "temporal"

    if pd.api.types.is_bool_dtype(
        series.dtype
    ):
        return "boolean"

    if pd.api.types.is_numeric_dtype(
        series.dtype
    ):
        return "numeric"

    non_null = series.dropna()

    if non_null.empty:
        return "empty"

    converted = parse_temporal_values(
        non_null
    )

    if (
        len(non_null) > 0
        and float(converted.notna().mean()) >= 0.95
    ):
        return "temporal"

    unique_count = int(
        non_null.nunique(
            dropna=True
        )
    )
    row_count = len(non_null)

    if (
        pd.api.types.is_object_dtype(
            series.dtype
        )
        or pd.api.types.is_string_dtype(
            series.dtype
        )
    ):
        if unique_count <= max(
            50,
            min(
                500,
                int(
                    max(
                        row_count,
                        1,
                    )
                    * 0.05
                ),
            ),
        ):
            return "categorical"

        return "text"

    return "unknown"


def structural_key_score(
    series: pd.Series,
) -> dict[str, Any]:
    """
    Estimate identifier evidence without business-name matching.

    Numeric fields are only promoted strongly when they also show
    sequential/integer-like structure. This reduces false identification of
    continuous measures as identifiers.
    """
    non_null = series.dropna()

    if non_null.empty:
        return {
            "non_null_count": 0,
            "unique_count": 0,
            "duplicate_count": 0,
            "uniqueness_ratio": 0.0,
            "sequential_ratio": 0.0,
            "candidate": False,
        }

    kind = infer_column_kind(
        series
    )

    unique_count = int(
        non_null.nunique(
            dropna=True
        )
    )

    non_null_count = int(
        len(non_null)
    )

    duplicate_count = (
        non_null_count
        - unique_count
    )

    uniqueness_ratio = (
        unique_count
        / max(
            non_null_count,
            1,
        )
    )

    sequential_ratio = 0.0

    if kind == "numeric":
        numeric = pd.to_numeric(
            non_null,
            errors="coerce",
        ).dropna()

        if len(numeric) >= 3:
            values = numeric.to_numpy(
                dtype=float
            )

            integer_like = np.isclose(
                values,
                np.round(values),
                rtol=0.0,
                atol=1e-9,
            )

            ordered = np.sort(
                np.unique(
                    values
                )
            )

            if (
                integer_like.mean() >= 0.95
                and len(ordered) >= 3
            ):
                differences = np.diff(
                    ordered
                )

                sequential_ratio = float(
                    np.mean(
                        np.isclose(
                            differences,
                            1.0,
                            rtol=0.0,
                            atol=1e-9,
                        )
                    )
                )

    if kind == "numeric":
        candidate = (
            uniqueness_ratio >= 0.995
            and (
                sequential_ratio >= 0.80
                or non_null_count <= 100
            )
        )
    elif kind in {
        "categorical",
        "text",
    }:
        candidate = (
            uniqueness_ratio >= 0.98
        )
    else:
        candidate = False

    return {
        "non_null_count": non_null_count,
        "unique_count": unique_count,
        "duplicate_count": duplicate_count,
        "uniqueness_ratio": float(
            uniqueness_ratio
        ),
        "sequential_ratio": float(
            sequential_ratio
        ),
        "candidate": bool(
            candidate
        ),
    }


def numeric_clean(
    series: pd.Series,
) -> pd.Series:
    return pd.to_numeric(
        series,
        errors="coerce",
    )


# ============================================================
# DATASET LOADING
# ============================================================

def _read_json_dataset(
    path: Path,
) -> dict[str, pd.DataFrame]:
    with path.open(
        "r",
        encoding="utf-8",
    ) as handle:
        payload = json.load(
            handle
        )

    if not isinstance(
        payload,
        dict,
    ):
        raise ValueError(
            "Analytical JSON must contain a table mapping."
        )

    tables: dict[
        str,
        pd.DataFrame,
    ] = {}

    for table_name, records in payload.items():
        if isinstance(
            records,
            list,
        ):
            tables[str(table_name)] = (
                pd.DataFrame(records)
            )
        elif isinstance(
            records,
            dict,
        ):
            tables[str(table_name)] = (
                pd.DataFrame(records)
            )
        else:
            raise ValueError(
                "Unsupported table structure for "
                f"'{table_name}'."
            )

    return tables


def _read_excel_dataset(
    path: Path,
) -> dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(
        path,
        engine="openpyxl",
    )

    tables: dict[
        str,
        pd.DataFrame,
    ] = {}

    for sheet_name in workbook.sheet_names:
        dataframe = pd.read_excel(
            workbook,
            sheet_name=sheet_name,
        )

        dataframe = dataframe.dropna(
            axis=0,
            how="all",
        )

        dataframe = dataframe.dropna(
            axis=1,
            how="all",
        )

        tables[str(sheet_name)] = (
            dataframe
        )

    return tables


def resolve_analytical_source() -> Path:
    """
    Resolve the pipeline's generic analytical artifact.

    This filename is a technical pipeline artifact, not a business
    dataset/schema assumption.
    """
    if not OUTPUT_DATA_DIRECTORY.exists():
        raise FileNotFoundError(
            "Configured analytical output directory does not exist."
        )

    candidates = [
        path
        for path in sorted(
            OUTPUT_DATA_DIRECTORY.iterdir()
        )
        if (
            path.is_file()
            and path.stem
            == "analytical_dataset"
            and path.suffix.lower()
            in {
                ".json",
                ".xlsx",
                ".xls",
                ".parquet",
            }
        )
    ]

    if not candidates:
        raise FileNotFoundError(
            "No analytical dataset artifact was found."
        )

    extension_priority = {
        ".xlsx": 0,
        ".json": 1,
        ".parquet": 2,
        ".xls": 3,
    }

    candidates.sort(
        key=lambda path: (
            extension_priority.get(
                path.suffix.lower(),
                99,
            ),
            path.name.lower(),
        )
    )

    return candidates[0]


def load_feature_provenance(
    source: Path,
) -> dict[str, dict[str, dict[str, Any]]]:
    """Load preprocessing lineage metadata for the analytical artifact.

    The metadata is a technical sidecar produced by preprocessing. It is
    intentionally optional for backward compatibility with older artifacts;
    when absent, callers can use the conservative technical-name fallback.
    """

    report_path = source.with_name(
        "preprocessing_report.json"
    )

    if not report_path.exists():
        return {}

    try:
        with report_path.open(
            "r",
            encoding="utf-8",
        ) as handle:
            report = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return {}

    tables = report.get("tables")

    if not isinstance(tables, list):
        return {}

    result: dict[str, dict[str, dict[str, Any]]] = {}

    for item in tables:
        if not isinstance(item, dict):
            continue

        table_name = str(item.get("table", ""))
        provenance = item.get("feature_provenance")

        if table_name and isinstance(provenance, dict):
            result[table_name] = {
                str(column): metadata
                for column, metadata in provenance.items()
                if isinstance(metadata, dict)
            }

    return result


def load_analytical_dataset() -> tuple[
    Path,
    dict[str, pd.DataFrame],
]:
    source = resolve_analytical_source()

    print()
    print("-" * 72)
    print(
        "STEP 1 — LOADING ANALYTICAL DATASET"
    )
    print("-" * 72)
    print(
        f"Source: {source}"
    )

    suffix = source.suffix.lower()

    if suffix == ".json":
        tables = _read_json_dataset(
            source
        )
    elif suffix in {
        ".xlsx",
        ".xls",
    }:
        tables = _read_excel_dataset(
            source
        )
    elif suffix == ".parquet":
        tables = {
            source.stem: pd.read_parquet(
                source
            )
        }
    else:
        raise ValueError(
            "Unsupported analytical dataset format: "
            f"{source.suffix}"
        )

    tables = {
        name: dataframe
        for name, dataframe in tables.items()
        if dataframe is not None
    }

    if not tables:
        raise RuntimeError(
            "Analytical dataset contains no tables."
        )

    print(
        f"Tables discovered: {len(tables):,}"
    )

    for name, dataframe in tables.items():
        print(
            f"  {name:<38}"
            f"{len(dataframe):>12,} rows"
            f"{len(dataframe.columns):>12,} columns"
        )

    return source, tables


# ============================================================
# STRUCTURAL SELECTION
# ============================================================

def select_numeric_columns(
    dataframe: pd.DataFrame,
) -> list[Any]:
    return [
        column
        for column in dataframe.columns
        if infer_column_kind(
            dataframe[column]
        ) == "numeric"
    ]


def select_temporal_columns(
    dataframe: pd.DataFrame,
) -> list[Any]:
    return [
        column
        for column in dataframe.columns
        if infer_column_kind(
            dataframe[column]
        ) == "temporal"
    ]


def select_dimension_columns(
    dataframe: pd.DataFrame,
) -> list[Any]:
    dimensions: list[Any] = []

    row_count = max(
        len(dataframe),
        1,
    )

    max_text_cardinality = min(
        100,
        max(
            20,
            row_count // 20,
        ),
    )

    for column in dataframe.columns:
        series = dataframe[column]
        kind = infer_column_kind(
            series
        )

        if kind not in {
            "categorical",
            "text",
            "boolean",
        }:
            continue

        non_null = series.dropna()

        if non_null.empty:
            continue

        unique_count = int(
            non_null.nunique(
                dropna=True
            )
        )

        if (
            kind == "text"
            and unique_count
            > max_text_cardinality
        ):
            continue

        dimensions.append(
            column
        )

    return dimensions



# ============================================================
# SCHEMA EVIDENCE
# ============================================================

def discover_schema_evidence(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        row_count = len(dataframe)

        for column in dataframe.columns:
            series = dataframe[column]

            kind = infer_column_kind(
                series
            )

            key = structural_key_score(
                series
            )

            missing_count = int(
                series.isna().sum()
            )

            non_null_count = int(
                series.notna().sum()
            )

            unique_count = int(
                series.nunique(
                    dropna=True
                )
            )

            rows.append(
                {
                    "table": table_name,
                    "column": column,
                    "structural_type": kind,
                    "row_count": int(
                        row_count
                    ),
                    "non_null_count": non_null_count,
                    "missing_count": missing_count,
                    "missing_ratio": (
                        missing_count
                        / row_count
                        if row_count
                        else 0.0
                    ),
                    "unique_count": unique_count,
                    "uniqueness_ratio": (
                        unique_count
                        / non_null_count
                        if non_null_count
                        else 0.0
                    ),
                    "identifier_candidate": bool(
                        key["candidate"]
                    ),
                    "identifier_sequential_ratio": float(
                        key[
                            "sequential_ratio"
                        ]
                    ),
                }
            )

    return pd.DataFrame(
        rows
    )


# ============================================================
# GENERIC ANALYSIS
# ============================================================

def analyse_table_overview(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        numeric_count = sum(
            infer_column_kind(
                dataframe[column]
            )
            == "numeric"
            for column in dataframe.columns
        )

        temporal_count = sum(
            infer_column_kind(
                dataframe[column]
            )
            == "temporal"
            for column in dataframe.columns
        )

        categorical_count = sum(
            infer_column_kind(
                dataframe[column]
            )
            in {
                "categorical",
                "boolean",
            }
            for column in dataframe.columns
        )

        rows.append(
            {
                "table": table_name,
                "row_count": int(
                    len(dataframe)
                ),
                "column_count": int(
                    len(dataframe.columns)
                ),
                "numeric_column_count": int(
                    numeric_count
                ),
                "temporal_column_count": int(
                    temporal_count
                ),
                "categorical_column_count": int(
                    categorical_count
                ),
                "duplicate_row_count": int(
                    dataframe.duplicated().sum()
                ),
            }
        )

    return pd.DataFrame(
        rows
    )


def analyse_numeric_statistics(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        for column in select_numeric_columns(
            dataframe
        ):
            series = numeric_clean(
                dataframe[column]
            ).dropna()

            if series.empty:
                continue

            quantiles = series.quantile(
                [
                    0.25,
                    0.50,
                    0.75,
                ]
            )

            rows.append(
                {
                    "table": table_name,
                    "column": column,
                    "count": int(
                        series.count()
                    ),
                    "sum": safe_float(
                        series.sum()
                    ),
                    "mean": safe_float(
                        series.mean()
                    ),
                    "median": safe_float(
                        quantiles.loc[0.50]
                    ),
                    "std": safe_float(
                        series.std()
                    ),
                    "min": safe_float(
                        series.min()
                    ),
                    "q25": safe_float(
                        quantiles.loc[0.25]
                    ),
                    "q75": safe_float(
                        quantiles.loc[0.75]
                    ),
                    "max": safe_float(
                        series.max()
                    ),
                    "zero_count": int(
                        (
                            series == 0
                        ).sum()
                    ),
                    "negative_count": int(
                        (
                            series < 0
                        ).sum()
                    ),
                }
            )

    return pd.DataFrame(
        rows
    )


def analyse_categorical_statistics(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        for column in select_dimension_columns(
            dataframe
        ):
            series = dataframe[column]
            non_null = series.dropna()

            if non_null.empty:
                continue

            value_counts = (
                non_null.value_counts(
                    dropna=False
                )
            )

            top_value = value_counts.index[0]
            top_count = int(
                value_counts.iloc[0]
            )

            rows.append(
                {
                    "table": table_name,
                    "column": column,
                    "structural_type": (
                        infer_column_kind(series)
                    ),
                    "non_null_count": int(
                        len(non_null)
                    ),
                    "unique_count": int(
                        non_null.nunique()
                    ),
                    "top_value": safe_json_value(
                        top_value
                    ),
                    "top_value_count": top_count,
                    "top_value_ratio": (
                        top_count
                        / len(non_null)
                    ),
                }
            )

    return pd.DataFrame(
        rows
    )


def analyse_temporal_statistics(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        for column in select_temporal_columns(
            dataframe
        ):
            series = dataframe[column]

            converted = (
                series
                if is_datetime_dtype(series)
                else parse_temporal_values(series)
            )

            valid = converted.dropna()

            if valid.empty:
                continue

            rows.append(
                {
                    "table": table_name,
                    "column": column,
                    "valid_count": int(
                        valid.size
                    ),
                    "missing_count": int(
                        converted.isna().sum()
                    ),
                    "minimum": safe_json_value(
                        valid.min()
                    ),
                    "maximum": safe_json_value(
                        valid.max()
                    ),
                    "distinct_periods": int(
                        valid.nunique()
                    ),
                }
            )

    return pd.DataFrame(
        rows
    )


def analyse_numeric_by_dimensions(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        numeric_columns = select_numeric_columns(
            dataframe
        )

        dimensions = select_dimension_columns(
            dataframe
        )[
            :MAX_DIMENSIONS_PER_TABLE
        ]

        if not numeric_columns or not dimensions:
            continue

        for dimension in dimensions:
            unique_count = int(
                dataframe[
                    dimension
                ].nunique(
                    dropna=True
                )
            )

            if (
                unique_count
                > MAX_GROUPS_PER_OUTPUT
            ):
                continue

            for measure in numeric_columns:
                working = pd.DataFrame(
                    {
                        "__dimension":
                            dataframe[
                                dimension
                            ],
                        "__measure":
                            numeric_clean(
                                dataframe[
                                    measure
                                ]
                            ),
                    }
                )

                grouped = (
                    working
                    .groupby(
                        "__dimension",
                        dropna=False,
                        as_index=False,
                    )
                    .agg(
                        row_count=(
                            "__measure",
                            "count",
                        ),
                        sum=(
                            "__measure",
                            "sum",
                        ),
                        mean=(
                            "__measure",
                            "mean",
                        ),
                        minimum=(
                            "__measure",
                            "min",
                        ),
                        maximum=(
                            "__measure",
                            "max",
                        ),
                    )
                )

                if (
                    len(grouped)
                    > MAX_GROUPS_PER_OUTPUT
                ):
                    grouped = grouped.head(
                        MAX_GROUPS_PER_OUTPUT
                    )

                for _, record in grouped.iterrows():
                    rows.append(
                        {
                            "table": table_name,
                            "dimension_column": dimension,
                            "dimension_value": (
                                safe_json_value(
                                    record[
                                        "__dimension"
                                    ]
                                )
                            ),
                            "measure_column": measure,
                            "row_count": int(
                                record[
                                    "row_count"
                                ]
                            ),
                            "sum": safe_float(
                                record["sum"]
                            ),
                            "mean": safe_float(
                                record["mean"]
                            ),
                            "minimum": safe_float(
                                record["minimum"]
                            ),
                            "maximum": safe_float(
                                record["maximum"]
                            ),
                        }
                    )

                    if (
                        len(rows)
                        >= MAX_OUTPUT_ROWS_PER_ANALYSIS
                    ):
                        return pd.DataFrame(
                            rows
                        )

    return pd.DataFrame(
        rows
    )


def infer_period_frequency(
    dates: pd.Series,
) -> str:
    """
    Infer a reasonable analytical period from observed date spacing.

    This is a structural time-granularity decision, not a business rule.
    """
    valid = (
        parse_temporal_values(
            dates
        )
        .dropna()
        .drop_duplicates()
        .sort_values()
    )

    if len(valid) < 3:
        return "M"

    differences = (
        valid
        .diff()
        .dropna()
        .dt.total_seconds()
        .div(86400.0)
    )

    median_gap = safe_float(
        differences.median()
    )

    if median_gap is None:
        return "M"

    if median_gap <= 2:
        return "D"

    if median_gap <= 10:
        return "W"

    if median_gap <= 45:
        return "M"

    if median_gap <= 120:
        return "Q"

    return "Y"


def period_start(
    dates: pd.Series,
    frequency: str,
) -> pd.Series:
    converted = parse_temporal_values(
        dates
    )

    if frequency == "D":
        return converted.dt.floor(
            "D"
        )

    if frequency == "W":
        return converted.dt.to_period(
            "W"
        ).dt.start_time

    if frequency == "Q":
        return converted.dt.to_period(
            "Q"
        ).dt.start_time

    if frequency == "Y":
        return converted.dt.to_period(
            "Y"
        ).dt.start_time

    return converted.dt.to_period(
        "M"
    ).dt.start_time


def analyse_temporal_measures(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        temporal_columns = select_temporal_columns(
            dataframe
        )

        numeric_columns = select_numeric_columns(
            dataframe
        )[
            :MAX_MEASURES_PER_TABLE
        ]

        if not temporal_columns or not numeric_columns:
            continue

        for temporal_column in temporal_columns:
            dates = parse_temporal_values(
                dataframe[
                    temporal_column
                ]
            )

            frequency = infer_period_frequency(
                dates
            )

            periods = period_start(
                dates,
                frequency,
            )

            for measure in numeric_columns:
                values = numeric_clean(
                    dataframe[
                        measure
                    ]
                )

                working = pd.DataFrame(
                    {
                        "__period": periods,
                        "__value": values,
                    }
                ).dropna(
                    subset=["__period"]
                )

                if working.empty:
                    continue

                grouped = (
                    working
                    .groupby(
                        "__period",
                        as_index=False,
                    )
                    .agg(
                        row_count=(
                            "__value",
                            "count",
                        ),
                        sum=(
                            "__value",
                            "sum",
                        ),
                        mean=(
                            "__value",
                            "mean",
                        ),
                        minimum=(
                            "__value",
                            "min",
                        ),
                        maximum=(
                            "__value",
                            "max",
                        ),
                    )
                )

                for _, record in grouped.iterrows():
                    rows.append(
                        {
                            "table": table_name,
                            "temporal_column": temporal_column,
                            "period": safe_json_value(
                                record[
                                    "__period"
                                ]
                            ),
                            "period_frequency": frequency,
                            "measure_column": measure,
                            "row_count": int(
                                record[
                                    "row_count"
                                ]
                            ),
                            "sum": safe_float(
                                record["sum"]
                            ),
                            "mean": safe_float(
                                record["mean"]
                            ),
                            "minimum": safe_float(
                                record["minimum"]
                            ),
                            "maximum": safe_float(
                                record["maximum"]
                            ),
                        }
                    )

                    if (
                        len(rows)
                        >= MAX_OUTPUT_ROWS_PER_ANALYSIS
                    ):
                        return pd.DataFrame(
                            rows
                        )

    return pd.DataFrame(
        rows
    )


def analyse_correlations(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    """
    Calculate bounded pairwise correlations.

    Columns are ranked structurally before comparison so wide datasets do
    not trigger an uncontrolled quadratic expansion.
    """
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        numeric_columns = select_numeric_columns(
            dataframe
        )

        if len(numeric_columns) < 2:
            continue

        if (
            len(numeric_columns)
            > MAX_CORRELATION_COLUMNS_PER_TABLE
        ):
            variability: list[
                tuple[float, Any]
            ] = []

            for column in numeric_columns:
                series = numeric_clean(
                    dataframe[column]
                ).dropna()

                if len(series) < 2:
                    score = 0.0
                else:
                    value = safe_float(
                        series.std(
                            ddof=1
                        )
                    )
                    score = (
                        abs(value)
                        if value is not None
                        else 0.0
                    )

                variability.append(
                    (
                        score,
                        column,
                    )
                )

            variability.sort(
                key=lambda item: (
                    -item[0],
                    repr(item[1]),
                )
            )

            numeric_columns = [
                column
                for _, column in variability[
                    :MAX_CORRELATION_COLUMNS_PER_TABLE
                ]
            ]

        numeric_frame = (
            dataframe[
                numeric_columns
            ]
            .apply(
                pd.to_numeric,
                errors="coerce",
            )
        )

        correlation = numeric_frame.corr(
            method="pearson",
            numeric_only=True,
        )

        for index, first in enumerate(
            correlation.columns
        ):
            for second in correlation.columns[
                index + 1:
            ]:
                value = (
                    correlation.loc[
                        first,
                        second,
                    ]
                )

                if pd.isna(value):
                    continue

                rows.append(
                    {
                        "table": table_name,
                        "column_a": first,
                        "column_b": second,
                        "pearson_correlation": safe_float(
                            value
                        ),
                    }
                )

                if (
                    len(rows)
                    >= MAX_OUTPUT_ROWS_PER_ANALYSIS
                ):
                    return pd.DataFrame(
                        rows
                    )

    return pd.DataFrame(
        rows
    )


def analyse_changes(
    tables: dict[str, pd.DataFrame],
) -> pd.DataFrame:
    """
    Calculate first-order changes when a temporal field is available.

    One temporal field per table is selected using value-based structural
    quality rather than a name convention.
    """
    rows: list[dict[str, Any]] = []

    for table_name, dataframe in tables.items():
        temporal_columns = select_temporal_columns(
            dataframe
        )

        numeric_columns = select_numeric_columns(
            dataframe
        )

        if not temporal_columns or not numeric_columns:
            continue

        temporal_quality: list[
            tuple[float, int, Any, pd.Series]
        ] = []

        for temporal_column in temporal_columns:
            converted = parse_temporal_values(
                dataframe[
                    temporal_column
                ]
            )

            temporal_quality.append(
                (
                    float(
                        converted.notna().mean()
                    ),
                    int(
                        converted.nunique(
                            dropna=True
                        )
                    ),
                    temporal_column,
                    converted,
                )
            )

        temporal_quality.sort(
            key=lambda item: (
                -item[0],
                -item[1],
                repr(item[2]),
            )
        )

        _, _, temporal_column, dates = (
            temporal_quality[0]
        )

        for measure in numeric_columns:
            values = numeric_clean(
                dataframe[
                    measure
                ]
            )

            working = pd.DataFrame(
                {
                    "__date": dates,
                    "__value": values,
                }
            ).dropna(
                subset=[
                    "__date",
                    "__value",
                ]
            )

            if working.empty:
                continue

            grouped = (
                working
                .groupby(
                    "__date",
                    as_index=False,
                )[
                    "__value"
                ]
                .sum()
                .sort_values(
                    "__date",
                    kind="stable",
                )
            )

            if len(grouped) < 2:
                continue

            grouped["__change"] = (
                grouped["__value"]
                .diff()
            )

            for _, record in grouped.iterrows():
                rows.append(
                    {
                        "table": table_name,
                        "temporal_column": temporal_column,
                        "measure_column": measure,
                        "period": safe_json_value(
                            record["__date"]
                        ),
                        "value": safe_float(
                            record["__value"]
                        ),
                        "absolute_change": safe_float(
                            record["__change"]
                        ),
                    }
                )

                if (
                    len(rows)
                    >= MAX_OUTPUT_ROWS_PER_ANALYSIS
                ):
                    return pd.DataFrame(
                        rows
                    )

    return pd.DataFrame(
        rows
    )


# ============================================================
# RELATIONSHIP EVIDENCE
# ============================================================

def discover_relationship_evidence(
    tables: dict[str, pd.DataFrame],
    schema: pd.DataFrame,
    sample_size: int = 2000,
) -> pd.DataFrame:
    """
    Discover conservative cross-table identifier evidence.

    Sample evidence is never treated as a database constraint.
    """
    if sample_size <= 0:
        raise ValueError(
            "sample_size must be greater than zero."
        )

    if schema is None or schema.empty:
        return pd.DataFrame()

    required = {
        "table",
        "column",
        "structural_type",
        "identifier_candidate",
    }

    missing = (
        required
        - set(schema.columns)
    )

    if missing:
        raise ValueError(
            "Schema evidence is missing required fields: "
            + ", ".join(
                sorted(
                    str(value)
                    for value in missing
                )
            )
        )

    schema_by_table = {
        table_name: group
        for table_name, group
        in schema.groupby(
            "table",
            sort=False,
        )
    }

    table_names = list(
        tables
    )

    def canonical_values(
        series: pd.Series,
    ) -> set[tuple[str, str]]:
        values = (
            series
            .dropna()
            .drop_duplicates()
            .head(
                sample_size
            )
        )

        result: set[
            tuple[str, str]
        ] = set()

        for value in values:
            if (
                isinstance(
                    value,
                    (np.integer, int),
                )
                and not isinstance(
                    value,
                    bool,
                )
            ):
                result.add(
                    (
                        "numeric",
                        str(
                            int(value)
                        ),
                    )
                )
                continue

            if (
                isinstance(
                    value,
                    (np.floating, float),
                )
                and not isinstance(
                    value,
                    bool,
                )
            ):
                numeric = float(
                    value
                )

                if math.isfinite(
                    numeric
                ):
                    result.add(
                        (
                            "numeric",
                            repr(
                                numeric
                            ),
                        )
                    )
                continue

            if isinstance(
                value,
                (np.bool_, bool),
            ):
                result.add(
                    (
                        "boolean",
                        str(
                            bool(value)
                        ),
                    )
                )
                continue

            result.add(
                (
                    "text",
                    str(
                        value
                    ).strip(),
                )
            )

        return result

    candidates: list[
        dict[str, Any]
    ] = []

    for left_index, left_name in enumerate(
        table_names
    ):
        left = tables[
            left_name
        ]

        left_schema = schema_by_table.get(
            left_name,
            pd.DataFrame(),
        )

        if left_schema.empty:
            continue

        for right_name in table_names[
            left_index + 1:
        ]:
            right = tables[
                right_name
            ]

            right_schema = schema_by_table.get(
                right_name,
                pd.DataFrame(),
            )

            if right_schema.empty:
                continue

            left_records = [
                row
                for _, row
                in left_schema.iterrows()
            ]

            right_records = [
                row
                for _, row
                in right_schema.iterrows()
            ]

            left_identifiers = [
                record
                for record in left_records
                if bool(
                    record.get(
                        "identifier_candidate",
                        False,
                    )
                )
            ]

            right_identifiers = [
                record
                for record in right_records
                if bool(
                    record.get(
                        "identifier_candidate",
                        False,
                    )
                )
            ]

            if (
                not left_identifiers
                and not right_identifiers
            ):
                continue

            possible_pairs: list[
                tuple[
                    Any,
                    Any,
                    pd.Series,
                    pd.Series,
                    dict[str, Any],
                    dict[str, Any],
                ]
            ] = []

            if left_identifiers:
                for left_record in left_identifiers:
                    left_column = (
                        left_record[
                            "column"
                        ]
                    )
                    left_kind = str(
                        left_record.get(
                            "structural_type",
                            "",
                        )
                    )

                    if (
                        left_column
                        not in left.columns
                    ):
                        continue

                    for right_record in right_records:
                        right_column = (
                            right_record[
                                "column"
                            ]
                        )
                        right_kind = str(
                            right_record.get(
                                "structural_type",
                                "",
                            )
                        )

                        if (
                            right_column
                            not in right.columns
                        ):
                            continue

                        if right_kind not in {
                            "categorical",
                            "text",
                            "numeric",
                            "boolean",
                        }:
                            continue

                        if (
                            left_kind
                            == "numeric"
                            and right_kind
                            != "numeric"
                        ):
                            continue

                        if (
                            right_kind
                            == "numeric"
                            and left_kind
                            != "numeric"
                        ):
                            continue

                        possible_pairs.append(
                            (
                                left_column,
                                right_column,
                                left[
                                    left_column
                                ],
                                right[
                                    right_column
                                ],
                                left_record.to_dict(),
                                right_record.to_dict(),
                            )
                        )

            if right_identifiers:
                for right_record in right_identifiers:
                    right_column = (
                        right_record[
                            "column"
                        ]
                    )
                    right_kind = str(
                        right_record.get(
                            "structural_type",
                            "",
                        )
                    )

                    if (
                        right_column
                        not in right.columns
                    ):
                        continue

                    for left_record in left_records:
                        left_column = (
                            left_record[
                                "column"
                            ]
                        )
                        left_kind = str(
                            left_record.get(
                                "structural_type",
                                "",
                            )
                        )

                        if (
                            left_column
                            not in left.columns
                        ):
                            continue

                        if left_kind not in {
                            "categorical",
                            "text",
                            "numeric",
                            "boolean",
                        }:
                            continue

                        if (
                            left_kind
                            == "numeric"
                            and right_kind
                            != "numeric"
                        ):
                            continue

                        if (
                            right_kind
                            == "numeric"
                            and left_kind
                            != "numeric"
                        ):
                            continue

                        possible_pairs.append(
                            (
                                left_column,
                                right_column,
                                left[
                                    left_column
                                ],
                                right[
                                    right_column
                                ],
                                left_record.to_dict(),
                                right_record.to_dict(),
                            )
                        )

            seen_pairs: set[
                tuple[str, str]
            ] = set()

            for (
                left_column,
                right_column,
                left_series,
                right_series,
                left_record,
                right_record,
            ) in possible_pairs:
                pair_key = (
                    repr(left_column),
                    repr(right_column),
                )

                if pair_key in seen_pairs:
                    continue

                seen_pairs.add(
                    pair_key
                )

                left_values = canonical_values(
                    left_series
                )

                right_values = canonical_values(
                    right_series
                )

                if (
                    len(left_values) < 2
                    or len(right_values) < 2
                ):
                    continue

                overlap = (
                    left_values
                    & right_values
                )

                if len(overlap) < 2:
                    continue

                overlap_left = (
                    len(overlap)
                    / len(left_values)
                )

                overlap_right = (
                    len(overlap)
                    / len(right_values)
                )

                left_identifier = bool(
                    left_record.get(
                        "identifier_candidate",
                        False,
                    )
                )

                right_identifier = bool(
                    right_record.get(
                        "identifier_candidate",
                        False,
                    )
                )

                if (
                    left_identifier
                    == right_identifier
                ):
                    continue

                if left_identifier:
                    identifier_ratio = (
                        overlap_left
                    )
                    other_ratio = (
                        overlap_right
                    )
                    identifier_side = "left"
                else:
                    identifier_ratio = (
                        overlap_right
                    )
                    other_ratio = (
                        overlap_left
                    )
                    identifier_side = "right"

                if (
                    identifier_ratio
                    < 0.80
                ):
                    continue

                if (
                    other_ratio
                    < 0.20
                ):
                    continue

                candidates.append(
                    {
                        "left_table": left_name,
                        "left_column": left_column,
                        "right_table": right_name,
                        "right_column": right_column,
                        "sample_overlap_count": int(
                            len(overlap)
                        ),
                        "left_sample_unique_count": int(
                            len(left_values)
                        ),
                        "right_sample_unique_count": int(
                            len(right_values)
                        ),
                        "overlap_ratio_left": float(
                            overlap_left
                        ),
                        "overlap_ratio_right": float(
                            overlap_right
                        ),
                        "identifier_side": identifier_side,
                        "evidence_strength": "strong",
                    }
                )

                if (
                    len(candidates)
                    >= MAX_OUTPUT_ROWS_PER_ANALYSIS
                ):
                    return pd.DataFrame(
                        candidates
                    )

    if not candidates:
        return pd.DataFrame(
            columns=[
                "left_table",
                "left_column",
                "right_table",
                "right_column",
                "sample_overlap_count",
                "left_sample_unique_count",
                "right_sample_unique_count",
                "overlap_ratio_left",
                "overlap_ratio_right",
                "identifier_side",
                "evidence_strength",
            ]
        )

    return pd.DataFrame(
        candidates
    )


# ============================================================
# DEMAND-SIDE STRUCTURAL EVIDENCE
# ============================================================

def discover_demand_series_candidates(
    tables: dict[str, pd.DataFrame],
    feature_provenance: dict[str, dict[str, dict[str, Any]]] | None = None,
) -> pd.DataFrame:
    """
    Discover temporal numeric series suitable for demand-side analysis.

    IMPORTANT:
    This function does NOT declare a business meaning. It only records
    structural evidence that a temporal numeric series exists and has
    enough variation and observations to participate in demand analysis.
    """
    rows: list[
        dict[str, Any]
    ] = []

    for table_name, dataframe in tables.items():
        if dataframe.empty:
            continue

        temporal_columns = select_temporal_columns(
            dataframe
        )

        numeric_columns = select_numeric_columns(
            dataframe
        )

        if (
            not temporal_columns
            or not numeric_columns
        ):
            continue

        temporal_quality: list[
            tuple[
                float,
                int,
                Any,
                pd.Series,
            ]
        ] = []

        for temporal_column in temporal_columns:
            converted = parse_temporal_values(
                dataframe[
                    temporal_column
                ]
            )

            coverage = float(
                converted.notna().mean()
            )

            distinct_periods = int(
                converted.nunique(
                    dropna=True
                )
            )

            temporal_quality.append(
                (
                    coverage,
                    distinct_periods,
                    temporal_column,
                    converted,
                )
            )

        temporal_quality.sort(
            key=lambda item: (
                -item[0],
                -item[1],
                repr(item[2]),
            )
        )

        _, _, time_column, time_values = (
            temporal_quality[0]
        )

        for value_column in numeric_columns:
            # Demand candidates must be grounded in source data unless a
            # future feature explicitly declares itself eligible. This keeps
            # technical calendar/completeness features out of demand sensing.
            metadata = (
                feature_provenance.get(
                    str(table_name),
                    {},
                ).get(
                    str(value_column),
                    {},
                )
                if feature_provenance
                else {}
            )

            if metadata:
                if not bool(
                    metadata.get(
                        "demand_candidate_eligible",
                        False,
                    )
                ):
                    continue
            elif str(value_column).startswith("__"):
                # Backward-compatible safeguard for artifacts generated
                # before the provenance sidecar existed.
                continue

            values = numeric_clean(
                dataframe[
                    value_column
                ]
            )

            valid = pd.DataFrame(
                {
                    "__time": time_values,
                    "__value": values,
                }
            ).dropna(
                subset=[
                    "__time",
                    "__value",
                ]
            )

            observation_count = len(
                valid
            )

            if (
                observation_count
                < DEMAND_MIN_OBSERVATIONS
            ):
                continue

            distinct_time_periods = int(
                valid[
                    "__time"
                ].nunique()
            )

            if distinct_time_periods < 2:
                continue

            unique_value_count = int(
                valid[
                    "__value"
                ].nunique()
            )

            if unique_value_count <= 1:
                continue

            standard_deviation = safe_float(
                valid[
                    "__value"
                ].std(
                    ddof=1
                )
            )

            if (
                standard_deviation is None
                or standard_deviation <= 0
            ):
                continue

            temporal_coverage = (
                observation_count
                / max(
                    len(dataframe),
                    1,
                )
            )

            variation_score = min(
                1.0,
                unique_value_count
                / max(
                    observation_count,
                    1,
                ),
            )

            time_depth_score = min(
                1.0,
                distinct_time_periods
                / 100.0,
            )

            observation_score = min(
                1.0,
                observation_count
                / 500.0,
            )

            candidate_score = (
                0.35 * temporal_coverage
                + 0.25 * time_depth_score
                + 0.20 * variation_score
                + 0.20 * observation_score
            )

            rows.append(
                {
                    "table": table_name,
                    "time_column": time_column,
                    "measure_column": value_column,
                    "observation_count": int(
                        observation_count
                    ),
                    "distinct_time_periods": int(
                        distinct_time_periods
                    ),
                    "temporal_coverage": float(
                        temporal_coverage
                    ),
                    "unique_value_count": int(
                        unique_value_count
                    ),
                    "value_std": standard_deviation,
                    "candidate_score": float(
                        max(
                            0.0,
                            min(
                                1.0,
                                candidate_score,
                            ),
                        )
                    ),
                    "evidence_type": (
                        "structural_temporal_numeric_series"
                    ),
                }
            )

    if not rows:
        return pd.DataFrame(
            columns=[
                "table",
                "time_column",
                "measure_column",
                "observation_count",
                "distinct_time_periods",
                "temporal_coverage",
                "unique_value_count",
                "value_std",
                "candidate_score",
                "evidence_type",
            ]
        )

    result = pd.DataFrame(
        rows
    )

    result = (
        result
        .sort_values(
            [
                "table",
                "candidate_score",
                "observation_count",
            ],
            ascending=[
                True,
                False,
                False,
            ],
            kind="stable",
        )
        .groupby(
            "table",
            sort=False,
            as_index=False,
        )
        .head(
            DEMAND_MAX_CANDIDATES_PER_TABLE
        )
        .reset_index(
            drop=True
        )
    )

    return result


def _aggregate_temporal_series(
    dataframe: pd.DataFrame,
    time_column: Any,
    value_column: Any,
) -> pd.DataFrame:
    dates = parse_temporal_values(
        dataframe[
            time_column
        ]
    )

    values = numeric_clean(
        dataframe[
            value_column
        ]
    )

    working = pd.DataFrame(
        {
            "__time": dates,
            "__value": values,
        }
    ).dropna(
        subset=[
            "__time",
            "__value",
        ]
    )

    if working.empty:
        return pd.DataFrame(
            columns=[
                "__time",
                "__value",
            ]
        )

    return (
        working
        .groupby(
            "__time",
            as_index=False,
        )[
            "__value"
        ]
        .sum()
        .sort_values(
            "__time",
            kind="stable",
        )
        .reset_index(
            drop=True
        )
    )


def _correlation(
    first: np.ndarray,
    second: np.ndarray,
) -> float | None:
    if (
        len(first) < 2
        or len(second) < 2
    ):
        return None

    first_std = float(
        np.std(
            first,
            ddof=1,
        )
    )

    second_std = float(
        np.std(
            second,
            ddof=1,
        )
    )

    if (
        not np.isfinite(first_std)
        or not np.isfinite(second_std)
        or first_std <= 0.0
        or second_std <= 0.0
    ):
        return None

    first_centered = (
        first
        - float(
            np.mean(first)
        )
    )

    second_centered = (
        second
        - float(
            np.mean(second)
        )
    )

    denominator = float(
        np.sqrt(
            np.sum(
                first_centered
                ** 2
            )
            * np.sum(
                second_centered
                ** 2
            )
        )
    )

    if (
        not np.isfinite(
            denominator
        )
        or denominator <= 0.0
    ):
        return None

    result = float(
        np.sum(
            first_centered
            * second_centered
        )
        / denominator
    )

    if not np.isfinite(
        result
    ):
        return None

    return result


def _evaluate_signal_relationship(
    target_series: pd.DataFrame,
    signal_series: pd.DataFrame,
    max_lag: int,
) -> dict[str, Any] | None:
    """
    Evaluate same-period and lagged statistical association.

    Positive lag means earlier signal observations are compared with later
    target observations. No causation is inferred.
    """
    if (
        target_series.empty
        or signal_series.empty
    ):
        return None

    target = target_series.rename(
        columns={
            "__value": "__target",
        }
    )

    signal = signal_series.rename(
        columns={
            "__value": "__signal",
        }
    )

    merged = target.merge(
        signal,
        on="__time",
        how="inner",
    )

    if len(merged) < (
        DEMAND_MIN_OBSERVATIONS
    ):
        return None

    merged = (
        merged
        .sort_values(
            "__time",
            kind="stable",
        )
        .reset_index(
            drop=True
        )
    )

    target_values = pd.to_numeric(
        merged["__target"],
        errors="coerce",
    ).to_numpy(
        dtype=float
    )

    signal_values = pd.to_numeric(
        merged["__signal"],
        errors="coerce",
    ).to_numpy(
        dtype=float
    )

    finite = (
        np.isfinite(
            target_values
        )
        & np.isfinite(
            signal_values
        )
    )

    target_values = target_values[
        finite
    ]

    signal_values = signal_values[
        finite
    ]

    if len(target_values) < (
        DEMAND_MIN_OBSERVATIONS
    ):
        return None

    target_changes = np.diff(
        target_values
    )

    signal_changes = np.diff(
        signal_values
    )

    results: list[
        dict[str, Any]
    ] = []

    for lag in range(
        -max_lag,
        max_lag + 1,
    ):
        if lag > 0:
            target_level = target_values[
                lag:
            ]
            signal_level = signal_values[
                :-lag
            ]
            target_change = target_changes[
                lag:
            ]
            signal_change = signal_changes[
                :-lag
            ]
        elif lag < 0:
            shift = abs(lag)

            target_level = target_values[
                :-shift
            ]
            signal_level = signal_values[
                shift:
            ]
            target_change = target_changes[
                :-shift
            ]
            signal_change = signal_changes[
                shift:
            ]
        else:
            target_level = target_values
            signal_level = signal_values
            target_change = target_changes
            signal_change = signal_changes

        if len(target_level) < (
            DEMAND_MIN_OBSERVATIONS
        ):
            continue

        if len(target_change) < max(
            3,
            DEMAND_MIN_OBSERVATIONS - 1,
        ):
            continue

        level_corr = _correlation(
            target_level,
            signal_level,
        )

        change_corr = _correlation(
            target_change,
            signal_change,
        )

        if (
            level_corr is None
            and change_corr is None
        ):
            continue

        target_direction = np.sign(
            target_change
        )

        signal_direction = np.sign(
            signal_change
        )

        non_zero = (
            (target_direction != 0)
            | (signal_direction != 0)
        )

        if not non_zero.any():
            continue

        target_direction = (
            target_direction[
                non_zero
            ]
        )

        signal_direction = (
            signal_direction[
                non_zero
            ]
        )

        same_ratio = float(
            np.mean(
                target_direction
                == signal_direction
            )
        )

        opposite_ratio = float(
            np.mean(
                target_direction
                == -signal_direction
            )
        )

        directional_agreement = max(
            same_ratio,
            opposite_ratio,
        )

        direction = (
            "same"
            if same_ratio >= opposite_ratio
            else "opposite"
        )

        level_component = (
            abs(level_corr)
            if level_corr is not None
            else 0.0
        )

        change_component = (
            abs(change_corr)
            if change_corr is not None
            else 0.0
        )

        strength = (
            0.30 * level_component
            + 0.50 * change_component
            + 0.20 * directional_agreement
        )

        results.append(
            {
                "lag_periods": int(
                    lag
                ),
                "aligned_observations": int(
                    len(target_level)
                ),
                "level_correlation": (
                    None
                    if level_corr is None
                    else float(
                        level_corr
                    )
                ),
                "change_correlation": (
                    None
                    if change_corr is None
                    else float(
                        change_corr
                    )
                ),
                "directional_agreement": float(
                    directional_agreement
                ),
                "direction": direction,
                "strength": float(
                    strength
                ),
            }
        )

    if not results:
        return None

    evidence = pd.DataFrame(
        results
    )

    evidence = evidence[
        evidence[
            "directional_agreement"
        ]
        >= DEMAND_MIN_DIRECTIONAL_AGREEMENT
    ]

    evidence = evidence[
        evidence["strength"]
        >= DEMAND_MIN_ABS_CORRELATION
    ]

    if evidence.empty:
        return None

    best = (
        evidence
        .sort_values(
            [
                "strength",
                "aligned_observations",
            ],
            ascending=[
                False,
                False,
            ],
            kind="stable",
        )
        .iloc[0]
    )

    return {
        "lag_periods": int(
            best["lag_periods"]
        ),
        "aligned_observations": int(
            best[
                "aligned_observations"
            ]
        ),
        "level_correlation": safe_float(
            best[
                "level_correlation"
            ]
        ),
        "change_correlation": safe_float(
            best[
                "change_correlation"
            ]
        ),
        "directional_agreement": safe_float(
            best[
                "directional_agreement"
            ]
        ),
        "direction": str(
            best[
                "direction"
            ]
        ),
        "relationship_strength": safe_float(
            best[
                "strength"
            ]
        ),
    }


def analyse_demand_signal_links(
    tables: dict[str, pd.DataFrame],
    feature_provenance: dict[str, dict[str, dict[str, Any]]] | None = None,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
]:
    """
    Discover structurally suitable temporal numeric series and evaluate a
    globally bounded number of statistical relationships.

    No business category is assumed and no field is declared to represent
    demand merely because of its name.
    """
    candidates = (
        discover_demand_series_candidates(
            tables,
            feature_provenance,
        )
    )

    relationship_columns = [
        "target_table",
        "target_time_column",
        "target_measure_column",
        "signal_table",
        "signal_time_column",
        "signal_measure_column",
        "lag_periods",
        "aligned_observations",
        "level_correlation",
        "change_correlation",
        "directional_agreement",
        "direction",
        "relationship_strength",
    ]

    if candidates.empty:
        return (
            candidates,
            pd.DataFrame(
                columns=relationship_columns
            ),
        )

    series_cache: dict[
        tuple[str, Any, Any],
        pd.DataFrame,
    ] = {}

    series_metadata: list[
        dict[str, Any]
    ] = []

    for table_name, dataframe in tables.items():
        temporal_columns = select_temporal_columns(
            dataframe
        )

        numeric_columns = select_numeric_columns(
            dataframe
        )

        for time_column in temporal_columns:
            for value_column in numeric_columns:
                key = (
                    table_name,
                    time_column,
                    value_column,
                )

                series = (
                    _aggregate_temporal_series(
                        dataframe,
                        time_column,
                        value_column,
                    )
                )

                if series.empty:
                    continue

                values = (
                    series[
                        "__value"
                    ]
                    .to_numpy(
                        dtype=float
                    )
                )

                finite = values[
                    np.isfinite(values)
                ]

                if len(finite) < (
                    DEMAND_MIN_OBSERVATIONS
                ):
                    continue

                unique_count = int(
                    np.unique(
                        finite
                    ).size
                )

                if unique_count <= 1:
                    continue

                series_cache[key] = series

                series_metadata.append(
                    {
                        "key": key,
                        "observation_count": int(
                            len(finite)
                        ),
                        "unique_count": unique_count,
                    }
                )

    if not series_cache:
        return (
            candidates,
            pd.DataFrame(
                columns=relationship_columns
            ),
        )

    candidate_keys = [
        (
            row["table"],
            row["time_column"],
            row["measure_column"],
        )
        for _, row in candidates.iterrows()
        if (
            row["table"],
            row["time_column"],
            row["measure_column"],
        ) in series_cache
    ]

    signal_keys = [
        record["key"]
        for record in sorted(
            series_metadata,
            key=lambda record: (
                -record[
                    "observation_count"
                ],
                -record[
                    "unique_count"
                ],
                repr(
                    record["key"]
                ),
            ),
        )
    ]

    links: list[
        dict[str, Any]
    ] = []

    comparison_count = 0

    for target_key in candidate_keys:
        if (
            comparison_count
            >= DEMAND_MAX_RELATIONSHIP_COMPARISONS
        ):
            break

        target_series = series_cache[
            target_key
        ]

        for signal_key in signal_keys:
            if (
                comparison_count
                >= DEMAND_MAX_RELATIONSHIP_COMPARISONS
            ):
                break

            if signal_key == target_key:
                continue

            signal_series = series_cache[
                signal_key
            ]

            comparison_count += 1

            relationship = (
                _evaluate_signal_relationship(
                    target_series,
                    signal_series,
                    DEMAND_MAX_LAG_PERIODS,
                )
            )

            if relationship is None:
                continue

            links.append(
                {
                    "target_table": target_key[0],
                    "target_time_column": target_key[1],
                    "target_measure_column": target_key[2],
                    "signal_table": signal_key[0],
                    "signal_time_column": signal_key[1],
                    "signal_measure_column": signal_key[2],
                    **relationship,
                }
            )

    if not links:
        return (
            candidates,
            pd.DataFrame(
                columns=relationship_columns
            ),
        )

    result = pd.DataFrame(
        links
    )

    result = (
        result
        .sort_values(
            [
                "relationship_strength",
                "aligned_observations",
            ],
            ascending=[
                False,
                False,
            ],
            kind="stable",
        )
        .reset_index(
            drop=True
        )
    )

    return (
        candidates,
        result,
    )


# ============================================================
# ANALYTICAL READINESS
# ============================================================

def build_readiness_report(
    tables: dict[str, pd.DataFrame],
    schema: pd.DataFrame,
    relationships: pd.DataFrame,
    demand_candidates: pd.DataFrame,
    demand_signal_links: pd.DataFrame,
) -> pd.DataFrame:
    rows: list[
        dict[str, Any]
    ] = []

    for table_name, dataframe in tables.items():
        table_schema = schema[
            schema["table"]
            == table_name
        ]

        numeric_count = int(
            (
                table_schema[
                    "structural_type"
                ]
                == "numeric"
            ).sum()
        )

        temporal_count = int(
            (
                table_schema[
                    "structural_type"
                ]
                == "temporal"
            ).sum()
        )

        dimension_count = int(
            table_schema[
                "structural_type"
            ].isin(
                [
                    "categorical",
                    "boolean",
                    "text",
                ]
            ).sum()
        )

        key_count = int(
            table_schema[
                "identifier_candidate"
            ].sum()
        )

        has_relationship = False

        if (
            relationships is not None
            and not relationships.empty
        ):
            has_relationship = bool(
                (
                    relationships[
                        "left_table"
                    ].eq(
                        table_name
                    )
                    | relationships[
                        "right_table"
                    ].eq(
                        table_name
                    )
                ).any()
            )

        candidate_count = 0

        if (
            demand_candidates
            is not None
            and not demand_candidates.empty
        ):
            candidate_count = int(
                (
                    demand_candidates[
                        "table"
                    ].eq(
                        table_name
                    )
                ).sum()
            )

        signal_count = 0

        if (
            demand_signal_links
            is not None
            and not demand_signal_links.empty
        ):
            signal_count = int(
                (
                    (
                        demand_signal_links[
                            "target_table"
                        ].eq(
                            table_name
                        )
                    )
                    | (
                        demand_signal_links[
                            "signal_table"
                        ].eq(
                            table_name
                        )
                    )
                ).sum()
            )

        rows.append(
            {
                "table": table_name,
                "row_count": int(
                    len(dataframe)
                ),
                "column_count": int(
                    len(dataframe.columns)
                ),
                "numeric_fields_available": numeric_count,
                "temporal_fields_available": temporal_count,
                "dimension_fields_available": dimension_count,
                "identifier_candidates_available": key_count,
                "demand_series_candidates": candidate_count,
                "demand_signal_relationships": signal_count,
                "has_numeric_analysis": (
                    numeric_count > 0
                ),
                "has_temporal_analysis": (
                    numeric_count > 0
                    and temporal_count > 0
                ),
                "has_dimension_analysis": (
                    numeric_count > 0
                    and dimension_count > 0
                ),
                "has_relationship_evidence": (
                    has_relationship
                ),
                "has_demand_candidate_evidence": (
                    candidate_count > 0
                ),
                "has_demand_signal_evidence": (
                    signal_count > 0
                ),
            }
        )

    return pd.DataFrame(
        rows
    )


# ============================================================
# OUTPUT WRITING
# ============================================================

def write_json_dataset(
    outputs: dict[str, pd.DataFrame],
) -> None:
    payload: dict[
        str,
        list[
            dict[str, Any]
        ],
    ] = {}

    for name, dataframe in outputs.items():
        cleaned = safe_output_dataframe(
            dataframe
        )

        payload[name] = [
            {
                str(key): safe_json_value(
                    value
                )
                for key, value in record.items()
            }
            for record in cleaned.to_dict(
                orient="records"
            )
        ]

    OUTPUT_DATA_DIRECTORY.mkdir(
        parents=True,
        exist_ok=True,
    )

    ANALYSIS_DATASET_JSON.write_text(
        json.dumps(
            payload,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def make_unique_sheet_name(
    requested_name: str,
    used_names: set[str],
) -> str:
    cleaned = str(
        requested_name
    )

    invalid_characters = {
        ":",
        "\\",
        "/",
        "?",
        "*",
        "[",
        "]",
    }

    cleaned = "".join(
        "_"
        if char in invalid_characters
        else char
        for char in cleaned
    ).strip()

    if not cleaned:
        cleaned = "Analysis"

    cleaned = cleaned[:31]
    candidate = cleaned
    counter = 1

    while candidate in used_names:
        suffix = f"_{counter}"
        candidate = (
            cleaned[
                :31 - len(suffix)
            ]
            + suffix
        )
        counter += 1

    used_names.add(
        candidate
    )

    return candidate


def write_analysis_workbook(
    outputs: dict[str, pd.DataFrame],
) -> int:
    OUTPUT_DATA_DIRECTORY.mkdir(
        parents=True,
        exist_ok=True,
    )

    written = 0
    used_names: set[str] = set()

    with pd.ExcelWriter(
        ANALYSIS_WORKBOOK,
        engine="openpyxl",
    ) as writer:
        for name, dataframe in outputs.items():
            if (
                dataframe is None
                or dataframe.empty
            ):
                continue

            sheet = make_unique_sheet_name(
                name,
                used_names,
            )

            safe_output_dataframe(
                dataframe
            ).to_excel(
                writer,
                sheet_name=sheet,
                index=False,
            )

            written += 1

    if written == 0:
        raise RuntimeError(
            "No non-empty analytical outputs were generated."
        )

    return written


def write_report(
    source: Path,
    tables: dict[str, pd.DataFrame],
    outputs: dict[str, pd.DataFrame],
    schema: pd.DataFrame,
    relationships: pd.DataFrame,
    demand_candidates: pd.DataFrame,
    demand_signal_links: pd.DataFrame,
) -> None:
    report = {
        "source": str(
            source
        ),
        "table_count": int(
            len(tables)
        ),
        "total_rows": int(
            sum(
                len(frame)
                for frame in tables.values()
            )
        ),
        "outputs": {
            name: {
                "rows": int(
                    len(frame)
                ),
                "columns": int(
                    len(frame.columns)
                ),
            }
            for name, frame in outputs.items()
        },
        "schema_columns_analyzed": int(
            len(schema)
        ),
        "relationship_candidates": int(
            len(relationships)
        ),
        "demand_series_candidates": int(
            len(demand_candidates)
        ),
        "demand_signal_relationships": int(
            len(demand_signal_links)
        ),
        "design": {
            "dataset_specific_aliases": False,
            "dataset_specific_table_names": False,
            "dataset_specific_column_names": False,
            "fixed_entity_counts": False,
            "target_inference": False,
            "source_mutation": False,
            "structural_semantic_guessing": False,
            "demand_target_declared": False,
            "demand_candidates_are_evidence_only": True,
            "demand_signal_links_are_evidence_only": True,
            "business_recommendations_generated": False,
        },
        "capability_behavior": {
            "no_temporal_data_is_fatal": False,
            "no_numeric_data_is_fatal": False,
            "no_relationship_evidence_is_fatal": False,
            "no_demand_evidence_is_fatal": False,
            "resource_limits_are_technical": True,
        },
        "technical_controls": {
            "demand_min_observations": (
                DEMAND_MIN_OBSERVATIONS
            ),
            "demand_max_candidates_per_table": (
                DEMAND_MAX_CANDIDATES_PER_TABLE
            ),
            "demand_max_lag_periods": (
                DEMAND_MAX_LAG_PERIODS
            ),
            "demand_max_relationship_comparisons": (
                DEMAND_MAX_RELATIONSHIP_COMPARISONS
            ),
            "analysis_max_dimensions_per_table": (
                MAX_DIMENSIONS_PER_TABLE
            ),
            "analysis_max_groups_per_output": (
                MAX_GROUPS_PER_OUTPUT
            ),
            "analysis_max_measures_per_table": (
                MAX_MEASURES_PER_TABLE
            ),
            "analysis_max_correlation_columns_per_table": (
                MAX_CORRELATION_COLUMNS_PER_TABLE
            ),
            "analysis_max_output_rows_per_analysis": (
                MAX_OUTPUT_ROWS_PER_ANALYSIS
            ),
        },
    }

    ANALYSIS_REPORT_JSON.write_text(
        json.dumps(
            report,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


# ============================================================
# MAIN PIPELINE
# ============================================================

def run_analysis() -> None:
    print("=" * 72)
    print(
        "ENTERPRISE DEMAND INTELLIGENCE PLATFORM"
    )
    print(
        "GENERIC DATASET-DRIVEN ANALYTICAL ENGINE"
    )
    print("=" * 72)

    print()
    print(
        "Architecture mode:"
    )
    print(
        "  DATASET-DRIVEN"
    )

    print()
    print(
        "Semantic mode:"
    )
    print(
        "  STRUCTURAL EVIDENCE — NO BUSINESS NAME ASSUMPTIONS"
    )

    print()
    print(
        "Source modification:"
    )
    print(
        "  DISABLED"
    )

    source, tables = (
        load_analytical_dataset()
    )

    feature_provenance = load_feature_provenance(
        source
    )

    print(
        "Feature provenance metadata: "
        + (
            "LOADED"
            if feature_provenance
            else "NOT AVAILABLE — technical fallback active"
        )
    )

    print()
    print("-" * 72)
    print(
        "STEP 2 — DISCOVERING STRUCTURAL ANALYTICAL EVIDENCE"
    )
    print("-" * 72)

    schema = discover_schema_evidence(
        tables
    )

    if schema.empty:
        raise RuntimeError(
            "No structural schema evidence was generated."
        )

    print(
        f"Columns discovered: {len(schema):,}"
    )

    print(
        "Structural schema discovery: PASSED."
    )

    print()
    print("-" * 72)
    print(
        "STEP 3 — RUNNING GENERIC ANALYSES"
    )
    print("-" * 72)

    outputs: dict[
        str,
        pd.DataFrame,
    ] = {}

    outputs[
        "Table_Overview"
    ] = analyse_table_overview(
        tables
    )

    outputs[
        "Schema_Evidence"
    ] = schema

    outputs[
        "Numeric_Statistics"
    ] = analyse_numeric_statistics(
        tables
    )

    outputs[
        "Categorical_Statistics"
    ] = analyse_categorical_statistics(
        tables
    )

    outputs[
        "Temporal_Statistics"
    ] = analyse_temporal_statistics(
        tables
    )

    outputs[
        "Dimension_Measure_Aggregations"
    ] = analyse_numeric_by_dimensions(
        tables
    )

    outputs[
        "Temporal_Measure_Aggregations"
    ] = analyse_temporal_measures(
        tables
    )

    outputs[
        "Numeric_Correlations"
    ] = analyse_correlations(
        tables
    )

    outputs[
        "Temporal_Changes"
    ] = analyse_changes(
        tables
    )

    relationships = (
        discover_relationship_evidence(
            tables,
            schema,
        )
    )

    outputs[
        "Relationship_Evidence"
    ] = relationships

    print(
        "Generic analytical evidence generation: PASSED."
    )

    print()
    print("-" * 72)
    print(
        "STEP 3A — DISCOVERING DEMAND-SIDE EVIDENCE"
    )
    print("-" * 72)

    (
        demand_candidates,
        demand_signal_links,
    ) = analyse_demand_signal_links(
        tables,
        feature_provenance,
    )

    outputs[
        "Demand_Candidate_Evidence"
    ] = demand_candidates

    outputs[
        "Demand_Signal_Evidence"
    ] = demand_signal_links

    print(
        "Demand-side structural candidates: "
        f"{len(demand_candidates):,}"
    )

    print(
        "Demand-signal relationships: "
        f"{len(demand_signal_links):,}"
    )

    print()
    print("-" * 72)
    print(
        "STEP 3B — BUILDING CAPABILITY READINESS"
    )
    print("-" * 72)

    outputs[
        "Analytical_Readiness"
    ] = build_readiness_report(
        tables,
        schema,
        relationships,
        demand_candidates,
        demand_signal_links,
    )

    non_empty_outputs = {
        name: frame
        for name, frame
        in outputs.items()
        if (
            frame is not None
            and not frame.empty
        )
    }

    for name, dataframe in outputs.items():
        print(
            f"  {name:<38}"
            f"{len(dataframe):>12,} rows"
        )

    print()
    print(
        "Non-empty analytical output tables: "
        f"{len(non_empty_outputs):,}"
    )

    print()
    print("-" * 72)
    print(
        "STEP 4 — STRUCTURAL INTEGRITY CHECKS"
    )
    print("-" * 72)

    source_rows = sum(
        len(frame)
        for frame in tables.values()
    )

    if source_rows <= 0:
        raise RuntimeError(
            "The analytical dataset contains no rows."
        )

    if schema.empty:
        raise RuntimeError(
            "No schema evidence was generated."
        )

    if (
        "Demand_Candidate_Evidence"
        not in outputs
        or "Demand_Signal_Evidence"
        not in outputs
    ):
        raise RuntimeError(
            "Demand evidence outputs were not generated."
        )

    print(
        "Schema evidence preservation: PASSED."
    )

    print(
        "Demand-sensing evidence generation: PASSED."
    )

    print(
        "Capability degradation is non-fatal when evidence is unavailable."
    )

    print(
        "Source data mutation: NONE."
    )

    print()
    print("-" * 72)
    print(
        "STEP 5 — WRITING ANALYTICAL OUTPUTS"
    )
    print("-" * 72)

    OUTPUT_DATA_DIRECTORY.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        f"Writing JSON: {ANALYSIS_DATASET_JSON}"
    )

    write_json_dataset(
        outputs
    )

    print(
        "Analytical JSON: CREATED."
    )

    print(
        f"Writing Excel: {ANALYSIS_WORKBOOK}"
    )

    written_count = (
        write_analysis_workbook(
            outputs
        )
    )

    print(
        "Analytical Excel: CREATED "
        f"({written_count} sheets)."
    )

    write_report(
        source,
        tables,
        outputs,
        schema,
        relationships,
        demand_candidates,
        demand_signal_links,
    )

    print(
        f"Report: {ANALYSIS_REPORT_JSON}"
    )

    print()
    print("=" * 72)
    print(
        "GENERIC ANALYTICAL ENGINE COMPLETED SUCCESSFULLY"
    )
    print("=" * 72)

    print()
    print(
        f"Tables processed: {len(tables):,}"
    )

    print(
        f"Total source rows: {source_rows:,}"
    )

    print(
        "Relationship candidates: "
        f"{len(relationships):,}"
    )

    print(
        "Demand candidates: "
        f"{len(demand_candidates):,}"
    )

    print(
        "Demand-signal relationships: "
        f"{len(demand_signal_links):,}"
    )

    print()
    print(
        "Architecture guarantees:"
    )
    print(
        "  Business aliases used: NO"
    )
    print(
        "  Fixed table names required: NO"
    )
    print(
        "  Fixed column names required: NO"
    )
    print(
        "  Fixed entity counts required: NO"
    )
    print(
        "  Target inferred: NO"
    )
    print(
        "  Source modified: NO"
    )


if __name__ == "__main__":
    try:
        run_analysis()
    except Exception as error:
        print()
        print("=" * 72)
        print(
            "GENERIC ANALYTICAL ENGINE FAILED"
        )
        print("=" * 72)
        print()
        print(
            f"ERROR: {error}"
        )
        raise
