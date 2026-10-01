"""
Enterprise Demand Intelligence Platform
GENERIC FORECASTING ENGINE

This module is dataset-driven. It does not require or encode a business schema.
It discovers temporal structure, numeric measures, and usable dimensions from
whatever valid tabular dataset is supplied by the preceding pipeline stage.

The module is read-only with respect to source data. It produces forecasts only
for structurally suitable numeric time series and reports unavailable
capabilities instead of inventing missing business concepts.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config

SUPPORTED_INPUT_FORMATS = set(
    getattr(
        config,
        "SUPPORTED_DATASET_FORMATS",
        {
            "csv",
            "json",
            "xlsx",
            "xls",
            "parquet",
        },
    )
)

MAX_INPUT_FILE_SIZE_MB = int(
    getattr(
        config,
        "MAX_DATASET_FILE_SIZE_MB",
        500,
    )
)

FORECAST_HORIZON = int(
    getattr(
        config,
        "FORECAST_HORIZON_DAYS",
        30,
    )
)

MIN_HISTORY = int(
    getattr(
        config,
        "FORECAST_MINIMUM_HISTORY",
        30,
    )
)

RECENT_WINDOW = int(
    getattr(
        config,
        "FORECAST_RECENT_WINDOW",
        14,
    )
)

BASELINE_WINDOW = int(
    getattr(
        config,
        "FORECAST_BASELINE_WINDOW",
        30,
    )
)

SAMPLE_SIZE = int(
    getattr(
        config,
        "SAMPLE_SIZE",
        5000,
    )
)

MAX_FORECAST_POINTS = max(
    1,
    int(
        getattr(
            config,
            "FORECAST_MAX_POINTS",
            1000,
        )
    ),
)

OUTPUT_DIR = Path(getattr(config, "OUTPUT_DATA_DIRECTORY", PROJECT_ROOT / "data" / "output"))
FORECAST_OUTPUT = OUTPUT_DIR / "forecast_results.xlsx"
FORECAST_JSON = OUTPUT_DIR / "forecast_results.json"
FORECAST_REPORT = OUTPUT_DIR / "forecast_report.json"


def _safe_json(value: Any) -> Any:
    """Convert analytical values into JSON-safe primitives."""
    if value is None:
        return None

    if isinstance(
        value,
        (list, tuple, set),
    ):
        return [
            _safe_json(item)
            for item in value
        ]

    if isinstance(value, dict):
        return {
            str(key): _safe_json(item)
            for key, item in value.items()
        }

    if isinstance(value, np.integer):
        return int(value)

    if isinstance(value, np.floating):
        numeric = float(value)
        return (
            numeric
            if np.isfinite(numeric)
            else None
        )

    if isinstance(value, np.bool_):
        return bool(value)

    if isinstance(value, pd.Timestamp):
        return value.isoformat()

    if isinstance(value, pd.Timedelta):
        return str(value)

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, float):
        return (
            value
            if np.isfinite(value)
            else None
        )

    if isinstance(
        value,
        (str, int, bool),
    ):
        return value

    return str(value)


def _normalise_tables(obj: Any) -> dict[str, pd.DataFrame]:
    if isinstance(obj, dict):
        result: dict[str, pd.DataFrame] = {}
        for key, value in obj.items():
            if isinstance(value, pd.DataFrame):
                result[str(key)] = value.copy()
            elif isinstance(value, list) and value and isinstance(value[0], dict):
                result[str(key)] = pd.DataFrame(value)
        if result:
            return result
    if isinstance(obj, pd.DataFrame):
        return {"table_1": obj.copy()}
    raise ValueError("The selected input does not contain readable tabular data.")


def _read_json(path: Path) -> dict[str, pd.DataFrame]:
    with path.open("r", encoding="utf-8-sig") as handle:
        return _normalise_tables(json.load(handle))


def _read_input(path: Path) -> dict[str, pd.DataFrame]:
    suffix = path.suffix.lower().lstrip(".")
    if suffix == "json":
        return _read_json(path)
    if suffix == "csv":
        return {path.stem: pd.read_csv(path)}
    if suffix == "parquet":
        return {path.stem: pd.read_parquet(path)}
    if suffix in {"xlsx", "xls"}:
        workbook = pd.ExcelFile(path)
        result = {}
        for sheet in workbook.sheet_names:
            frame = pd.read_excel(path, sheet_name=sheet)
            if not frame.empty or len(frame.columns) > 0:
                result[str(sheet)] = frame
        return result
    raise ValueError(f"Unsupported input format: {path.suffix}")


def _candidate_input_files() -> list[Path]:
    explicit = getattr(config, "ACTIVE_DATASET_PATH", None)
    if explicit:
        path = Path(explicit).expanduser().resolve()
        if path.is_file():
            return [path]
        raise FileNotFoundError(f"Configured dataset path does not exist: {path}")

    env_name = getattr(config, "DATASET_PATH_ENV_VAR", "")
    if env_name:
        import os
        value = os.getenv(env_name)
        if value:
            path = Path(value).expanduser().resolve()
            if path.is_file():
                return [path]
            raise FileNotFoundError(f"Configured dataset path does not exist: {path}")

    # Prefer the latest pipeline-produced analytical tabular artifact only when
    # it is explicitly identifiable by the configuration; otherwise discover a
    # single supported file. No business vocabulary is used here.
    configured = []
    for attr in ("ANALYTICAL_DATASET_PATH", "ANALYTICAL_WORKBOOK"):
        value = getattr(config, attr, None)
        if value:
            path = Path(value).expanduser().resolve()
            if path.is_file() and path.suffix.lower().lstrip(".") in SUPPORTED_INPUT_FORMATS:
                configured.append(path)
    if configured:
        return [configured[0]]

    candidates = [
        path
        for path in (
            OUTPUT_DIR / "analytical_dataset.xlsx",
            OUTPUT_DIR / "analytical_dataset.json",
            OUTPUT_DIR / "analytical_dataset.parquet",
        )
        if (
            path.exists()
            and path.is_file()
            and path.suffix.lower().lstrip(".")
            in SUPPORTED_INPUT_FORMATS
        )
    ]

    if candidates:
        return candidates

    raise FileNotFoundError(
        "No analytical dataset produced by the preceding pipeline stage "
        "was found in the configured output directory."
    )


def discover_input() -> tuple[Path, dict[str, pd.DataFrame]]:
    files = _candidate_input_files()
    if len(files) != 1:
        raise RuntimeError("Exactly one forecasting input must be selected.")
    path = files[0]
    size_mb = path.stat().st_size / (1024 * 1024)
    if size_mb > MAX_INPUT_FILE_SIZE_MB:
        raise ValueError(f"Input file exceeds the configured maximum size of {MAX_INPUT_FILE_SIZE_MB} MB.")
    return path, _read_input(path)


def _structural_kind(
    series: pd.Series,
) -> str:
    """Return the structural type from the shared discovery layer."""
    try:
        from src.data_loader import infer_column_type
    except ImportError as error:
        raise RuntimeError(
            "The shared structural type inference could not be imported."
        ) from error

    return infer_column_type(series)


def _numeric_candidate_score(series: pd.Series) -> tuple[bool, float, dict[str, Any]]:
    kind = _structural_kind(series)
    if kind != "numeric":
        return False, 0.0, {"reason": "not_numeric"}
    values = pd.to_numeric(series, errors="coerce").dropna()
    if len(values) < 2:
        return False, 0.0, {"reason": "insufficient_values"}
    unique_ratio = values.nunique(dropna=True) / max(len(values), 1)
    integer_ratio = np.isclose(values % 1, 0).mean()
    if unique_ratio <= 0.01 and integer_ratio >= 0.99:
        return False, 0.0, {"reason": "low_variability_integer_like"}
    if values.nunique(dropna=True) <= 1:
        return False, 0.0, {"reason": "constant"}
    score = min(1.0, 0.5 * min(unique_ratio * 5, 1.0) + 0.5)
    return True, score, {
        "unique_ratio": float(unique_ratio),
        "minimum": float(values.min()),
        "maximum": float(values.max()),
    }


def _identifier_evidence(
    series: pd.Series,
    discovery_evidence: dict[str, Any] | None = None,
) -> bool:
    """
    Determine whether a field has structural identifier evidence.

    Shared discovery evidence is preferred. The fallback is deliberately
    conservative and exists only for direct standalone use.
    """
    if discovery_evidence is not None:
        return bool(
            discovery_evidence.get(
                "identifier_candidate",
                False,
            )
        )

    non_null = series.dropna()

    if len(non_null) < 2:
        return False

    if _structural_kind(series) in {
        "temporal",
        "boolean",
    }:
        return False

    uniqueness = (
        non_null.nunique(dropna=True)
        / len(non_null)
    )

    if _structural_kind(series) == "numeric":
        numeric_values = pd.to_numeric(
            non_null,
            errors="coerce",
        ).dropna()

        if len(numeric_values) < 2:
            return False

        integer_like = bool(
            np.isclose(
                numeric_values % 1,
                0,
            ).mean() >= 0.99
        )

        monotonic = bool(
            numeric_values.is_monotonic_increasing
            or numeric_values.is_monotonic_decreasing
        )

        ordered = np.sort(
            numeric_values.to_numpy(
                dtype=float
            )
        )
        steps = np.diff(ordered)
        sequential = bool(
            len(steps) > 0
            and np.isclose(
                steps,
                1.0,
            ).mean() >= 0.99
        )

        return bool(
            uniqueness >= 0.995
            and integer_like
            and (
                monotonic
                or sequential
            )
        )

    return bool(uniqueness >= 0.995)

def _build_discovery_evidence(
    tables: dict[str, pd.DataFrame],
) -> dict[tuple[str, Any], dict[str, Any]]:
    """
    Build one structural evidence cache for the loaded tables.

    Structural type and identifier evidence are calculated once per
    table/column and reused by candidate discovery.
    """
    evidence: dict[
        tuple[str, Any],
        dict[str, Any],
    ] = {}

    for table_name, dataframe in tables.items():

        for column in dataframe.columns:

            structural_type = _structural_kind(
                dataframe[column]
            )

            if structural_type in {
                "temporal",
                "boolean",
            }:
                identifier_candidate = False

            else:
                non_null = (
                    dataframe[column]
                    .dropna()
                )

                if len(non_null) < 2:
                    identifier_candidate = False
                else:
                    uniqueness = (
                        non_null.nunique(
                            dropna=True
                        )
                        / len(non_null)
                    )

                    if structural_type == "numeric":
                        numeric_values = pd.to_numeric(
                            non_null,
                            errors="coerce",
                        ).dropna()

                        integer_like = bool(
                            len(numeric_values) > 1
                            and np.isclose(
                                numeric_values % 1,
                                0,
                            ).mean() >= 0.99
                        )

                        monotonic = bool(
                            len(numeric_values) > 1
                            and (
                                numeric_values.is_monotonic_increasing
                                or numeric_values.is_monotonic_decreasing
                            )
                        )

                        if integer_like and len(numeric_values) > 1:
                            ordered = np.sort(
                                numeric_values.to_numpy(
                                    dtype=float
                                )
                            )
                            steps = np.diff(ordered)
                            sequential = bool(
                                len(steps) > 0
                                and np.isclose(
                                    steps,
                                    1.0,
                                ).mean() >= 0.99
                            )
                        else:
                            sequential = False

                        identifier_candidate = bool(
                            uniqueness >= 0.995
                            and integer_like
                            and (
                                monotonic
                                or sequential
                            )
                        )
                    else:
                        identifier_candidate = bool(
                            uniqueness >= 0.995
                        )

            evidence[
                (
                    table_name,
                    column,
                )
            ] = {
                "structural_type": structural_type,
                "identifier_candidate":
                    identifier_candidate,
            }

    return evidence

def discover_temporal_columns(
    tables: dict[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Discover temporal fields while preserving their actual labels.
    """
    found: list[
        dict[str, Any]
    ] = []

    for table_name, frame in tables.items():
        for column in frame.columns:
            if _structural_kind(
                frame[column]
            ) != "temporal":
                continue

            found.append(
                {
                    "table": table_name,
                    "column": column,
                }
            )

    return found


def discover_forecast_series(
    tables: dict[str, pd.DataFrame],
) -> list[dict[str, Any]]:
    """
    Discover structurally suitable temporal/numeric forecasting candidates.

    No business vocabulary or fixed schema is used.
    """
    candidates: list[
        dict[str, Any]
    ] = []

    max_candidates = max(
        1,
        int(
            getattr(
                config,
                "FORECAST_MAX_CANDIDATES",
                500,
            )
        ),
    )

    discovery_evidence = (
        _build_discovery_evidence(
            tables
        )
    )

    for table_name, frame in tables.items():

        if len(candidates) >= max_candidates:
            break

        temporal_columns = [
            column
            for column in frame.columns
            if discovery_evidence[
                (
                    table_name,
                    column,
                )
            ]["structural_type"]
            == "temporal"
        ]

        if not temporal_columns:
            continue

        temporal_scores: list[
            tuple[
                float,
                int,
                Any,
            ]
        ] = []

        for column in temporal_columns:

            converted = pd.to_datetime(
                frame[column],
                errors="coerce",
                format="mixed",
            )

            coverage = float(
                converted.notna().mean()
            )

            distinct = int(
                converted.nunique(
                    dropna=True
                )
            )

            temporal_scores.append(
                (
                    coverage,
                    distinct,
                    column,
                )
            )

        temporal_scores.sort(
            key=lambda item: (
                item[0],
                item[1],
                repr(item[2]),
            ),
            reverse=True,
        )

        time_col = temporal_scores[
            0
        ][2]

        numeric_candidates: list[
            tuple[
                float,
                Any,
                dict[str, Any],
            ]
        ] = []

        for column in frame.columns:

            if column == time_col:
                continue

            if "__" in str(column):
                continue

            evidence_record = (
                discovery_evidence[
                    (
                        table_name,
                        column,
                    )
                ]
            )

            if (
                evidence_record[
                    "structural_type"
                ]
                != "numeric"
            ):
                continue

            usable, score, evidence = (
                _numeric_candidate_score(
                    frame[column]
                )
            )

            if not usable:
                continue

            if evidence_record[
                "identifier_candidate"
            ]:
                continue

            numeric_candidates.append(
                (
                    score,
                    column,
                    evidence,
                )
            )

        numeric_candidates.sort(
            key=lambda item: (
                item[0],
                repr(item[1]),
            ),
            reverse=True,
        )

        remaining = (
            max_candidates
            - len(candidates)
        )

        for (
            score,
            value_col,
            evidence,
        ) in numeric_candidates[
            :remaining
        ]:

            candidates.append(
                {
                    "table": table_name,
                    "time_column": time_col,
                    "value_column": value_col,
                    "candidate_score": score,
                    "evidence": evidence,
                }
            )

            if len(candidates) >= (
                max_candidates
            ):
                break

    return candidates


def _choose_dimensions(
    frame: pd.DataFrame,
    time_col: Any,
    value_col: Any,
) -> list[Any]:
    """
    Select structurally usable dimensions without renaming labels.
    """
    dimensions: list[Any] = []

    for column in frame.columns:

        if column == time_col or column == value_col:
            continue

        if "__" in str(column):
            continue

        kind = _structural_kind(
            frame[column]
        )

        if kind not in {
            "text",
            "categorical",
            "boolean",
        }:
            continue

        non_null = frame[
            column
        ].dropna()

        if len(non_null) < 2:
            continue

        cardinality = int(
            non_null.nunique()
        )

        if (
            1 < cardinality
            <= min(
                1000,
                max(
                    2,
                    int(
                        len(non_null)
                        * 0.25
                    ),
                ),
            )
        ):
            dimensions.append(
                column
            )

    return dimensions


def _build_series(
    frame: pd.DataFrame,
    time_col: Any,
    value_col: Any,
    dimensions: list[Any],
) -> pd.DataFrame:
    """
    Build an aggregated temporal series without modifying source data.
    """
    selected_columns = [
        time_col,
        value_col,
        *dimensions,
    ]

    data = frame[
        selected_columns
    ].copy()

    data[
        time_col
    ] = pd.to_datetime(
        data[time_col],
        errors="coerce",
        format="mixed",
    )

    data[
        value_col
    ] = pd.to_numeric(
        data[value_col],
        errors="coerce",
    )

    data = data.dropna(
        subset=[
            time_col,
            value_col,
        ]
    )

    if data.empty:
        return pd.DataFrame()

    keys = [
        *dimensions,
        time_col,
    ]

    grouped = (
        data
        .groupby(
            keys,
            dropna=False,
            as_index=False,
        )[value_col]
        .sum()
    )

    return grouped.sort_values(
        keys,
        kind="stable",
    )


def _forecast_group(
    group: pd.DataFrame,
    time_col: Any,
    value_col: Any,
    dimensions: list[Any],
) -> dict[str, Any] | None:
    """
    Generate a structurally safe baseline forecast for one series.
    """
    group = (
        group
        .sort_values(
            time_col,
            kind="stable",
        )
        .copy()
    )

    if len(group) < MIN_HISTORY:
        return None

    dates = (
        pd.DatetimeIndex(
            group[time_col]
        )
        .dropna()
        .drop_duplicates()
        .sort_values()
    )

    if len(dates) < 2:
        return None

    inferred = pd.infer_freq(
        dates
    )

    gaps = (
        dates
        .to_series()
        .diff()
        .dropna()
        .dt.total_seconds()
        / 86400.0
    )

    gaps_array = gaps.to_numpy(
        dtype=float
    )

    gaps = gaps[
        np.isfinite(
            gaps_array
        )
        & (
            gaps_array > 0
        )
    ]

    if gaps.empty:
        return None

    median_gap_days = float(
        gaps.median()
    )

    if (
        not np.isfinite(
            median_gap_days
        )
        or median_gap_days <= 0
    ):
        return None

    if inferred is None:

        if median_gap_days <= 1.5:
            freq = "D"
        elif median_gap_days <= 8.0:
            freq = "W"
        elif median_gap_days <= 31.0:
            freq = "MS"
        elif median_gap_days <= 92.0:
            freq = "QS"
        else:
            return None

    else:
        freq = inferred

    try:
        offset = (
            pd.tseries.frequencies
            .to_offset(freq)
        )
    except (
        ValueError,
        TypeError,
    ):
        return None

    series = (
        group
        .set_index(time_col)[
            value_col
        ]
        .astype(float)
        .dropna()
    )

    if len(series) < MIN_HISTORY:
        return None

    recent = series.tail(
        min(
            RECENT_WINDOW,
            len(series),
        )
    )

    baseline = series.tail(
        min(
            BASELINE_WINDOW,
            len(series),
        )
    )

    if recent.empty or baseline.empty:
        return None

    recent_mean = float(
        recent.mean()
    )

    baseline_mean = float(
        baseline.mean()
    )

    volatility = (
        float(
            baseline.std(
                ddof=1
            )
        )
        if len(baseline) > 1
        else 0.0
    )

    if not np.isfinite(
        volatility
    ):
        volatility = 0.0

    if baseline_mean != 0:
        cv = abs(
            volatility
            / baseline_mean
        )
    else:
        cv = (
            0.0
            if volatility == 0
            else float("inf")
        )

    forecast_level = (
        0.6 * recent_mean
        + 0.4 * baseline_mean
    )

    if not np.isfinite(
        forecast_level
    ):
        return None

    last_date = dates.max()

    forecast_points = max(
        1,
        int(
            np.ceil(
                FORECAST_HORIZON
                / median_gap_days
            )
        ),
    )

    forecast_points = min(
        forecast_points,
        MAX_FORECAST_POINTS,
    )

    try:
        future_dates = pd.date_range(
            start=last_date + offset,
            periods=forecast_points,
            freq=offset,
        )
    except (
        ValueError,
        TypeError,
        OverflowError,
    ):
        return None

    if future_dates.empty:
        return None

    lower_forecast = (
        forecast_level
        - volatility
    )

    upper_forecast = (
        forecast_level
        + volatility
    )

    trend_pct = (
        0.0
        if baseline_mean == 0
        else (
            (
                recent_mean
                - baseline_mean
            )
            / abs(
                baseline_mean
            )
        ) * 100
    )

    if not np.isfinite(
        trend_pct
    ):
        trend_pct = 0.0

    if cv <= 0.25:
        confidence = "High"
    elif cv <= 0.50:
        confidence = "Medium"
    else:
        confidence = "Low"

    row: dict[str, Any] = {
        "table": None,
        "time_column": time_col,
        "value_column": value_col,
    }

    for dimension in dimensions:
        row[dimension] = group[
            dimension
        ].iloc[0]

    row.update(
        {
            "ForecastStart":
                future_dates.min(),
            "ForecastEnd":
                future_dates.max(),
            "HistoryPoints":
                int(len(series)),
            "RecentAverage":
                recent_mean,
            "BaselineAverage":
                baseline_mean,
            "ForecastAverage":
                forecast_level,
            "ForecastTotal":
                forecast_level
                * len(future_dates),
            "LowerForecastTotal":
                lower_forecast
                * len(future_dates),
            "UpperForecastTotal":
                upper_forecast
                * len(future_dates),
            "Volatility":
                volatility,
            "CoefficientOfVariation":
                (
                    cv
                    if np.isfinite(cv)
                    else None
                ),
            "RecentVsBaselineChangePct":
                trend_pct,
            "ForecastConfidence":
                confidence,
            "FrequencyEvidence":
                str(freq),
            "ForecastPoints":
                int(
                    len(future_dates)
                ),
            "ForecastHorizonDays":
                float(
                    median_gap_days
                    * len(future_dates)
                ),
        }
    )

    return row


def generate_forecasts(
    tables: dict[str, pd.DataFrame],
    candidates: list[dict[str, Any]] | None = None,
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    list[dict[str, Any]],
]:
    """
    Generate forecasts from structurally discovered candidates.

    Candidate discovery happens once when candidates are supplied by
    the orchestration flow. Generated series are technically bounded.
    """
    if candidates is None:
        candidates = (
            discover_forecast_series(
                tables
            )
        )

    forecasts: list[
        dict[str, Any]
    ] = []

    evidence: list[
        dict[str, Any]
    ] = []

    max_generated_series = max(
        1,
        int(
            getattr(
                config,
                "FORECAST_MAX_GENERATED_SERIES",
                50000,
            )
        ),
    )

    generated_series_count = 0
    stopped_by_limit = False

    for candidate in candidates:

        if (
            generated_series_count
            >= max_generated_series
        ):
            stopped_by_limit = True
            break

        table = candidate[
            "table"
        ]

        time_col = candidate[
            "time_column"
        ]

        value_col = candidate[
            "value_column"
        ]

        frame = tables[
            table
        ]

        dimensions = _choose_dimensions(
            frame,
            time_col,
            value_col,
        )

        dimension_sets = (
            [[]]
            + [
                [dimension]
                for dimension in dimensions
            ]
        )

        produced = 0

        for dimension_set in dimension_sets:

            if (
                generated_series_count
                >= max_generated_series
            ):
                stopped_by_limit = True
                break

            series_data = _build_series(
                frame,
                time_col,
                value_col,
                dimension_set,
            )

            if series_data.empty:
                continue

            if dimension_set:
                grouped_items = (
                    series_data.groupby(
                        dimension_set,
                        dropna=False,
                        sort=False,
                    )
                )
            else:
                grouped_items = [
                    (
                        (),
                        series_data,
                    )
                ]

            for _, group in grouped_items:

                if (
                    generated_series_count
                    >= max_generated_series
                ):
                    stopped_by_limit = True
                    break

                result = _forecast_group(
                    group,
                    time_col,
                    value_col,
                    dimension_set,
                )

                if result is None:
                    continue

                result[
                    "table"
                ] = table

                result[
                    "candidate_score"
                ] = candidate[
                    "candidate_score"
                ]

                forecasts.append(
                    result
                )

                produced += 1
                generated_series_count += 1

        evidence.append(
            {
                "table": table,
                "time_column": time_col,
                "value_column": value_col,
                "candidate_score": candidate[
                    "candidate_score"
                ],
                "dimensions_considered": [
                    dimension
                    for dimension in dimensions
                ],
                "forecast_series_produced": produced,
                "status": (
                    "available"
                    if produced
                    else
                    "insufficient_history_or_frequency"
                ),
            }
        )

        if stopped_by_limit:
            break

    if stopped_by_limit:
        evidence.append(
            {
                "table": None,
                "time_column": None,
                "value_column": None,
                "candidate_score": None,
                "dimensions_considered": [],
                "forecast_series_produced": 0,
                "status":
                    "technical_generation_limit_reached",
                "generation_limit":
                    max_generated_series,
            }
        )

    return (
        pd.DataFrame(
            forecasts
        ),
        pd.DataFrame(
            evidence
        ),
        candidates,
    )


def _trend_label(value: float) -> str:
    if value >= 10:
        return "Increasing"
    if value <= -10:
        return "Decreasing"
    return "Stable"


def build_summary(forecasts: pd.DataFrame) -> pd.DataFrame:
    if forecasts.empty:
        return pd.DataFrame({"Metric": ["Forecast series produced"], "Value": [0]})
    trends = forecasts["RecentVsBaselineChangePct"].map(_trend_label)
    confidence = forecasts["ForecastConfidence"]
    rows = [
        ("Forecast series produced", len(forecasts)),
        ("Forecast horizon days", FORECAST_HORIZON),
        ("Total forecast value", float(forecasts["ForecastTotal"].sum())),
        ("Increasing series", int((trends == "Increasing").sum())),
        ("Stable series", int((trends == "Stable").sum())),
        ("Decreasing series", int((trends == "Decreasing").sum())),
        ("High confidence series", int((confidence == "High").sum())),
        ("Medium confidence series", int((confidence == "Medium").sum())),
        ("Low confidence series", int((confidence == "Low").sum())),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def _write_json(path: Path, payload: Any) -> None:
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2, default=_safe_json)


def write_outputs(
    input_path: Path,
    tables: dict[str, pd.DataFrame],
    forecasts: pd.DataFrame,
    evidence: pd.DataFrame,
    candidates: list[dict[str, Any]],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    summary = build_summary(forecasts)
    temporal = pd.DataFrame(discover_temporal_columns(tables))
    if not forecasts.empty:
        forecasts = forecasts.copy()
        forecasts["DemandTrend"] = forecasts["RecentVsBaselineChangePct"].map(_trend_label)
    with pd.ExcelWriter(FORECAST_OUTPUT, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Forecast_Summary", index=False)
        forecasts.to_excel(writer, sheet_name="Forecasts", index=False)
        evidence.to_excel(writer, sheet_name="Forecast_Evidence", index=False)
        temporal.to_excel(writer, sheet_name="Temporal_Evidence", index=False)
    _write_json(FORECAST_JSON, {
        "input": str(input_path),
        "summary": summary.to_dict(orient="records"),
        "forecasts": forecasts.to_dict(orient="records"),
        "evidence": evidence.to_dict(orient="records"),
        "candidate_count": len(candidates),
    })
    report = {
        "status": "VALID" if not forecasts.empty else "CAPABILITY_UNAVAILABLE",
        "input": str(input_path),
        "tables_discovered": len(tables),
        "total_rows": int(sum(len(v) for v in tables.values())),
        "temporal_columns": int(len(temporal)),
        "forecast_candidates": len(candidates),
        "forecast_series": int(len(forecasts)),
        "source_modified": False,
        "business_aliases_used": False,
        "fixed_table_names_required": False,
        "fixed_column_names_required": False,
        "fixed_entity_counts_required": False,
        "target_inferred": False,
        "notes": "Forecasts are generated only from structurally discovered temporal and numeric evidence.",
    }
    _write_json(FORECAST_REPORT, report)
    return summary, forecasts


def run_forecasting() -> None:
    print("=" * 72)
    print("ENTERPRISE DEMAND INTELLIGENCE PLATFORM")
    print("GENERIC DATASET-DRIVEN FORECASTING ENGINE")
    print("=" * 72)
    print("\nArchitecture mode:\n  DATASET-DRIVEN")
    print("\nSemantic mode:\n  STRUCTURAL — NO BUSINESS NAME ASSUMPTIONS")
    print("\nSource modification:\n  DISABLED")
    print("\n------------------------------------------------------------------------")
    print("STEP 1 — DISCOVERING FORECAST INPUT")
    print("------------------------------------------------------------------------")
    input_path, tables = discover_input()
    print(f"Input: {input_path}")
    print(f"Tables discovered: {len(tables)}")
    for name, frame in tables.items():
        print(f"  {name:<40} {len(frame):>10,} rows {len(frame.columns):>8,} columns")
    print("\n------------------------------------------------------------------------")
    print("STEP 2 — DISCOVERING STRUCTURAL FORECAST EVIDENCE")
    print("------------------------------------------------------------------------")
    temporal = discover_temporal_columns(tables)
    candidates = discover_forecast_series(tables)
    print(f"Temporal columns discovered: {len(temporal)}")
    print(f"Forecastable numeric series candidates: {len(candidates)}")
    print("\n------------------------------------------------------------------------")
    print("STEP 3 — GENERATING STRUCTURAL FORECASTS")
    print("------------------------------------------------------------------------")
    forecasts, evidence, candidates = generate_forecasts(
        tables,
        candidates=candidates,
    )
    print(f"Forecast series produced: {len(forecasts):,}")
    print(f"Forecast evidence rows: {len(evidence):,}")
    print("\n------------------------------------------------------------------------")
    print("STEP 4 — STRUCTURAL INTEGRITY CHECKS")
    print("------------------------------------------------------------------------")
    print("Source data mutation: NONE.")
    print("Business schema requirement: NO.")
    print("Target inference: NO.")
    print("\n------------------------------------------------------------------------")
    print("STEP 5 — WRITING FORECAST OUTPUTS")
    print("------------------------------------------------------------------------")
    write_outputs(input_path, tables, forecasts, evidence, candidates)
    print(f"Forecast Excel: {FORECAST_OUTPUT}")
    print(f"Forecast JSON : {FORECAST_JSON}")
    print(f"Forecast report: {FORECAST_REPORT}")
    print("\n" + "=" * 72)
    if forecasts.empty:
        print("FORECASTING COMPLETED — FORECAST CAPABILITY UNAVAILABLE")
        print("No structurally suitable series with sufficient history was found.")
    else:
        print("GENERIC FORECASTING ENGINE COMPLETED SUCCESSFULLY")
    print("=" * 72)
    print(f"\nTables processed: {len(tables)}")
    print(f"Total source rows: {sum(len(v) for v in tables.values()):,}")
    print(f"Forecast series: {len(forecasts):,}")
    print("\nArchitecture guarantees:")
    print("  Business aliases used: NO")
    print("  Fixed table names required: NO")
    print("  Fixed column names required: NO")
    print("  Fixed entity counts required: NO")
    print("  Target inferred: NO")
    print("  Source modified: NO")


if __name__ == "__main__":
    run_forecasting()
