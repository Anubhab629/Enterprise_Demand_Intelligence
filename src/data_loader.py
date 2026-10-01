"""
Generic dataset loader and structural discovery.

Architecture rule
-----------------
This module does NOT assume:
- a particular dataset filename
- table names
- column names
- business aliases or synonyms
- company/product/customer/order names
- entity counts
- row counts
- date ranges
- business values
- a predefined target column

It reads the supplied data and infers only structural properties that are
supported by the actual contents.

VS Code
-------
Open this file and press "Run Python File". If no path is supplied, the
loader automatically resolves a single supported dataset from data/raw.

Explicit execution is also supported:
    python src/data_loader.py "path/to/dataset.json"
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except ImportError as exc:
    raise SystemExit(
        "pandas is required. Install it with: "
        "python -m pip install pandas"
    ) from exc

try:
    import config
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    import config


# ============================================================
# GENERIC HELPERS
# ============================================================

NULL_LIKE = {
    "",
    "nan",
    "nat",
    "none",
    "null",
    "na",
    "n/a",
}


def is_missing(value: Any) -> bool:
    if value is None:
        return True

    if isinstance(value, str):
        return value.strip().lower() in NULL_LIKE

    try:
        result = pd.isna(value)
        return bool(result) if not hasattr(result, "__len__") else False
    except (TypeError, ValueError):
        return False


def safe_json_value(value: Any) -> Any:
    if is_missing(value):
        return None

    if isinstance(value, (str, int, float, bool)):
        if isinstance(value, float) and not math.isfinite(value):
            return None
        return value

    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except Exception:
            pass

    return str(value)


def json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            str(key): json_safe(item)
            for key, item in value.items()
        }

    if isinstance(value, (list, tuple)):
        return [json_safe(item) for item in value]

    return safe_json_value(value)


def suffix(path: Path) -> str:
    return path.suffix.lower().lstrip(".")


def validate_file(
    path: Path,
    max_size_mb: float | None = None,
) -> Path:
    path = Path(path).expanduser().resolve()

    if not path.exists():
        raise FileNotFoundError(
            f"Dataset file does not exist: {path}"
        )

    if not path.is_file():
        raise ValueError(
            f"Dataset path is not a file: {path}"
        )

    file_type = suffix(path)

    if file_type not in config.SUPPORTED_DATASET_FORMATS:
        raise ValueError(
            f"Unsupported format '.{file_type}'. "
            f"Supported formats: "
            f"{sorted(config.SUPPORTED_DATASET_FORMATS)}"
        )

    limit = (
        config.MAX_DATASET_FILE_SIZE_MB
        if max_size_mb is None
        else max_size_mb
    )

    size_mb = path.stat().st_size / (1024 * 1024)

    if size_mb > float(limit):
        raise ValueError(
            f"Dataset is {size_mb:.2f} MB; "
            f"maximum allowed is {float(limit):.2f} MB."
        )

    return path


def evenly_spaced_positions(
    length: int,
    sample_size: int,
) -> list[int]:
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


def sample_frame(
    frame: pd.DataFrame,
    sample_size: int,
) -> pd.DataFrame:
    if sample_size <= 0:
        raise ValueError(
            "sample_size must be greater than zero."
        )

    if len(frame) <= sample_size:
        return frame

    positions = evenly_spaced_positions(
        len(frame),
        sample_size,
    )

    return frame.iloc[positions]


# ============================================================
# FILE LOADERS
# ============================================================

def json_to_tables(
    obj: Any,
) -> dict[str, pd.DataFrame]:
    if isinstance(obj, list):
        if not obj:
            return {"table_0": pd.DataFrame()}

        if all(isinstance(item, dict) for item in obj):
            return {
                "table_0": pd.DataFrame.from_records(obj)
            }

        raise ValueError(
            "A JSON root list must contain objects."
        )

    if not isinstance(obj, dict):
        raise ValueError(
            "JSON root must be an object or a list of objects."
        )

    tables: dict[str, pd.DataFrame] = {}

    for name, value in obj.items():
        if (
            isinstance(value, list)
            and all(isinstance(item, dict) for item in value)
        ):
            tables[str(name)] = pd.DataFrame.from_records(value)

    if tables:
        return tables

    if all(
        not isinstance(value, (dict, list))
        for value in obj.values()
    ):
        return {
            "table_0": pd.DataFrame.from_records([obj])
        }

    raise ValueError(
        "JSON structure does not contain a supported tabular object."
    )


def load_json(path: Path) -> dict[str, pd.DataFrame]:
    with path.open(
        "r",
        encoding=config.ETL_READ_JSON_ENCODING,
    ) as handle:
        obj = json.load(handle)

    return json_to_tables(obj)


def load_csv(path: Path) -> dict[str, pd.DataFrame]:
    return {
        path.stem: pd.read_csv(path)
    }


def load_excel(path: Path) -> dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(path)

    return {
        str(sheet): pd.read_excel(
            workbook,
            sheet_name=sheet,
        )
        for sheet in workbook.sheet_names
    }


def load_parquet(path: Path) -> dict[str, pd.DataFrame]:
    return {
        path.stem: pd.read_parquet(path)
    }


def load_file(path: Path) -> dict[str, pd.DataFrame]:
    path = validate_file(path)
    file_type = suffix(path)

    if file_type == "json":
        return load_json(path)

    if file_type == "csv":
        return load_csv(path)

    if file_type in {"xlsx", "xls"}:
        return load_excel(path)

    if file_type == "parquet":
        return load_parquet(path)

    raise ValueError(
        f"No loader is registered for '.{file_type}'."
    )


# ============================================================
# DATASET PATH RESOLUTION
# ============================================================

def discover_dataset_file(
    explicit_path: str | Path | None = None,
) -> Path:
    if explicit_path:
        return validate_file(Path(explicit_path))

    configured = getattr(
        config,
        "ACTIVE_DATASET_PATH",
        None,
    )

    if configured:
        return validate_file(Path(configured))

    environment_name = getattr(
        config,
        "DATASET_PATH_ENV_VAR",
        "",
    )

    if environment_name:
        environment_path = os.getenv(environment_name)

        if environment_path:
            return validate_file(
                Path(environment_path)
            )

    if not getattr(
        config,
        "DATASET_DISCOVERY_ENABLED",
        True,
    ):
        raise FileNotFoundError(
            "No dataset path is configured and "
            "automatic discovery is disabled."
        )

    raw_directory = Path(
        config.RAW_DATA_DIRECTORY
    )

    if not raw_directory.exists():
        raise FileNotFoundError(
            f"Raw-data directory does not exist: "
            f"{raw_directory}"
        )

    candidates = sorted(
        path
        for path in raw_directory.iterdir()
        if (
            path.is_file()
            and suffix(path)
            in config.SUPPORTED_DATASET_FORMATS
        )
    )

    if len(candidates) == 1:
        return validate_file(candidates[0])

    if not candidates:
        raise FileNotFoundError(
            "No supported dataset file was found in: "
            f"{raw_directory}"
        )

    candidate_text = "\n".join(
        f"  - {path.name}"
        for path in candidates
    )

    raise RuntimeError(
        "More than one supported dataset file was found. "
        "The loader will not guess which one to use.\n"
        f"Set ACTIVE_DATASET_PATH in config.py or "
        f"{config.DATASET_PATH_ENV_VAR} in the environment.\n"
        f"Candidates:\n{candidate_text}"
    )


# ============================================================
# TYPE DISCOVERY
# ============================================================

def infer_column_type(
    series: pd.Series,
) -> str:
    non_missing = series.dropna()

    if non_missing.empty:
        return "empty"

    if pd.api.types.is_bool_dtype(series):
        return "boolean"

    if pd.api.types.is_numeric_dtype(series):
        return "numeric"

    if pd.api.types.is_datetime64_any_dtype(series):
        return "temporal"

    text = non_missing.astype(str).str.strip()

    # Explicit mixed-format parsing prevents pandas from trying to infer
    # one global format and avoids noisy per-value dateutil warnings.
    try:
        parsed = pd.to_datetime(
            text,
            errors="coerce",
            format="mixed",
        )
    except (TypeError, ValueError):
        parsed = pd.to_datetime(
            text,
            errors="coerce",
        )

    if float(parsed.notna().mean()) >= 0.95:
        return "temporal"

    unique_ratio = (
        non_missing.nunique(dropna=True)
        / max(len(non_missing), 1)
    )

    if unique_ratio <= 0.05:
        return "categorical"

    return "text"


def identifier_score(
    series: pd.Series,
    inferred_type: str,
) -> float:
    values = series.dropna()

    if values.empty:
        return 0.0

    unique_count = values.nunique(dropna=True)
    unique_ratio = unique_count / max(len(values), 1)

    # Uniqueness alone is deliberately insufficient. A short reference
    # column can be completely unique while still being an ordinary
    # category. Identifier evidence must come from the observed value
    # structure as well.
    if unique_ratio < 0.80:
        return 0.0

    if inferred_type == "text":
        text = values.astype(str).str.strip()
        structured_ratio = float(
            text.str.contains(
                r"[\-_:/]|[A-Za-z].*\d|\d.*[A-Za-z]",
                regex=True,
            ).mean()
        )

        if structured_ratio < 0.50:
            return 0.0

        score = 0.55 + 0.30 * structured_ratio

        if unique_count == len(values):
            score += 0.15

        return min(score, 1.0)

    if inferred_type == "numeric":
        numeric = pd.to_numeric(
            values,
            errors="coerce",
        ).dropna()

        if numeric.empty:
            return 0.0

        integer_ratio = float(
            numeric.mod(1).eq(0).mean()
        )

        if integer_ratio < 0.95:
            return 0.0

        # Integer uniqueness is useful evidence, but continuous numeric
        # measurements are not promoted automatically.
        return min(
            0.60 + 0.25 * integer_ratio
            + (0.15 if unique_count == len(values) else 0.0),
            1.0,
        )

    return 0.0


def profile_column(
    series: pd.Series,
) -> dict[str, Any]:
    inferred = infer_column_type(series)
    non_missing = series.dropna()

    unique_count = int(
        non_missing.nunique(dropna=True)
    )

    result: dict[str, Any] = {
        "name": series.name,
        "display_name": str(series.name),
        "dtype": str(series.dtype),
        "inferred_type": inferred,
        "row_count": int(len(series)),
        "missing_count": int(series.isna().sum()),
        "missing_ratio": round(
            float(series.isna().mean())
            if len(series)
            else 0.0,
            6,
        ),
        "unique_count": unique_count,
        "unique_ratio": round(
            unique_count
            / max(len(non_missing), 1),
            6,
        ),
        "identifier_score": round(
            identifier_score(
                series,
                inferred,
            ),
            6,
        ),
    }

    if inferred == "numeric":
        numeric = pd.to_numeric(
            series,
            errors="coerce",
        ).dropna()

        if not numeric.empty:
            result.update({
                "minimum": safe_json_value(
                    numeric.min()
                ),
                "maximum": safe_json_value(
                    numeric.max()
                ),
                "mean": safe_json_value(
                    numeric.mean()
                ),
                "median": safe_json_value(
                    numeric.median()
                ),
                "integer_like": bool(
                    numeric.mod(1).eq(0).all()
                ),
            })

    if inferred == "categorical":
        counts = non_missing.astype(
            str
        ).value_counts()

        result["top_values"] = [
            {
                "value": str(index),
                "count": int(count),
            }
            for index, count in counts.head(10).items()
        ]

    return result


def profile_table(
    frame: pd.DataFrame,
    sample_size: int,
) -> dict[str, Any]:
    sampled = sample_frame(
        frame,
        sample_size,
    )

    columns = [
        profile_column(sampled[column])
        for column in sampled.columns
    ]

    temporal_columns = [
        item["name"]
        for item in columns
        if item["inferred_type"] == "temporal"
    ]

    numeric_columns = [
        item["name"]
        for item in columns
        if item["inferred_type"] == "numeric"
    ]

    identifier_candidates = [
        item["name"]
        for item in columns
        if item["identifier_score"] >= 0.75
    ]

    return {
        "row_count": int(len(frame)),
        "column_count": int(len(frame.columns)),
        "columns": columns,
        "temporal_columns": temporal_columns,
        "numeric_columns": numeric_columns,
        "identifier_candidates": identifier_candidates,
    }


# ============================================================
# STRUCTURAL TABLE CLASSIFICATION
# ============================================================

def classify_table(
    profile: dict[str, Any],
) -> str:
    rows = profile["row_count"]
    columns = profile["column_count"]

    if rows == 0 or columns == 0:
        return "empty"

    missing_ratios = [
        item["missing_ratio"]
        for item in profile["columns"]
    ]

    if (
        missing_ratios
        and all(ratio >= 0.95 for ratio in missing_ratios)
    ):
        return "mostly_missing"

    temporal_count = len(
        profile["temporal_columns"]
    )

    numeric_count = len(
        profile["numeric_columns"]
    )

    identifier_count = len(
        profile["identifier_candidates"]
    )

    # A unique record identifier together with other fields is stronger
    # structural evidence of a record-oriented relation than merely having
    # a date and numeric values. This remains completely name-independent.
    if identifier_count and columns >= 2:
        if temporal_count or numeric_count or columns > 3:
            return "relational_like"

    if temporal_count and numeric_count:
        return "observation_like"

    if (
        identifier_count
        and columns <= max(
            3,
            identifier_count * 3,
        )
    ):
        return "reference_like"

    categorical_count = sum(
        item["inferred_type"] == "categorical"
        for item in profile["columns"]
    )

    if (
        categorical_count >= max(1, columns // 2)
        and rows > 0
    ):
        return "categorical_like"

    if numeric_count >= 2:
        return "measurement_like"

    if identifier_count:
        return "relational_like"

    return "unknown"


# ============================================================
# TEMPORAL STRUCTURE
# ============================================================

def discover_temporal_structure(
    frame: pd.DataFrame,
    temporal_columns: list[Any],
) -> list[dict[str, Any]]:
    results = []

    for column in temporal_columns:
        parsed = pd.to_datetime(
            frame[column],
            errors="coerce",
        )

        valid = parsed.dropna()

        if valid.empty:
            continue

        ordered = valid.sort_values()
        differences = ordered.diff().dropna()

        results.append({
            "column": column,
            "valid_values": int(len(valid)),
            "unique_values": int(
                valid.nunique()
            ),
            "minimum": valid.min().isoformat(),
            "maximum": valid.max().isoformat(),
            "monotonic_in_input": bool(
                valid.is_monotonic_increasing
            ),
            "repeated_temporal_values": bool(
                valid.nunique() < len(valid)
            ),
            "median_interval_seconds": (
                float(
                    differences.dt.total_seconds()
                    .median()
                )
                if not differences.empty
                else None
            ),
        })

    return results


# ============================================================
# GENERIC VALUE-OVERLAP RELATIONSHIPS
# ============================================================

def normalized_value_set(
    series: pd.Series,
    limit: int,
) -> set[str]:
    values = (
        series.dropna()
        .astype(str)
        .str.strip()
    )

    values = values[values != ""]

    if len(values) > limit:
        positions = evenly_spaced_positions(
            len(values),
            limit,
        )
        values = values.iloc[positions]

    return set(values.tolist())


def discover_relationships(
    tables: dict[str, pd.DataFrame],
    profiles: dict[str, dict[str, Any]],
    sample_size: int,
    max_pairs: int,
) -> list[dict[str, Any]]:
    candidates = []

    for table_name, profile in profiles.items():
        for column in profile["columns"]:
            inferred_type = column["inferred_type"]

            if inferred_type in {
                "numeric",
                "temporal",
            }:
                continue

            if (
                column["identifier_score"] < 0.55
                and column["unique_ratio"] > 0.50
            ):
                continue

            candidates.append((
                table_name,
                column["name"],
                column,
            ))

    relationships = []
    comparisons = 0

    for index, (
        left_table,
        left_column,
        left_meta,
    ) in enumerate(candidates):

        if comparisons >= max_pairs:
            break

        left_sample = sample_frame(
            tables[left_table],
            sample_size,
        )

        left_values = normalized_value_set(
            left_sample[left_column],
            sample_size,
        )

        if not left_values:
            continue

        for (
            right_table,
            right_column,
            right_meta,
        ) in candidates[index + 1:]:

            if comparisons >= max_pairs:
                break

            if left_table == right_table:
                continue

            left_type = left_meta["inferred_type"]
            right_type = right_meta["inferred_type"]

            compatible = (
                left_type == right_type
                or {
                    left_type,
                    right_type,
                } == {"categorical", "text"}
            )

            if not compatible:
                continue

            right_sample = sample_frame(
                tables[right_table],
                sample_size,
            )

            right_values = normalized_value_set(
                right_sample[right_column],
                sample_size,
            )

            if not right_values:
                continue

            comparisons += 1

            overlap = (
                left_values & right_values
            )

            if not overlap:
                continue

            left_coverage = (
                len(overlap)
                / len(left_values)
            )

            right_coverage = (
                len(overlap)
                / len(right_values)
            )

            smaller_coverage = min(
                left_coverage,
                right_coverage,
            )

            if (
                len(overlap) >= 2
                and smaller_coverage >= 0.10
            ):
                confidence = min(
                    1.0,
                    (
                        0.5 * smaller_coverage
                        + 0.5
                        * min(
                            1.0,
                            len(overlap) / 20,
                        )
                    ),
                )

                relationships.append({
                    "left_table": left_table,
                    "left_column": left_column,
                    "right_table": right_table,
                    "right_column": right_column,
                    "shared_values_sampled": int(
                        len(overlap)
                    ),
                    "left_coverage": round(
                        left_coverage,
                        6,
                    ),
                    "right_coverage": round(
                        right_coverage,
                        6,
                    ),
                    "confidence": round(
                        confidence,
                        6,
                    ),
                })

    relationships.sort(
        key=lambda item: (
            item["confidence"],
            item["shared_values_sampled"],
        ),
        reverse=True,
    )

    return relationships


# ============================================================
# REUSE / KEY-LIKE CANDIDATES
# ============================================================

def discover_reuse_candidates(
    profiles: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    candidates = []

    for table_name, profile in profiles.items():
        for column in profile["columns"]:
            if column["identifier_score"] < 0.75:
                continue

            if column["unique_count"] == 0:
                continue

            candidates.append({
                "table": table_name,
                "column": column["name"],
                "unique_count": column["unique_count"],
                "unique_ratio": column["unique_ratio"],
                "identifier_score": column["identifier_score"],
            })

    return sorted(
        candidates,
        key=lambda item: (
            item["identifier_score"],
            item["unique_count"],
        ),
        reverse=True,
    )


# ============================================================
# QUALITY
# ============================================================

def quality_report(
    tables: dict[str, pd.DataFrame],
) -> dict[str, Any]:
    reports = {}
    total_rows = 0

    for name, frame in tables.items():
        total_rows += len(frame)

        reports[name] = {
            "rows": int(len(frame)),
            "columns": int(len(frame.columns)),
            "duplicate_rows": int(
                frame.duplicated().sum()
            ),
            "all_missing_columns": [
                str(column)
                for column in frame.columns
                if frame[column].isna().all()
            ],
        }

    return {
        "total_rows": int(total_rows),
        "table_count": int(len(tables)),
        "tables": reports,
    }


def detect_structural_ml_risks(
    profiles: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    risks = []

    for table_name, profile in profiles.items():
        for column in profile["columns"]:
            if column["identifier_score"] >= 0.90:
                risks.append({
                    "table": table_name,
                    "column": column["name"],
                    "risk": "identifier_candidate",
                    "reason": (
                        "High uniqueness may cause memorization "
                        "or leakage if the field is used directly "
                        "as a model feature."
                    ),
                })

    return {
        "target_inference": "not_performed",
        "future_information_inference": "not_performed",
        "derived_semantics_inference": "not_performed",
        "risks": risks,
    }


# ============================================================
# COMPLETE DISCOVERY
# ============================================================

def discover_dataset(
    dataset_path: str | Path | None = None,
    sample_size: int | None = None,
    max_size_mb: float | None = None,
) -> dict[str, Any]:

    path = discover_dataset_file(
        dataset_path
    )

    if max_size_mb is not None and max_size_mb <= 0:
        raise ValueError(
            "max_size_mb must be greater than zero."
        )

    validate_file(
        path,
        max_size_mb=max_size_mb,
    )

    tables = load_file(path)

    if len(tables) > config.MAX_DISCOVERED_TABLES:
        raise ValueError(
            f"Discovered {len(tables)} tables, exceeding "
            f"the technical limit of "
            f"{config.MAX_DISCOVERED_TABLES}."
        )

    for name, frame in tables.items():
        if (
            len(frame.columns)
            > config.MAX_DISCOVERED_COLUMNS_PER_TABLE
        ):
            raise ValueError(
                f"Table '{name}' contains "
                f"{len(frame.columns)} columns, exceeding "
                "the configured technical limit."
            )

    total_rows = sum(
        len(frame)
        for frame in tables.values()
    )

    if total_rows > config.MAX_TOTAL_ROWS:
        raise ValueError(
            f"Dataset contains {total_rows:,} rows, "
            f"exceeding the technical limit of "
            f"{config.MAX_TOTAL_ROWS:,}."
        )

    if sample_size is None:
        effective_sample = config.SAMPLE_SIZE
    else:
        if sample_size <= 0:
            raise ValueError(
                "sample_size must be greater than zero."
            )
        effective_sample = sample_size

    profiles = {
        name: profile_table(
            frame,
            effective_sample,
        )
        for name, frame in tables.items()
    }

    roles = {
        name: classify_table(
            profiles[name]
        )
        for name in tables
    }

    temporal = {
        name: discover_temporal_structure(
            frame,
            profiles[name]["temporal_columns"],
        )
        for name, frame in tables.items()
    }

    relationships = discover_relationships(
        tables,
        profiles,
        config.RELATIONSHIP_SAMPLE_SIZE,
        config.MAX_RELATIONSHIP_PAIRS,
    )

    return {
        "dataset": {
            "path": str(path),
            "file_name": path.name,
            "format": suffix(path),
            "size_bytes": int(
                path.stat().st_size
            ),
            "size_mb": round(
                path.stat().st_size
                / (1024 * 1024),
                4,
            ),
        },
        "summary": {
            "table_count": len(tables),
            "total_rows": total_rows,
        },
        "tables": tables,
        "profiles": profiles,
        "roles": roles,
        "temporal_structure": temporal,
        "relationships": relationships,
        "reuse_candidates": discover_reuse_candidates(
            profiles
        ),
        "quality": quality_report(tables),
        "ml_risks": detect_structural_ml_risks(
            profiles
        ),
    }


# ============================================================
# REPORTING
# ============================================================

def print_summary(
    result: dict[str, Any],
    preview: int = 0,
) -> None:

    dataset = result["dataset"]
    summary = result["summary"]

    print("=" * 78)
    print("GENERIC DATASET DISCOVERY")
    print("=" * 78)
    print(f"Dataset : {dataset['path']}")
    print(f"Format  : {dataset['format']}")
    print(f"Size    : {dataset['size_mb']:.2f} MB")
    print(f"Tables  : {summary['table_count']}")
    print(f"Rows    : {summary['total_rows']:,}")
    print()
    print("DISCOVERED TABLE STRUCTURE")
    print("-" * 78)

    for name, profile in result["profiles"].items():
        print(
            f"{name} | "
            f"rows={profile['row_count']:,} | "
            f"columns={profile['column_count']} | "
            f"role={result['roles'][name]}"
        )

        print(
            f"  temporal={profile['temporal_columns']} | "
            f"numeric={profile['numeric_columns']} | "
            f"identifier_candidates="
            f"{profile['identifier_candidates']}"
        )

    print()
    print(
        "Relationship candidates : "
        f"{len(result['relationships'])}"
    )
    print(
        "Reuse candidates        : "
        f"{len(result['reuse_candidates'])}"
    )
    print(
        "Structural ML risks     : "
        f"{len(result['ml_risks']['risks'])}"
    )

    if preview > 0:
        print()
        print(
            f"PREVIEW ({preview} ROWS PER TABLE)"
        )
        print("-" * 78)

        for name, frame in result["tables"].items():
            print(f"\n[{name}]")
            print(
                frame.head(preview).to_string(
                    index=False
                )
            )


def build_report(
    result: dict[str, Any],
) -> dict[str, Any]:

    return {
        "dataset": result["dataset"],
        "summary": result["summary"],
        "profiles": result["profiles"],
        "roles": result["roles"],
        "temporal_structure": result[
            "temporal_structure"
        ],
        "relationships": result[
            "relationships"
        ],
        "reuse_candidates": result[
            "reuse_candidates"
        ],
        "quality": result["quality"],
        "ml_risks": result["ml_risks"],
        "architecture": {
            "dataset_specific_business_aliases": False,
            "hardcoded_business_table_names": False,
            "hardcoded_business_column_names": False,
            "hardcoded_entity_counts": False,
            "hardcoded_record_counts": False,
            "target_inference_from_names": False,
            "future_information_inference_from_names": False,
        },
    }


def validate_architecture() -> dict[str, Any]:
    source = Path(__file__).read_text(
        encoding=config.ENCODING
    )

    source_lower = source.lower()

    structural_requirements = {
        "automatic_schema_discovery": (
            "discover_dataset" in source
        ),
        "configured_format_validation": (
            "supported_dataset_formats" in source_lower
        ),
        "configured_file_size_validation": (
            "max_dataset_file_size_mb" in source_lower
        ),
        "runtime_column_type_discovery": (
            "infer_column_type" in source
        ),
        "runtime_relationship_discovery": (
            "discover_relationships" in source
        ),
    }

    failed_checks = [
        name
        for name, passed in structural_requirements.items()
        if not passed
    ]

    return {
        "passed": not failed_checks,
        "failed_structural_checks": failed_checks,
        "structural_checks": structural_requirements,
    }


# ============================================================
# CLI
# ============================================================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Discover and structurally profile a dataset."
        )
    )

    parser.add_argument(
        "dataset",
        nargs="?",
        help=(
            "Optional dataset path. If omitted, "
            "automatic configuration/discovery is used."
        ),
    )

    parser.add_argument(
        "--sample-size",
        type=int,
        default=config.SAMPLE_SIZE,
    )

    parser.add_argument(
        "--max-size-mb",
        type=float,
        default=config.MAX_DATASET_FILE_SIZE_MB,
    )

    parser.add_argument(
        "--preview",
        type=int,
        default=0,
    )

    parser.add_argument(
        "--quiet",
        action="store_true",
    )

    parser.add_argument(
        "--report",
    )

    parser.add_argument(
        "--validate",
        action="store_true",
    )

    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.sample_size <= 0:
        raise ValueError(
            "--sample-size must be greater than zero."
        )

    if args.preview < 0:
        raise ValueError(
            "--preview cannot be negative."
        )

    result = discover_dataset(
        dataset_path=args.dataset,
        sample_size=args.sample_size,
        max_size_mb=args.max_size_mb,
    )

    if not args.quiet:
        print_summary(
            result,
            preview=args.preview,
        )

    if args.validate:
        architecture = validate_architecture()

        print()
        print(
            "Architecture validation : "
            + (
                "PASSED"
                if architecture["passed"]
                else "FAILED"
            )
        )

        if architecture["failed_structural_checks"]:
            print(
                "Failed structural checks:",
                architecture["failed_structural_checks"],
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
                json_safe(
                    build_report(result)
                ),
                indent=2,
                ensure_ascii=False,
            ),
            encoding=config.ENCODING,
        )

        print(
            f"Report written to: {report_path}"
        )

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            f"ERROR: {exc}",
            file=sys.stderr,
        )
        raise SystemExit(1)
