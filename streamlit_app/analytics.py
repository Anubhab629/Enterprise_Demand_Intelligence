"""Dataset-independent analytics helpers for the EDI Streamlit dashboard.

These functions discover data characteristics dynamically. They do not assume
business-specific table names, column names, KPI definitions, or forecast fields.
All transformations operate on copies and preserve source DataFrames.
"""

from __future__ import annotations

import ast
import json
import re
from collections.abc import Mapping
from typing import Any

import pandas as pd


_NESTED_TYPES = (dict, list, tuple, set, frozenset)

_IDENTIFIER_NAME_PATTERN = re.compile(
    r"(^|[_.\-\s])"
    r"(id|key|uuid|guid|index|identifier|code|reference|ref|"
    r"record_id|row_id|customer_id|product_id|order_id|"
    r"transaction_id|recommendation_id)"
    r"($|[_.\-\s])",
    re.IGNORECASE,
)

_DATE_NAME_PATTERN = re.compile(
    r"(^|[_.\-\s])"
    r"(date|datetime|timestamp|time|created_at|updated_at|"
    r"event_date|order_date|transaction_date|period|month|year)"
    r"($|[_.\-\s])",
    re.IGNORECASE,
)

_DATE_VALUE_PATTERN = re.compile(
    r"^\s*(?:"
    r"\d{4}[-/]\d{1,2}[-/]\d{1,2}"
    r"|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}"
    r"|\d{4}-\d{2}"
    r"|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}"
    r")"
    r"(?:[T\s].*)?\s*$"
)


def _is_nested(value: Any) -> bool:
    """Return True when a cell contains a nested Python container."""
    return isinstance(value, _NESTED_TYPES)


def _is_missing_scalar(value: Any) -> bool:
    """Check missingness without treating nested containers as arrays."""
    if _is_nested(value):
        return False

    try:
        result = pd.isna(value)
        return bool(result) if not hasattr(result, "__len__") else False
    except (TypeError, ValueError):
        return False


def _canonical_cell(value: Any) -> str:
    """Create a stable text representation for nested or unhashable values."""
    try:
        return json.dumps(
            value,
            sort_keys=True,
            ensure_ascii=False,
            separators=(",", ":"),
            default=str,
        )
    except (TypeError, ValueError):
        return repr(value)


def _parse_nested_string(value: Any) -> Any:
    """Safely parse JSON/Python-literal containers, including quoted containers."""
    if not isinstance(value, str):
        return value

    candidate = value.strip()
    if len(candidate) < 2:
        return value

    # Some Excel exports wrap a dictionary string in an additional quote layer.
    if candidate[0] in ('"', "'") and candidate[-1] == candidate[0]:
        try:
            unwrapped = ast.literal_eval(candidate)
        except (SyntaxError, ValueError, TypeError, MemoryError, RecursionError):
            unwrapped = value

        if isinstance(unwrapped, str) and unwrapped.strip() != candidate:
            nested = _parse_nested_string(unwrapped)
            if isinstance(nested, (dict, list, tuple, set, frozenset)):
                return nested

    if candidate[0] not in "{[(" or candidate[-1] not in "}])":
        return value

    try:
        parsed = json.loads(candidate)
        if isinstance(parsed, (dict, list, tuple)):
            return parsed
    except (json.JSONDecodeError, TypeError, ValueError):
        pass

    try:
        parsed = ast.literal_eval(candidate)
    except (SyntaxError, ValueError, TypeError, MemoryError, RecursionError):
        return value

    return (
        parsed
        if isinstance(parsed, (dict, list, tuple, set, frozenset))
        else value
    )


def _normalise_nested_value(
    value: Any,
    *,
    parse_json_strings: bool,
) -> Any:
    """Normalize supported nested values for consistent expansion."""
    if parse_json_strings:
        value = _parse_nested_string(value)

    if isinstance(value, tuple):
        return list(value)

    if isinstance(value, (set, frozenset)):
        return sorted(value, key=_canonical_cell)

    return value


def _unique_column_name(name: str, used: set[str]) -> str:
    """Prevent generated flattened names from overwriting other columns."""
    candidate = name
    suffix = 2

    while candidate in used:
        candidate = f"{name}_{suffix}"
        suffix += 1

    used.add(candidate)
    return candidate


def expand_nested_columns(
    frame: pd.DataFrame,
    *,
    max_depth: int = 6,
    separator: str = ".",
    parse_json_strings: bool = True,
) -> pd.DataFrame:
    """Flatten nested dictionaries and quoted JSON without changing row count."""
    if not isinstance(frame, pd.DataFrame):
        raise TypeError("expand_nested_columns expects a pandas DataFrame")
    if frame.empty or len(frame.columns) == 0:
        return frame.copy(deep=True)

    depth_limit = max(0, int(max_depth))
    separator = str(separator) if separator else "."
    source = frame.copy(deep=True)

    def parse_value(value: Any) -> Any:
        return _normalise_nested_value(
            value,
            parse_json_strings=parse_json_strings,
        )

    def readable(value: Any) -> Any:
        if _is_nested(value):
            try:
                return json.dumps(
                    value,
                    ensure_ascii=False,
                    default=str,
                    sort_keys=True,
                )
            except (TypeError, ValueError):
                return repr(value)
        return value

    def unique_child_name(name: str, used: set[str]) -> str:
        return _unique_column_name(name, used)

    def expand_series(
        series: pd.Series,
        prefix: str,
        depth: int,
        used: set[str],
    ) -> dict[str, pd.Series]:
        values = series.map(parse_value)
        present = values[~values.map(_is_missing_scalar)]

        if depth >= depth_limit or present.empty:
            return {prefix: values.map(readable)}

        dict_values = present[
            present.map(lambda value: isinstance(value, Mapping))
        ]
        list_values = present[
            present.map(lambda value: isinstance(value, list))
        ]
        result: dict[str, pd.Series] = {}

        # Expand dictionary cells even if other rows contain ordinary text.
        if not dict_values.empty:
            keys: list[Any] = []
            seen: set[str] = set()

            for item in dict_values:
                for key in item:
                    text_key = str(key)
                    if text_key not in seen:
                        seen.add(text_key)
                        keys.append(key)

            scalar_values = values.map(
                lambda item: (
                    item
                    if not isinstance(item, Mapping)
                    and not isinstance(item, list)
                    else pd.NA
                )
            )
            if scalar_values.map(
                lambda value: not _is_missing_scalar(value)
            ).any():
                result[prefix] = scalar_values.map(readable)

            for key in keys:
                child_name = unique_child_name(
                    f"{prefix}{separator}{key}",
                    used,
                )
                child = values.map(
                    lambda item, k=key: (
                        item.get(k, pd.NA)
                        if isinstance(item, Mapping)
                        else pd.NA
                    )
                )
                result.update(
                    expand_series(
                        child,
                        child_name,
                        depth + 1,
                        used,
                    )
                )

        # Expand lists of objects by index, preserving the original row grain.
        if not list_values.empty and list_values.map(
            lambda items: all(isinstance(item, Mapping) for item in items)
        ).all():
            max_items = min(
                max((len(items) for items in list_values), default=0),
                20,
            )

            for item_index in range(max_items):
                item_series = values.map(
                    lambda items, i=item_index: (
                        items[i]
                        if isinstance(items, list) and i < len(items)
                        else pd.NA
                    )
                )
                child_prefix = unique_child_name(
                    f"{prefix}{separator}{item_index + 1}",
                    used,
                )
                result.update(
                    expand_series(
                        item_series,
                        child_prefix,
                        depth + 1,
                        used,
                    )
                )

            other_values = values.map(
                lambda item: (
                    item
                    if not isinstance(item, list)
                    and not isinstance(item, Mapping)
                    else pd.NA
                )
            )
            if other_values.map(
                lambda value: not _is_missing_scalar(value)
            ).any():
                result[prefix] = other_values.map(readable)

        elif not dict_values.empty:
            # Keep mixed list cells readable rather than silently discarding them.
            mixed_lists = values.map(
                lambda item: item if isinstance(item, list) else pd.NA
            )
            if mixed_lists.map(
                lambda value: not _is_missing_scalar(value)
            ).any():
                result[prefix] = mixed_lists.map(readable)

        elif dict_values.empty:
            # Primitive or heterogeneous lists are not exploded into false records.
            return {prefix: values.map(readable)}

        return result or {prefix: values.map(readable)}

    used_names = {str(column) for column in source.columns}
    output: dict[str, pd.Series] = {}

    for column in source.columns:
        name = str(column)
        expanded = expand_series(
            source[column],
            name,
            0,
            used_names,
        )

        for expanded_name, values in expanded.items():
            final_name = expanded_name
            if final_name in output:
                final_name = unique_child_name(final_name, used_names)
            output[final_name] = values.reindex(source.index)

    result = pd.DataFrame(output, index=source.index)
    if len(result) != len(source):
        raise RuntimeError(
            "Nested expansion unexpectedly changed the record count"
        )

    return result


def _comparable_frame(frame: pd.DataFrame) -> pd.DataFrame:
    """Normalize nested cells in a copy for duplicate-row detection."""
    comparable = frame.copy(deep=True)

    for column in comparable.columns:
        try:
            has_nested = comparable[column].map(_is_nested).any()
        except (TypeError, ValueError):
            has_nested = True

        if has_nested:
            comparable[column] = comparable[column].map(
                lambda value: (
                    _canonical_cell(value) if _is_nested(value) else value
                )
            )

    return comparable


def dataset_metrics(
    tables: Mapping[str, pd.DataFrame] | None,
) -> dict[str, int]:
    """Calculate aggregate metrics for a collection of DataFrames."""
    frames = [
        frame
        for frame in (tables or {}).values()
        if isinstance(frame, pd.DataFrame)
    ]

    duplicate_rows = 0

    for frame in frames:
        if frame.empty or len(frame.columns) == 0:
            continue

        try:
            duplicate_rows += int(
                _comparable_frame(frame).duplicated().sum()
            )
        except (TypeError, ValueError):
            continue

    return {
        "tables": len(frames),
        "rows": int(sum(len(frame.index) for frame in frames)),
        "columns": int(sum(len(frame.columns) for frame in frames)),
        "missing_cells": int(
            sum(int(frame.isna().sum().sum()) for frame in frames)
        ),
        "duplicate_rows": duplicate_rows,
    }


def numeric_columns(frame: pd.DataFrame) -> list[Any]:
    """Return all numeric columns except Boolean and complex values."""
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return []

    return [
        column
        for column in frame.columns
        if pd.api.types.is_numeric_dtype(frame[column].dtype)
        and not pd.api.types.is_bool_dtype(frame[column].dtype)
        and not pd.api.types.is_complex_dtype(frame[column].dtype)
    ]


def identifier_columns(frame: pd.DataFrame) -> list[Any]:
    """Identify likely record identifiers without deleting source columns.

    Name-based matches are treated as identifiers. Numeric fields with nearly
    unique values are also flagged when they resemble row keys. This is a
    heuristic for chart selection, not proof that a field is an identifier.
    """
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return []

    found: list[Any] = []
    row_count = len(frame)

    for column in frame.columns:
        name = str(column).strip()

        if _IDENTIFIER_NAME_PATTERN.search(name):
            found.append(column)
            continue

        series = frame[column].dropna()

        if (
            len(series) >= 10
            and pd.api.types.is_numeric_dtype(frame[column].dtype)
            and not pd.api.types.is_bool_dtype(frame[column].dtype)
        ):
            unique_count = int(series.nunique(dropna=True))
            unique_ratio = unique_count / max(1, len(series))
            normalized_name = re.sub(r"[^a-z0-9]", "", name.casefold())

            # High uniqueness alone is not enough: require a key-like name.
            key_like = any(
                token in normalized_name
                for token in ("record", "row", "serial", "sequence", "number")
            )
            if key_like and unique_ratio >= 0.9:
                found.append(column)

    return found


def numeric_measure_columns(frame: pd.DataFrame) -> list[Any]:
    """Return numeric fields suitable for generic measure charts."""
    excluded = set(identifier_columns(frame))
    return [
        column
        for column in numeric_columns(frame)
        if column not in excluded
    ]


def boolean_columns(frame: pd.DataFrame) -> list[Any]:
    """Discover Boolean columns and simple Boolean-like categorical fields."""
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return []

    found: list[Any] = []
    boolean_tokens = {"true", "false", "yes", "no", "y", "n"}

    for column in frame.columns:
        series = frame[column]

        if pd.api.types.is_bool_dtype(series.dtype):
            found.append(column)
            continue

        non_null = series.dropna()

        if (
            non_null.empty
            or pd.api.types.is_numeric_dtype(series.dtype)
            or non_null.map(_is_nested).any()
        ):
            continue

        values = {
            str(value).strip().casefold()
            for value in non_null.unique()
        }

        if len(values) == 2 and values.issubset(boolean_tokens):
            found.append(column)

    return found


def datetime_columns(frame: pd.DataFrame) -> list[Any]:
    """Discover datetime fields using types and conservative value parsing."""
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return []

    found: list[Any] = []

    for column in frame.columns:
        series = frame[column]

        if pd.api.types.is_datetime64_any_dtype(series.dtype):
            found.append(column)
            continue

        if (
            pd.api.types.is_numeric_dtype(series.dtype)
            or pd.api.types.is_bool_dtype(series.dtype)
        ):
            continue

        non_null = series.dropna()

        if non_null.empty or non_null.map(_is_nested).any():
            continue

        sample = non_null.astype(str).head(200)
        name_suggests_date = bool(_DATE_NAME_PATTERN.search(str(column)))
        values_look_like_dates = sample.str.match(
            _DATE_VALUE_PATTERN, na=False
        )

        # Without a temporal-looking column name, require date-shaped values.
        if not name_suggests_date and (
            len(values_look_like_dates) == 0
            or float(values_look_like_dates.mean()) < 0.8
        ):
            continue

        try:
            parsed = pd.to_datetime(
                sample,
                errors="coerce",
                format="mixed",
            )
        except (TypeError, ValueError):
            parsed = pd.to_datetime(sample, errors="coerce")

        ratio = float(parsed.notna().mean()) if len(parsed) else 0.0

        if ratio >= 0.9 and int(parsed.notna().sum()) >= 2:
            found.append(column)

    return found


def categorical_columns(
    frame: pd.DataFrame,
    max_unique: int = 50,
) -> list[Any]:
    """Discover low-cardinality fields suitable for category comparisons."""
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return []

    temporal = set(datetime_columns(frame))
    identifiers = set(identifier_columns(frame))
    found: list[Any] = []

    for column in frame.columns:
        series = frame[column]

        if (
            column in temporal
            or column in identifiers
            or pd.api.types.is_numeric_dtype(series.dtype)
            or pd.api.types.is_bool_dtype(series.dtype)
        ):
            continue

        non_null = series.dropna()

        if non_null.empty or non_null.map(_is_nested).any():
            continue

        # Avoid turning long free-text fields into category charts.
        text_lengths = non_null.astype(str).str.len()
        if len(text_lengths) and float(text_lengths.median()) > 100:
            continue

        try:
            unique_count = int(non_null.nunique(dropna=True))
        except (TypeError, ValueError):
            unique_count = int(non_null.map(repr).nunique())

        if 1 < unique_count <= max(1, int(max_unique)):
            found.append(column)

    return found


def numeric_summary(frame: pd.DataFrame) -> pd.DataFrame:
    """Describe discovered numeric fields, or return an empty DataFrame."""
    columns = numeric_columns(frame)

    if not columns:
        return pd.DataFrame()

    return (
        frame[columns]
        .describe()
        .transpose()
        .rename_axis("column")
        .reset_index()
    )


def numeric_distribution(
    frame: pd.DataFrame,
    column: Any,
) -> pd.Series:
    """Return non-missing numeric observations for a supported column."""
    if (
        column not in frame.columns
        or column not in numeric_columns(frame)
    ):
        return pd.Series(dtype="float64", name=column)

    return pd.to_numeric(frame[column], errors="coerce").dropna()


def categorical_counts(
    frame: pd.DataFrame,
    column: Any,
) -> pd.DataFrame:
    """Count observed values in a supported categorical field."""
    if column not in frame.columns:
        return pd.DataFrame(columns=["category", "count"])

    series = frame[column].dropna()

    if series.map(_is_nested).any():
        return pd.DataFrame(columns=["category", "count"])

    try:
        counts = series.astype(str).value_counts(dropna=False)
    except (TypeError, ValueError):
        counts = series.map(repr).value_counts(dropna=False)

    return counts.rename_axis("category").reset_index(name="count")


def schema_profile(frame: pd.DataFrame) -> pd.DataFrame:
    """Return observed schema and data-quality facts for each column."""
    records: list[dict[str, Any]] = []

    for column in frame.columns:
        series = frame[column]
        non_null = series.dropna()

        try:
            if non_null.map(_is_nested).any():
                distinct = int(non_null.map(_canonical_cell).nunique())
            else:
                distinct = int(non_null.nunique(dropna=True))
        except (TypeError, ValueError):
            distinct = int(non_null.map(repr).nunique())

        records.append({
            "Column": str(column),
            "Data type": str(series.dtype),
            "Missing values": int(series.isna().sum()),
            "Distinct values": distinct,
        })

    return pd.DataFrame(
        records,
        columns=[
            "Column",
            "Data type",
            "Missing values",
            "Distinct values",
        ],
    )


def format_nested_values(frame: pd.DataFrame) -> pd.DataFrame:
    """Serialize remaining nested cells for display on a copy."""
    formatted = frame.copy(deep=True)

    for column in formatted.columns:
        formatted[column] = formatted[column].map(
            lambda value: (
                json.dumps(
                    value,
                    ensure_ascii=False,
                    indent=2,
                    default=str,
                )
                if _is_nested(value)
                else value
            )
        )

    return formatted