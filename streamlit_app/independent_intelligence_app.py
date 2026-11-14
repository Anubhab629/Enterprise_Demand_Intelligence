"""
Independent Company Intelligence
-------------------------------
A self-contained Streamlit feature for Excel workbook exploration.

Design guarantees:
- No imports from the EDI project.
- No fixed company-name/ID vocabulary.
- Schema discovery is based on observed data structure, not column labels.
- Company/entity identification is confidence-scored and user-confirmable.
- Original worksheet data is displayed without analytical transformations.
- Charts are selected from observed data types/cardinality.
- One cautious, evidence-grounded recommendation per identified entity.
- Recommendation output contains exactly: Company Name, Recommendation.

Run standalone:
    streamlit run independent_intelligence_app.py

Embed in another Streamlit app:
    from independent_intelligence_app import render_independent_intelligence
    render_independent_intelligence()
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import math
import re
import json
import urllib.request
import urllib.error

import pandas as pd
import plotly.express as px
import streamlit as st


# ----------------------------- Workbook ingestion -----------------------------

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
MAX_CHARTS_PER_SHEET = 10

# ----------------------------- Local AI configuration --------------------------

OLLAMA_ENTITY_MODEL = "qwen3:0.6b"
OLLAMA_RECOMMENDATION_MODEL = "qwen3:0.6b"

OLLAMA_API_URL = "http://localhost:11434/api/generate"
OLLAMA_TIMEOUT_SECONDS = 90


def read_workbook(uploaded_file: Any) -> dict[str, pd.DataFrame]:
    """Read all worksheets. Do not alter, filter, or drop source records."""
    suffix = Path(getattr(uploaded_file, "name", "")).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError("Unsupported workbook type. Upload .xlsx, .xlsm, or .xls.")

    excel = pd.ExcelFile(uploaded_file)
    frames: dict[str, pd.DataFrame] = {}
    for sheet in excel.sheet_names:
        # Keep empty sheets too: they are part of the uploaded workbook structure.
        frames[str(sheet)] = pd.read_excel(excel, sheet_name=sheet)
    return frames


# ----------------------------- Ollama AI integration ---------------------------


# ----------------------------- Ollama AI integration ---------------------------

def _ollama_generate_json(
    prompt: str,
    model: str,
) -> dict | None:
    """Generate JSON through Ollama and preserve diagnostic errors."""

    payload = {
        "model": model,
        "prompt": prompt,
        "format": "json",
        "stream": False,
        "keep_alive": "10m",
        "options": {
            "temperature": 0,
            "num_predict": 180,
            "num_ctx": 2048,
        },
    }

    request = urllib.request.Request(
        OLLAMA_API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(
            request,
            timeout=OLLAMA_TIMEOUT_SECONDS,
        ) as response:
            result = json.loads(response.read().decode("utf-8"))

        raw = result.get("response", "")

        if not raw:
            st.session_state["independent_ai_debug"] = (
                "Ollama returned an empty response."
            )
            return None

        parsed = json.loads(raw)

        if not isinstance(parsed, dict):
            st.session_state["independent_ai_debug"] = (
                "Ollama response was valid JSON but not a JSON object."
            )
            return None

        st.session_state["independent_ai_debug"] = (
            "Ollama response received and JSON parsed successfully."
        )

        return parsed

    except urllib.error.HTTPError as exc:
        try:
            details = exc.read().decode("utf-8", errors="replace")
        except Exception:
            details = str(exc)

        st.session_state["independent_ai_debug"] = (
            f"Ollama HTTP error {exc.code}: {details[:500]}"
        )

    except urllib.error.URLError as exc:
        st.session_state["independent_ai_debug"] = (
            f"Ollama connection error: {exc.reason}"
        )

    except TimeoutError:
        st.session_state["independent_ai_debug"] = (
            f"Ollama timed out after {OLLAMA_TIMEOUT_SECONDS} seconds."
        )

    except json.JSONDecodeError as exc:
        st.session_state["independent_ai_debug"] = (
            f"JSON parsing failed: {exc}"
        )

    except Exception as exc:
        st.session_state["independent_ai_debug"] = (
            f"{type(exc).__name__}: {exc}"
        )

    return None



def _build_entity_detection_profile(
    frame: pd.DataFrame,
) -> tuple[list[dict], set[str]]:
    """
    Profile every dataset column using observed values and structure.

    No columns are excluded based on predefined names or prefixes.
    """
    profiles = []
    eligible_columns = set()
    row_count = len(frame)

    for column in frame.columns:
        name = str(column)
        series = clean_entity_values(frame[column])

        values = series.dropna().astype(str).str.strip()
        values = values[values.ne("")]

        non_null = len(values)
        unique = values.nunique()

        kind, _ = _infer_kind(series)

        counts = values.value_counts()
        repeated_groups = int((counts >= 2).sum())

        is_candidate = (
            kind in {"categorical", "numeric"}
            and non_null >= 3
            and unique >= 2
            and unique <= max(2, min(500, int(non_null * 0.80)))
            and repeated_groups >= 2
        )

        if is_candidate:
            eligible_columns.add(name)

        profiles.append({
            "column": name,
            "data_type": str(frame[column].dtype),
            "inferred_kind": kind,
            "record_count": row_count,
            "non_null_count": non_null,
            "unique_count": int(unique),
            "uniqueness_ratio": round(
                unique / max(non_null, 1), 4
            ),
            "repeated_groups": repeated_groups,
            "eligible_grouping_candidate": is_candidate,
            "representative_values": (
                values.drop_duplicates().head(12).tolist()
            ),
        })

    return profiles, eligible_columns


def _ollama_detect_entity_column(
    frame: pd.DataFrame,
) -> str | None:
    """
    Identify the company/organization column through AI interpretation
    followed by independent Python validation.
    """
    profiles, eligible_columns = _build_entity_detection_profile(frame)

    if not eligible_columns:
        st.session_state["independent_ai_debug"] = (
            "Python structural profiling found no eligible "
            "repeated-value grouping candidates."
        )
        return None

    prompt = f"""
/no_think

You are an enterprise dataset schema interpretation system.

Identify the column that represents the company or organization
responsible for the records and can group those records by entity.

Inspect every supplied column profile.

RULES:

1. Do not assume every dataset contains a company column.
2. Do not invent column names or values.
3. Do not rely on a predefined company-name vocabulary.
4. Do not select a column merely because its values repeat.
5. Distinguish the organization responsible for records from
   products, categories, locations, customers, suppliers,
   employees, dates, transaction IDs and technical metadata.
6. Use column names, data types, representative values and
   structural evidence together.
7. Only select a column marked
   eligible_grouping_candidate=true.
8. If no suitable company/organization column exists,
   return null.
9. Return one best candidate only.
10. Return valid JSON without additional commentary.

DATASET PROFILES:

{json.dumps(profiles, ensure_ascii=False, default=str)}

Return exactly this structure:

{{
  "entity_column": "exact column name or null",
  "entity_type": "company, organization, or none",
  "confidence": 0.0,
  "reason": "brief explanation"
}}
"""

    result = _ollama_generate_json(
        prompt,
        OLLAMA_ENTITY_MODEL,
    )

    if not result:
        return None

    st.session_state["independent_ai_raw_result"] = result

    st.session_state["independent_ai_eligible_columns"] = sorted(
        eligible_columns
    )

    column = result.get("entity_column")
    entity_type = str(
        result.get("entity_type", "")
    ).casefold()

    try:
        confidence = float(result.get("confidence", 0))
    except (TypeError, ValueError):
        return None

    if not isinstance(column, str) or not column.strip():
        return None

    # The model cannot introduce a column that does not exist.
    if column not in eligible_columns:
        return None

    if entity_type not in {
        "company",
        "organization",
        "business",
    }:
        return None

    # The model's self-reported confidence is only one check.
    # It is not treated as proof of correctness.
    if not 0.70 <= confidence <= 1.0:
        return None

    values = (
        clean_entity_values(frame[column])
        .dropna()
        .astype(str)
        .str.strip()
    )

    values = values[values.ne("")]
    counts = values.value_counts()

    st.session_state["independent_ai_debug"] = (
        f"Validation details | "
        f"Column: {column!r} | "
        f"Eligible: {column in eligible_columns} | "
        f"Entity type: {entity_type!r} | "
        f"Confidence: {confidence} | "
        f"Populated values: {len(values)} | "
        f"Unique values: {values.nunique()} | "
        f"Repeated groups: {int((counts >= 2).sum())}"
    )

    if len(counts[counts >= 2]) < 2:
        st.session_state["independent_ai_debug"] = (
            "Selected column does not contain at least two "
            "repeated entity groups."
        )
        return None

    if len(values) < 3:
        st.session_state["independent_ai_debug"] = (
            "Selected column contains fewer than three populated values."
        )
        return None
    # All entity validation checks passed.
    st.session_state["independent_ai_debug"] = (
        f"Entity successfully validated: {column!r}"
    )

    return column


def _summarize_company_records(
    group: pd.DataFrame,
    full_frame: pd.DataFrame,
    entity_column: str,
    entity_value: str,
) -> dict:
    """
    Calculate analytical evidence using all available records
    belonging to one company.

    No predefined column names or metadata prefixes are excluded.
    """
    summary = {
        "company": entity_value,
        "total_records": int(len(group)),
        "total_columns": int(len(group.columns)),
        "columns": {},
    }

    for column in group.columns:
        name = str(column)

        if name == entity_column:
            continue

        series = group[column]
        non_null = series.dropna()
        missing_count = int(series.isna().sum())

        column_summary = {
            "data_type": str(series.dtype),
            "record_count": int(len(series)),
            "non_null_count": int(len(non_null)),
            "missing_count": missing_count,
            "missing_percentage": round(
                missing_count / max(len(series), 1) * 100,
                2,
            ),
            "unique_count": int(_safe_unique(series)),
        }

        kind, parsed_dates = _infer_kind(series)

        if kind == "numeric":
            values = pd.to_numeric(
                series,
                errors="coerce",
            ).dropna()

            if not values.empty:
                q1 = float(values.quantile(0.25))
                q3 = float(values.quantile(0.75))
                iqr = q3 - q1

                outlier_rate = 0.0

                if iqr > 0:
                    outliers = (
                        (values < q1 - 1.5 * iqr)
                        | (values > q3 + 1.5 * iqr)
                    )
                    outlier_rate = float(outliers.mean() * 100)

                column_summary["numeric_statistics"] = {
                    "minimum": float(values.min()),
                    "maximum": float(values.max()),
                    "mean": float(values.mean()),
                    "median": float(values.median()),
                    "standard_deviation": (
                        float(values.std())
                        if len(values) > 1 else 0.0
                    ),
                    "q1": q1,
                    "q3": q3,
                    "iqr_outlier_percentage": round(
                        outlier_rate, 2
                    ),
                }

                # Compare the company's median against the
                # median of the other identified companies.
                peer_values = pd.to_numeric(
                    full_frame[column],
                    errors="coerce",
                )

                entity_labels = (
                    full_frame[entity_column]
                    .astype("string")
                    .str.strip()
                )

                peer_data = pd.DataFrame({
                    "_entity": entity_labels,
                    "_value": peer_values,
                }).dropna()

                peer_medians = (
                    peer_data.groupby("_entity")["_value"]
                    .median()
                )

                other_peers = peer_medians.drop(
                    index=entity_value,
                    errors="ignore",
                )

                if not other_peers.empty:
                    column_summary["peer_median"] = float(
                        other_peers.median()
                    )

                    column_summary[
                        "difference_from_peer_median"
                    ] = float(
                        values.median() - other_peers.median()
                    )

        elif kind == "datetime" and parsed_dates is not None:
            dates = parsed_dates.dropna()

            if not dates.empty:
                column_summary["date_statistics"] = {
                    "earliest": str(dates.min()),
                    "latest": str(dates.max()),
                    "unique_dates": int(dates.nunique()),
                }

        else:
            text_values = non_null.astype(str).str.strip()
            frequencies = text_values.value_counts()

            column_summary["categorical_distribution"] = {
                "distinct_values": int(text_values.nunique()),
                "most_frequent_values": [
                    {
                        "value": str(value),
                        "count": int(count),
                    }
                    for value, count in frequencies.head(8).items()
                ],
            }

        summary["columns"][name] = column_summary

    if len(group) > 1:
        summary["possible_duplicate_record_percentage"] = round(
            float(
                group.duplicated(keep=False).mean() * 100
            ),
            2,
        )

    return summary


def _generate_ai_company_recommendation(
    entity_value: str,
    group: pd.DataFrame,
    full_frame: pd.DataFrame,
    entity_column: str,
) -> str | None:
    """
    Generate exactly one prioritized recommendation for one company.
    Record the precise reason whenever AI generation or validation fails.
    """

    diagnostics = st.session_state.setdefault(
        "independent_ai_recommendation_debug",
        {},
    )

    def fail(reason: str) -> None:
        diagnostics[entity_value] = reason
        st.session_state[
            "independent_ai_recommendation_debug"
        ] = diagnostics

    evidence = _summarize_company_records(
        group,
        full_frame,
        entity_column,
        entity_value,
    )

    prompt = f"""
/no_think

You are an enterprise business intelligence analyst.

Analyze the complete analytical profile of ONE company and
produce exactly ONE prioritized, actionable recommendation.

The profile was calculated by Python using ALL available
records belonging to this company.

RULES:

1. Use only the supplied analytical evidence.
2. Do not invent figures, trends, causes, business conditions,
   products or company details.
3. Do not assume higher or lower values are automatically better.
4. Do not confuse missing data with poor business performance.
5. Do not recommend an action unsupported by the evidence.
6. Prioritize the most important actionable finding.
7. If no material issue is supported, recommend sensible
   monitoring or validation.
8. Generate exactly ONE recommendation, not alternatives.
9. Write 2 to 4 professional, actionable sentences.
10. Make it specific to this company.
11. Evidence columns must be exact source column names.
12. Return valid JSON only.

COMPLETE COMPANY ANALYTICAL PROFILE:

{json.dumps(evidence, ensure_ascii=False, default=str)}

Return exactly this structure:

{{
  "recommendation": "One detailed actionable recommendation.",
  "evidence_columns": ["Exact source column names used"],
  "priority": "high, medium, or low"
}}
"""

    result = _ollama_generate_json(
        prompt,
        OLLAMA_RECOMMENDATION_MODEL,
    )

    if not result:
        reason = st.session_state.get(
            "independent_ai_debug",
            "Ollama did not return a valid JSON response.",
        )
        fail(f"AI generation failed: {reason}")
        return None

    recommendation = result.get("recommendation")
    evidence_columns = result.get("evidence_columns")
    priority = result.get("priority")

    if not isinstance(recommendation, str):
        fail("AI response validation failed: recommendation is not a string.")
        return None

    recommendation = recommendation.strip()

    if not recommendation:
        fail("AI response validation failed: recommendation is empty.")
        return None

    if len(recommendation) > 2000:
        fail("AI response validation failed: recommendation exceeds 2000 characters.")
        return None

    if not isinstance(evidence_columns, list) or not evidence_columns:
        fail("AI response validation failed: evidence_columns is missing or empty.")
        return None

    valid_columns = {
        str(column)
        for column in group.columns
        if str(column) != entity_column
    }

    invalid_columns = [
        column
        for column in evidence_columns
        if not isinstance(column, str) or column not in valid_columns
    ]

    if invalid_columns:
        fail(
            "AI response validation failed: invalid evidence columns: "
            + repr(invalid_columns)
        )
        return None

    if priority not in {"high", "medium", "low"}:
        fail(f"AI response validation failed: invalid priority {priority!r}.")
        return None

    diagnostics[entity_value] = (
        f"AI recommendation accepted. Priority: {priority}. "
        f"Evidence columns: {evidence_columns}"
    )

    st.session_state[
        "independent_ai_recommendation_debug"
    ] = diagnostics

    return recommendation



# ----------------------------- Schema discovery -------------------------------

@dataclass(frozen=True)
class ColumnProfile:
    name: str
    dtype: str
    non_null: int
    unique: int
    uniqueness_ratio: float
    inferred_kind: str
    entity_score: float
    entity_reason: str


def _safe_unique(series: pd.Series) -> int:
    try:
        return int(series.nunique(dropna=True))
    except (TypeError, ValueError):
        return int(series.dropna().astype(str).nunique())


def _infer_datetime(series: pd.Series) -> pd.Series | None:
    if pd.api.types.is_datetime64_any_dtype(series):
        parsed = pd.to_datetime(series, errors="coerce")
    elif (
        pd.api.types.is_object_dtype(series)
        or pd.api.types.is_string_dtype(series)
    ):
        sample = series.dropna()
        if sample.empty:
            return None
        # Avoid interpreting ordinary numeric identifiers as dates.
        if pd.api.types.is_numeric_dtype(sample):
            return None
        sample = sample.astype(str).head(500)
        try:
            parsed_sample = pd.to_datetime(sample, errors="coerce", format="mixed")
        except (TypeError, ValueError):
            parsed_sample = pd.to_datetime(sample, errors="coerce")
        if len(sample) == 0 or parsed_sample.notna().mean() < 0.85:
            return None
        try:
            parsed = pd.to_datetime(series, errors="coerce", format="mixed")
        except (TypeError, ValueError):
            parsed = pd.to_datetime(series, errors="coerce")
    else:
        return None

    if parsed.notna().sum() >= 3 and parsed.nunique(dropna=True) >= 2:
        return parsed
    return None


def _infer_kind(series: pd.Series) -> tuple[str, pd.Series | None]:
    if pd.api.types.is_bool_dtype(series):
        return "boolean", None
    if pd.api.types.is_numeric_dtype(series):
        return "numeric", None
    dates = _infer_datetime(series)
    if dates is not None:
        return "datetime", dates
    if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
        unique = _safe_unique(series)
        if unique <= 1:
            return "constant_or_empty", None
        return "categorical", None
    return "other", None


def profile_schema(frame: pd.DataFrame) -> list[ColumnProfile]:
    """
    Infer likely entity/grouping fields from structural evidence only.
    Column names are deliberately not used as a fixed keyword dictionary.
    """
    profiles: list[ColumnProfile] = []
    row_count = len(frame)

    for column in frame.columns:
        series = frame[column]
        non_null = int(series.notna().sum())
        unique = _safe_unique(series)
        ratio = unique / max(non_null, 1)
        kind, _ = _infer_kind(series)

        score = 0.0
        reasons: list[str] = []

        # A candidate entity key should normally be a repeated categorical value,
        # not a unique row identifier or a numeric measurement.
        if kind == "categorical":
            if 2 <= unique <= max(2, min(500, int(max(row_count, 1) * 0.5))):
                score += 0.35
                reasons.append("multiple distinct labels")
            if 0.01 <= ratio <= 0.50:
                score += 0.30
                reasons.append("repeated values across records")
            if non_null >= 4:
                score += 0.15
                reasons.append("sufficient populated records")
            if unique >= 2 and (non_null - unique) >= 2:
                score += 0.20
                reasons.append("most labels occur more than once")
        elif kind == "numeric":
            # Numeric categorical identifiers can be legitimate, but are weaker
            # candidates until a user confirms their meaning.
            if 2 <= unique <= min(100, max(2, int(max(row_count, 1) * 0.25))):
                score += 0.20
                reasons.append("low-cardinality numeric values")
                if non_null - unique >= 2:
                    score += 0.10
                    reasons.append("numeric values repeat")
        elif kind == "boolean":
            score += 0.05
            reasons.append("binary field; unlikely to identify entities alone")

        score = min(score, 1.0)
        profiles.append(
            ColumnProfile(
                name=str(column),
                dtype=str(series.dtype),
                non_null=non_null,
                unique=unique,
                uniqueness_ratio=round(ratio, 4),
                inferred_kind=kind,
                entity_score=round(score, 3),
                entity_reason=", ".join(reasons) if reasons else "no strong entity structure detected",
            )
        )

    return profiles

def clean_entity_values(series: pd.Series) -> pd.Series:
    """
    Exclude missing-value descriptions from entity identification.
    Preserve all genuine source values.
    """
    cleaned = series.copy()

    text_values = cleaned.astype("string").str.strip()

    invalid_markers = (
        text_values.str.casefold().str.startswith(
            "this field is not present in "
        )
        | text_values.str.casefold().isin({
            "not provided in the original record",
            "not provided",
            "nan",
            "none",
            "null",
            "n/a",
        })
    )

    cleaned = cleaned.mask(invalid_markers)
    cleaned = cleaned.mask(text_values.eq(""))

    return cleaned




def suggested_entity_column(frame: pd.DataFrame) -> Any | None:
    """
    Identify a company/organization grouping field dynamically.

    No predefined company column names, metadata conventions,
    company vocabularies, or dataset-specific assumptions.
    """
    if frame.empty or len(frame.columns) == 0:
        return None

    # Primary detection: local AI + Python validation.
    ai_column = _ollama_detect_entity_column(frame)

    if ai_column is not None:
        st.session_state[
            "independent_ai_entity_detection"
        ] = (
            "Identified using Qwen3 AI and "
            "Python structural validation."
        )
        return ai_column

    # Do not silently choose an arbitrary categorical column.
    st.session_state["independent_ai_entity_detection"] = (
        st.session_state.get(
            "independent_ai_debug",
            "Entity detection failed without a diagnostic.",
        )
    )

    return None


def get_usable_columns(
    frame: pd.DataFrame
) -> tuple[list[str], list[str], list[str]]:

    numeric = []
    categorical = []
    datetime_columns = []

    for column in frame.columns:

        name = str(column)

        # Exclude all technical provenance metadata.
        if name.startswith("__meta_"):
            continue

        kind, _ = _infer_kind(frame[column])

        if (
            kind == "numeric"
            and frame[column].notna().sum() >= 2
            and _safe_unique(frame[column]) > 1
        ):
            numeric.append(name)

        elif (
            kind == "categorical"
            and 1 < _safe_unique(frame[column]) <= 40
        ):
            categorical.append(name)

        elif kind == "datetime":
            datetime_columns.append(name)

    return numeric, categorical, datetime_columns


# ----------------------------- Recommendation engine --------------------------

def _missing_recommendation(
    group: pd.DataFrame,
    entity_column: str | None
) -> tuple[float, str] | None:

    columns = [
        column
        for column in group.columns
        if column != entity_column
        and not str(column).startswith("__meta_")
    ]

    if not columns or group.empty:
        return None

    rates = group[columns].isna().mean()

    if rates.empty:
        return None

    column = str(rates.idxmax())
    rate = float(rates[column])

    if rate >= 0.20:
        return rate, (
            f"Improve data completeness for '{column}': "
            f"{rate:.1%} of this company's records are missing "
            "this value. Investigate the source and collection process."
        )

    return None


def _duplicate_recommendation(
    group: pd.DataFrame
) -> tuple[float, str] | None:

    business_columns = [
        column
        for column in group.columns
        if not str(column).startswith("__meta_")
    ]

    if len(group) < 2 or not business_columns:
        return None

    business_data = group[business_columns]

    duplicate_rate = float(
        business_data.duplicated(keep=False).mean()
    )

    if duplicate_rate >= 0.10:
        return duplicate_rate, (
            f"Review possible duplicate records: "
            f"{duplicate_rate:.1%} of this company's business records "
            "match another complete record. Verify whether these "
            "are legitimate repeated transactions."
        )

    return None


def _outlier_recommendation(
    group: pd.DataFrame,
    numeric_columns: list[str],
) -> tuple[float, str] | None:
    best: tuple[float, str, float] | None = None
    for column in numeric_columns:
        values = pd.to_numeric(group[column], errors="coerce").dropna()
        if len(values) < 8 or values.nunique() < 4:
            continue
        median = float(values.median())
        mad = float((values - median).abs().median())
        if mad <= 0:
            continue
        robust_z = 0.6745 * (values - median).abs() / mad
        rate = float((robust_z > 3.5).mean())
        if rate >= 0.05 and (best is None or rate > best[0]):
            best = (rate, str(column), median)
    if best:
        rate, column, _ = best
        return rate, (
            f"Validate extreme values in '{column}': {rate:.1%} of this entity's "
            "available observations are statistically unusual relative to its own "
            "distribution. Check source records before taking operational action."
        )
    return None


def _peer_deviation_recommendation(
    entity_value: str,
    frame: pd.DataFrame,
    entity_column: str,
    numeric_columns: list[str],
) -> tuple[float, str] | None:
    if frame[entity_column].dropna().nunique() < 3:
        return None

    best: tuple[float, str, str] | None = None
    for column in numeric_columns:
        values = pd.to_numeric(frame[column], errors="coerce")
        entity_labels = frame[entity_column].astype("string").str.strip()
        grouped = (
            pd.DataFrame({"_entity": entity_labels, "_value": values})
            .dropna(subset=["_entity", "_value"])
            .groupby("_entity")["_value"]
            .median()
            .dropna()
        )
        if len(grouped) < 3 or entity_value not in grouped.index:
            continue

        own = float(grouped.loc[entity_value])
        peers = grouped.drop(index=entity_value, errors="ignore")
        if len(peers) < 2:
            continue
        peer_median = float(peers.median())
        peer_mad = float((peers - peer_median).abs().median())
        # A floor prevents tiny numerical noise from being labelled significant.
        scale = max(peer_mad * 1.4826, abs(peer_median) * 0.05, 1e-9)
        deviation = abs(own - peer_median) / scale
        if deviation >= 3.0 and (best is None or deviation > best[0]):
            direction = "higher" if own > peer_median else "lower"
            best = (deviation, str(column), direction)

    if best:
        deviation, column, direction = best
        return deviation, (
            f"Investigate '{column}': this entity's median is materially {direction} "
            "than the other entities in this workbook. Confirm that the records and "
            "measurement definitions are comparable before changing decisions."
        )
    return None


def recommend_for_entity(
    entity_value: str,
    group: pd.DataFrame,
    full_frame: pd.DataFrame,
    entity_column: str | None,
) -> str:
    """Choose one recommendation using transparent, conservative evidence priorities."""
    numeric_columns, _, _ = get_usable_columns(full_frame)

    candidates: list[tuple[float, int, str]] = []
    missing = _missing_recommendation(group, entity_column)
    if missing:
        candidates.append((missing[0], 4, missing[1]))

    duplicates = _duplicate_recommendation(group)
    if duplicates:
        candidates.append((duplicates[0], 3, duplicates[1]))

    outliers = _outlier_recommendation(group, numeric_columns)
    if outliers:
        candidates.append((outliers[0], 2, outliers[1]))

    if entity_column and entity_column in full_frame.columns:
        peer = _peer_deviation_recommendation(
            str(entity_value), full_frame, entity_column, numeric_columns
        )
        if peer:
            candidates.append((peer[0], 1, peer[1]))

    if candidates:
        # Priority favours data completeness, then duplicates, then unusual values,
        # then peer deviation. Within a priority, strongest signal wins.
        candidates.sort(key=lambda item: (item[1], item[0]), reverse=True)
        return candidates[0][2]

    if len(group) < 5:
        return (
            "Expand the available records before making company-level decisions; "
            "the current sample is too small for a dependable pattern assessment."
        )

    return (
        "No material data-quality issue or strong statistical deviation was detected "
        "by the available checks. Continue monitoring the recorded measures; this "
        "workbook alone does not justify a more specific action."
    )



def build_recommendations(
    frame: pd.DataFrame,
    entity_column: str | None,
    fallback_label: str,
) -> pd.DataFrame:
    """
    Generate exactly one recommendation per identified company.

    Primary:
        Full-record Python analysis + Qwen3 interpretation.

    Fallback:
        Existing Python recommendation engine.
    """
    output_columns = [
        "Company Name",
        "Recommendation",
    ]

    st.session_state[
        "independent_ai_recommendations_used"
    ] = 0

    st.session_state[
        "independent_python_fallback_recommendations"
    ] = 0

    if frame.empty:
        return pd.DataFrame(columns=output_columns)

    if entity_column and entity_column in frame.columns:
        entity_values = clean_entity_values(
            frame[entity_column]
        )

        labels = (
            entity_values.dropna()
            .astype(str)
            .str.strip()
        )

        unique_labels = list(
            dict.fromkeys(labels[labels.ne("")].tolist())
        )

        rows = []

        for label in unique_labels:
            mask = (
                entity_values.astype("string")
                .str.strip()
                .eq(label)
            )

            group = frame.loc[mask.fillna(False)]

            if group.empty:
                continue

            # Primary: AI recommendation based on the full
            # analytical profile of this company's records.
            recommendation = _generate_ai_company_recommendation(
                label,
                group,
                frame,
                entity_column,
            )

            if recommendation:
                st.session_state[
                    "independent_ai_recommendations_used"
                ] += 1

            else:
                # Fallback to the existing Python engine.
                recommendation = recommend_for_entity(
                    label,
                    group,
                    frame,
                    entity_column,
                )

                st.session_state[
                    "independent_python_fallback_recommendations"
                ] += 1

            rows.append({
                "Company Name": label,
                "Recommendation": recommendation,
            })

        return pd.DataFrame(
            rows,
            columns=output_columns,
        )

    # No entity identified: do not invent a company name.
    return pd.DataFrame(columns=output_columns)



# ---------------------------------- Charts -------------------------------------

def _candidate_geo_columns(frame: pd.DataFrame, numeric: list[str]) -> tuple[str, str] | None:
    """
    Infer a possible latitude/longitude pair from value ranges only.
    This is a candidate, not a guarantee; coordinate values are inherently ambiguous.
    """
    candidates = []
    for column in numeric:
        values = pd.to_numeric(frame[column], errors="coerce").dropna()
        if len(values) < 3:
            continue
        minimum, maximum = float(values.min()), float(values.max())
        if minimum >= -90 and maximum <= 90:
            candidates.append((column, "lat"))
        if minimum >= -180 and maximum <= 180:
            candidates.append((column, "lon"))

    for lat_column, lat_kind in candidates:
        if lat_kind != "lat":
            continue
        for lon_column, lon_kind in candidates:
            if lon_kind != "lon" or lon_column == lat_column:
                continue
            lat_values = pd.to_numeric(frame[lat_column], errors="coerce")
            lon_values = pd.to_numeric(frame[lon_column], errors="coerce")
            valid = lat_values.between(-90, 90) & lon_values.between(-180, 180)
            if int(valid.sum()) >= 3:
                return lat_column, lon_column
    return None


def render_adaptive_charts(frame: pd.DataFrame, sheet_name: str) -> None:
    """Create several chart families when the workbook contains suitable data."""
    if frame.empty or len(frame.columns) == 0:
        st.info("This worksheet has no records or columns to visualise.")
        return

    numeric, categorical, date_columns = get_usable_columns(frame)
    chart_count = 0
    template = "plotly_dark"

    # 1. Numeric distributions: histograms.
    for column in numeric[:2]:
        if chart_count >= MAX_CHARTS_PER_SHEET:
            break
        values = pd.to_numeric(frame[column], errors="coerce").dropna()
        if len(values) < 2:
            continue
        fig = px.histogram(
            x=values,
            nbins=min(40, max(10, int(math.sqrt(len(values))))),
            title=f"Distribution of {column}",
            labels={"x": column, "y": "Record count"},
            template=template,
        )
        st.plotly_chart(fig, use_container_width=True)
        chart_count += 1

    # 2. Category composition: donut and pie charts.
    for column in categorical[:1]:
        if chart_count >= MAX_CHARTS_PER_SHEET:
            break
        counts = frame[column].dropna().astype(str).value_counts().head(12)
        if len(counts) < 2:
            continue
        chart_data = counts.rename_axis(column).reset_index(name="Record count")
        fig = px.pie(
            chart_data,
            names=column,
            values="Record count",
            hole=0.48,
            title=f"Composition by {column} (Donut Chart)",
            template=template,
        )
        st.plotly_chart(fig, use_container_width=True)
        chart_count += 1

    # Also show a conventional pie chart for a second categorical field when
    # available; the donut above remains useful for the first field.
    if chart_count < MAX_CHARTS_PER_SHEET and len(categorical) > 1:
        column = categorical[1]
        counts = frame[column].dropna().astype(str).value_counts().head(10)
        if len(counts) >= 2:
            chart_data = counts.rename_axis(column).reset_index(name="Record count")
            fig = px.pie(
                chart_data,
                names=column,
                values="Record count",
                title=f"Proportion by {column} (Pie Chart)",
                template=template,
            )
            st.plotly_chart(fig, use_container_width=True)
            chart_count += 1

    # 6. Time-series: line charts only when a date-like field and numeric measure exist.
    if chart_count < MAX_CHARTS_PER_SHEET and numeric:
        for date_column in date_columns:
            parsed = _infer_datetime(frame[date_column])
            if parsed is None:
                continue
            measure = numeric[0]
            trend = pd.DataFrame({
                "_date": parsed,
                "_measure": pd.to_numeric(frame[measure], errors="coerce"),
            }).dropna()
            trend = (
                trend.groupby("_date", as_index=False)["_measure"]
                .mean()
                .sort_values("_date")
            )
            if len(trend) < 2:
                continue
            fig = px.line(
                trend,
                x="_date",
                y="_measure",
                markers=True,
                title=f"{measure} Trend Over Time",
                labels={"_date": date_column, "_measure": f"Average {measure}"},
                template=template,
            )
            st.plotly_chart(fig, use_container_width=True)
            chart_count += 1
            break

    # 7. Geographic scatter map when a plausible coordinate pair exists.
    if chart_count < MAX_CHARTS_PER_SHEET:
        geo_pair = _candidate_geo_columns(frame, numeric)
        if geo_pair:
            lat_col, lon_col = geo_pair
            geo_data = frame[[lat_col, lon_col]].copy()
            geo_data[lat_col] = pd.to_numeric(geo_data[lat_col], errors="coerce")
            geo_data[lon_col] = pd.to_numeric(geo_data[lon_col], errors="coerce")
            geo_data = geo_data.dropna()
            geo_data = geo_data[
                geo_data[lat_col].between(-90, 90)
                & geo_data[lon_col].between(-180, 180)
            ]
            if len(geo_data) >= 3:
                fig = px.scatter_geo(
                    geo_data,
                    lat=lat_col,
                    lon=lon_col,
                    title=f"Geographic Distribution ({lat_col}, {lon_col})",
                    template=template,
                )
                st.plotly_chart(fig, use_container_width=True)
                chart_count += 1

    # 3. Categorical comparison: horizontal bar charts.
    for category in categorical:
        if chart_count >= MAX_CHARTS_PER_SHEET:
            break
        counts = frame[category].dropna().astype(str).value_counts().head(15)
        if len(counts) < 2:
            continue
        chart_data = counts.rename_axis(category).reset_index(name="Record count")
        fig = px.bar(
            chart_data,
            x="Record count",
            y=category,
            orientation="h",
            title=f"Record Count by {category}",
            template=template,
        )
        st.plotly_chart(fig, use_container_width=True)
        chart_count += 1

    # 4. Numeric relationships: scatter plots.
    if chart_count < MAX_CHARTS_PER_SHEET and len(numeric) >= 2:
        x_col, y_col = numeric[0], numeric[1]
        scatter_data = frame[[x_col, y_col]].apply(pd.to_numeric, errors="coerce").dropna()
        if len(scatter_data) >= 3:
            fig = px.scatter(
                scatter_data,
                x=x_col,
                y=y_col,
                title=f"Relationship Between {x_col} and {y_col}",
                template=template,
            )
            st.plotly_chart(fig, use_container_width=True)
            chart_count += 1

    # 5. Numeric spread by category: box plots.
    if chart_count < MAX_CHARTS_PER_SHEET:
        for category in categorical:
            suitable = frame[category].dropna().astype(str).nunique()
            if not 2 <= suitable <= 12:
                continue
            for measure in numeric:
                plot_data = frame[[category, measure]].copy()
                plot_data[measure] = pd.to_numeric(plot_data[measure], errors="coerce")
                plot_data = plot_data.dropna()
                if len(plot_data) < 4:
                    continue
                fig = px.box(
                    plot_data,
                    x=category,
                    y=measure,
                    points="outliers",
                    title=f"Distribution of {measure} Across {category}",
                    template=template,
                )
                st.plotly_chart(fig, use_container_width=True)
                chart_count += 1
                break
            if chart_count >= MAX_CHARTS_PER_SHEET:
                break

    if chart_count == 0:
        st.info(
            "No meaningful chart could be generated from the detected data types. "
            "The original dataset is still available above."
        )


# ----------------------------------- UI ----------------------------------------

def render_independent_intelligence() -> None:
    """Render independent company intelligence."""

    st.title("Independent Company Intelligence")

    st.caption(
        "Upload an Excel workbook to inspect original worksheets, "
        "explore adaptive visualisations, and generate "
        "evidence-based entity recommendations."
    )

    uploaded = st.file_uploader(
        "Upload Excel workbook",
        type=["xlsx", "xlsm", "xls"],
        key="independent_intelligence_workbook",
        help="Every worksheet is read independently. "
             "Your uploaded workbook is not modified."
    )

    if uploaded is None:
        st.info("Upload an Excel workbook to begin.")
        return

    try:
        frames = read_workbook(uploaded)

    except Exception as exc:
        st.error(
            f"Unable to read this workbook: "
            f"{type(exc).__name__}: {exc}"
        )
        return

    if not frames:
        st.warning("No worksheets were found.")
        return

    # --------------------------------------------------------
    # 1. Select worksheet and identify the entity column
    # --------------------------------------------------------

    selected_sheet = st.selectbox(
        "Worksheet for analysis",
        list(frames.keys()),
        key="independent_intelligence_sheet"
    )

    frame = frames[selected_sheet]

    if frame.empty:
        st.warning("The selected worksheet contains no records.")
        return

    # --------------------------------------------------------
    # 2. Automatic entity identification and recommendations
    # --------------------------------------------------------

    st.subheader("Automatic Entity Analysis")

    entity_column = suggested_entity_column(frame)


    detection_status = st.session_state.get(
        "independent_ai_entity_detection",
        "Entity detection status unavailable.",
    )

    st.caption(f"Entity detection: {detection_status}")

    with st.expander("AI Detection Diagnostics"):

        st.write("Diagnostic message:")

        st.code(
            st.session_state.get(
                "independent_ai_debug",
                "No diagnostic information available.",
            )
        )

        st.write("Qwen3 raw JSON response:")

        st.json(
            st.session_state.get(
                "independent_ai_raw_result",
                {"status": "No response captured"},
            )
        )

        st.write("Python eligible columns:")

        st.write(
            st.session_state.get(
                "independent_ai_eligible_columns",
                [],
            )
        )


    if entity_column is not None:

        entity_values = clean_entity_values(
            frame[entity_column]
        )

        company_count = int(
            entity_values.nunique(dropna=True)
        )

        st.session_state["independent_ai_recommendations_used"] = 0

        recommendations = build_recommendations(
            frame,
            entity_column,
            ""
        )

        st.caption(
            f"Automatically identified entity field: {entity_column}. "
            "All available records belonging to each entity "
            "are analysed together."
        )

    else:

        company_count = 0

        recommendations = pd.DataFrame(
            columns=[
                "Company Name",
                "Recommendation"
            ]
        )

        st.warning(
            "The application could not reliably identify an entity "
            "grouping field in this dataset. Recommendations have not "
            "been fabricated. Check whether the uploaded data contains "
            "a repeated entity identifier."
        )

    
    if entity_column is not None:
        ai_count = st.session_state.get(
            "independent_ai_recommendations_used",
            0,
        )

        fallback_count = st.session_state.get(
            "independent_python_fallback_recommendations",
            0,
        )

        if ai_count:
            st.success(
                f"Qwen3 generated {ai_count} company-specific "
                "recommendation(s)."
            )

        if fallback_count:
            st.warning(
                f"{fallback_count} recommendation(s) used the "
                "existing Python fallback because AI generation "
                "was unavailable or failed validation."
            )

        with st.expander("AI Recommendation Diagnostics"):

            recommendation_debug = st.session_state.get(
                "independent_ai_recommendation_debug",
                {},
            )

            if recommendation_debug:
                for company, message in recommendation_debug.items():
                    st.markdown(f"**{company}**")
                    st.caption(message)
            else:
                st.info("No AI recommendation diagnostics available.")

    # --------------------------------------------------------
    # 3. KPI cards BEFORE uploaded dataset
    # --------------------------------------------------------

    st.divider()

    st.subheader("Dataset Overview")

    kpi1, kpi2, kpi3 = st.columns(3)

    kpi1.metric(
        "Total Dataset Records",
        f"{len(frame):,}"
    )

    kpi2.metric(
        "Total Companies / Entities",
        f"{company_count:,}" if company_count else "Not identified"
    )

    kpi3.metric(
        "Recommendations Generated",
        f"{len(recommendations):,}"
    )

    st.divider()

    # --------------------------------------------------------
    # 4. Original uploaded dataset
    # --------------------------------------------------------

    st.subheader("Uploaded Dataset")

    for sheet_name, sheet_frame in frames.items():

        with st.expander(
            f"{sheet_name} — "
            f"{len(sheet_frame):,} rows × "
            f"{len(sheet_frame.columns):,} columns",
            expanded=(sheet_name == selected_sheet)
        ):

            st.dataframe(
                sheet_frame,
                use_container_width=True,
                hide_index=True,
                height=560
            )

    st.divider()

    # --------------------------------------------------------
    # 5. Data visualisations
    # --------------------------------------------------------

    st.subheader("Data Visualisations")

    render_adaptive_charts(
        frame,
        selected_sheet
    )

    st.divider()

    # --------------------------------------------------------
    # 6. Company recommendations
    # --------------------------------------------------------

    st.subheader("Company Recommendations")

    st.caption(
        "Recommendations are generated separately for each "
        "identified entity using missing-value analysis, "
        "duplicate detection, statistical outlier detection "
        "and peer comparison where sufficient data exists."
    )

    if entity_column is None:

        st.info(
            "Company-level recommendations are unavailable "
            "because a reliable entity identifier could not be inferred."
        )

    elif recommendations.empty:

        st.info(
            "No valid entity records were available for analysis."
        )

    else:

        recommendations = recommendations[
            ["Company Name", "Recommendation"]
        ]

        # Dynamically sized recommendation panel.
        # Show more rows when more companies are identified.
        recommendation_panel_height = max(
            420,
            min(900, 140 * len(recommendations) + 80),
        )

        with st.container(
            height=recommendation_panel_height,
            border=True,
        ):

            for _, row in recommendations.iterrows():

                company_col, recommendation_col = st.columns(
                    [1, 3],
                    vertical_alignment="top"
                )

                with company_col:
                    st.markdown(
                        f"**{row['Company Name']}**"
                    )

                with recommendation_col:
                    st.write(
                        str(row["Recommendation"])
                    )

                st.divider()



def main() -> None:
    st.set_page_config(
        page_title="Independent Company Intelligence",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    render_independent_intelligence()


if __name__ == "__main__":
    main()
