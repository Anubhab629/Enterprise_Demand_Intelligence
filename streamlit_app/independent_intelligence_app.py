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
from geopy.geocoders import Nominatim 
from geopy.extra.rate_limiter import RateLimiter
import math
import re
import json
import hashlib
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
OLLAMA_RECOMMENDATION_MODEL = "qwen3:1.7b"

OLLAMA_API_URL = "http://localhost:11434/api/generate"
OLLAMA_TIMEOUT_SECONDS = 150


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
    generation_options: dict | None = None,
    timeout_seconds: int | None = None,
) -> dict | None:
    """Generate JSON through Ollama and preserve diagnostic errors."""

    payload = {
        "model": model,
        "prompt": prompt,
        "format": "json",
        "think": False,
        "stream": False,
        "keep_alive": "10m",
        "options": generation_options or {
            "temperature": 0,
            "num_predict": 768,
            "num_ctx": 4096,
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
            timeout=(
                timeout_seconds
                if timeout_seconds is not None
                else OLLAMA_TIMEOUT_SECONDS
            ),
        ) as response:
            result = json.loads(response.read().decode("utf-8"))

        raw = result.get("response", "")
        raw = raw if isinstance(raw, str) else ""

        # Keep the server metadata even when the model returns no text.
        metadata = {
            key: result.get(key)
            for key in (
                "model", "done", "done_reason", "prompt_eval_count",
                "eval_count", "total_duration", "load_duration",
                "prompt_eval_duration", "eval_duration"
            )
        }

        if not raw.strip():
            st.session_state["independent_ai_raw_result"] = {
                "status": "Ollama returned an empty response",
                "metadata": metadata,
            }
            st.session_state["independent_ai_debug"] = (
                "Ollama returned an empty response. "
                f"Server metadata: {metadata}"
            )
            return None

        if str(result.get("done_reason", "")).casefold() == "length":
            st.session_state["independent_ai_raw_result"] = {
                "status": "Ollama generation stopped at the token limit",
                "metadata": metadata,
                "response": raw[:4000],
            }
            st.session_state["independent_ai_debug"] = (
                "Ollama generation reached num_predict before completing. "
                "The response is treated as a failure so Python fallback runs. "
                f"Server metadata: {metadata}"
            )
            return None

        st.session_state["independent_ai_raw_result"] = {
            "status": "Response received",
            "metadata": metadata,
            "response": raw[:4000],
        }

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
        effective_timeout = (
            timeout_seconds
            if timeout_seconds is not None
            else OLLAMA_TIMEOUT_SECONDS
        )

        st.session_state["independent_ai_debug"] = (
            f"Ollama timed out after {effective_timeout} seconds."
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
    Profile every column from observed structure.

    Candidate screening is deliberately permissive: it only checks whether
    values can form repeated groups. It does not infer business meaning.
    Ollama interprets the full profiles; Python validates the selected field.
    No fixed column names, industry vocabulary, or dataset-specific rules.
    """
    profiles: list[dict] = []
    eligible_columns: set[str] = set()
    row_count = len(frame)

    for column in frame.columns:
        name = str(column)
        series = clean_entity_values(frame[column])
        values = series.dropna().astype(str).str.strip()
        values = values[values.ne("")]

        non_null = int(len(values))
        unique = int(values.nunique())
        kind, _ = _infer_kind(series)
        counts = values.value_counts()
        repeated_groups = int((counts >= 2).sum())
        repeated_records = int(counts[counts >= 2].sum())

        # This is only a structural pre-screen, not proof that the field
        # represents a company. Do not impose arbitrary cardinality caps.
        is_candidate = (
            kind in {"categorical", "numeric"}
            and non_null >= 4
            and unique >= 2
            and repeated_groups >= 2
        )

        if is_candidate:
            eligible_columns.add(name)

        examples = [str(v)[:100] for v in values.drop_duplicates().head(8).tolist()]
        profiles.append({
            "column": name,
            "data_type": str(frame[column].dtype),
            "inferred_kind": kind,
            "record_count": int(row_count),
            "non_null_count": non_null,
            "missing_count": int(row_count - non_null),
            "unique_count": unique,
            "uniqueness_ratio": round(unique / max(non_null, 1), 4),
            "repeated_groups": repeated_groups,
            "records_in_repeated_groups": repeated_records,
            "eligible_grouping_candidate": bool(is_candidate),
            "representative_values": examples,
        })

    # Save this before any network/model call so diagnostics remain truthful
    # even if Ollama is unavailable or returns an empty response.
    st.session_state["independent_ai_eligible_columns"] = sorted(eligible_columns)
    st.session_state["independent_ai_column_profiles"] = profiles
    return profiles, eligible_columns

def _ollama_detect_entity_column(
    frame: pd.DataFrame,
) -> str | None:
    """Ask Ollama to interpret runtime schema profiles, then validate in Python."""
    st.session_state["independent_ai_debug"] = "Entity detection started."
    st.session_state["independent_ai_raw_result"] = {
        "status": "No Ollama response received yet"
    }
    st.session_state["independent_ai_eligible_columns"] = []
    st.session_state["independent_ai_column_profiles"] = []

    if frame.empty or len(frame.columns) == 0:
        st.session_state["independent_ai_debug"] = "Selected worksheet has no data columns or records."
        return None

    profiles, eligible_columns = _build_entity_detection_profile(frame)

    # Even when the structural pre-screen finds no candidate, still attempt
    # Ollama using the complete runtime profiles. Python fallback will run
    # afterward if AI cannot identify a validated field.

    # Send compact profiles for every column, not only candidates. This lets
    # the model compare possible entity fields against dates, measures, IDs,
    # locations, products, customers, and other competing dimensions.
    prompt_profiles = [
        {
            "column": p["column"],
            "data_type": p["data_type"],
            "inferred_kind": p["inferred_kind"],
            "record_count": p["record_count"],
            "non_null_count": p["non_null_count"],
            "unique_count": p["unique_count"],
            "uniqueness_ratio": p["uniqueness_ratio"],
            "repeated_groups": p["repeated_groups"],
            "eligible_grouping_candidate": p["eligible_grouping_candidate"],
            "representative_values": p["representative_values"],
        }
        for p in profiles
    ]

    prompt = f"""
/no_think
You are a cautious enterprise data-schema interpreter.

Task: identify whether one column represents the organization/company
that owns, operates, or is the subject of the records, and can group
multiple records for that entity.

Important distinctions:
- A repeated value is not automatically a company.
- Customer, supplier, employee, product, category, region, location,
  date, status, and transaction identifiers are not the target unless
  the supplied schema evidence clearly establishes that role.
- Do not assume any particular industry, vocabulary, or column naming style.
- Do not invent a column name. Use an exact supplied column name.
- Select only a column with eligible_grouping_candidate=true.
- If the evidence is ambiguous or no organization field exists, return null.
- Confidence must reflect evidence quality, not certainty theater.

Return JSON only, using this exact structure:
{{
  "entity_column": "exact supplied column name or null",
  "entity_type": "company, organization, business, or none",
  "confidence": 0.0,
  "reason": "brief evidence-based explanation"
}}

RUNTIME COLUMN PROFILES:
{json.dumps(prompt_profiles, ensure_ascii=False, default=str)}
"""

    result = _ollama_generate_json(prompt, OLLAMA_ENTITY_MODEL)
    if not result:
        return None

    column = result.get("entity_column")
    entity_type = str(result.get("entity_type", "")).strip().casefold()
    try:
        confidence = float(result.get("confidence", 0))
    except (TypeError, ValueError):
        st.session_state["independent_ai_debug"] = "AI returned a non-numeric confidence value."
        return None

    if not isinstance(column, str) or not column.strip():
        st.session_state["independent_ai_debug"] = (
            "AI did not identify an entity column. "
            f"Reason: {str(result.get('reason', 'No reason supplied'))[:500]}"
        )
        return None

    # Exact matching prevents the model from creating or altering a field name.
    if column not in eligible_columns:
        st.session_state["independent_ai_debug"] = (
            f"AI selected {column!r}, but it is not a structurally eligible "
            "column in the uploaded worksheet."
        )
        return None

    if entity_type not in {"company", "organization", "business"}:
        st.session_state["independent_ai_debug"] = (
            f"AI returned unsupported entity_type: {entity_type!r}."
        )
        return None

    if not 0.70 <= confidence <= 1.0:
        st.session_state["independent_ai_debug"] = (
            f"AI confidence {confidence!r} is below the acceptance threshold. "
            "Choose the correct field manually if you can verify its meaning."
        )
        return None

    values = clean_entity_values(frame[column]).dropna().astype(str).str.strip()
    values = values[values.ne("")]
    counts = values.value_counts()
    repeated_groups = int((counts >= 2).sum())

    if len(values) < 4 or values.nunique() < 2 or repeated_groups < 2:
        st.session_state["independent_ai_debug"] = (
            "Python rejected the AI selection: it does not contain at least "
            "two distinct groups with repeated records."
        )
        return None

    st.session_state["independent_ai_debug"] = (
        f"AI suggested {column!r}; Python structural validation passed. "
        f"Confidence={confidence:.2f}; populated={len(values)}; "
        f"unique={values.nunique()}; repeated_groups={repeated_groups}. "
        "User confirmation is still required."
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
                    company_median = float(values.median())
                    peer_median = float(other_peers.median())
                    difference = company_median - peer_median

                    column_summary["peer_median"] = peer_median
                    column_summary[
                        "difference_from_peer_median"
                    ] = difference

                    # Explicit, same-statistic comparison for the AI.
                    # Never compare a company mean with a peer median.
                    column_summary["peer_comparison_reference"] = {
                        "statistic": "median",
                        "company_median": company_median,
                        "other_companies_median": peer_median,
                        "difference": difference,
                        "direction": (
                            "higher" if difference > 0
                            else "lower" if difference < 0
                            else "equal"
                        ),
                    }

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


def _generate_recommendation_with_retry(
    prompt: str,
    model: str,
    company_name: str,
) -> dict | None:
    """
    Retry Ollama recommendation generation when a response is
    incomplete or invalid. Never generate a Python recommendation here.
    """

    last_error = "No valid response received."

    for attempt in range(1, 4):
        result = _ollama_generate_json(prompt, model)

        if result is not None:
            recommendation = result.get("recommendation")
            evidence_columns = result.get("evidence_columns")
            priority = result.get("priority")

            if (
                isinstance(recommendation, str)
                and recommendation.strip()
                and len(recommendation.strip()) <= 2000
                and isinstance(evidence_columns, list)
                and len(evidence_columns) > 0
                and all(isinstance(column, str) for column in evidence_columns)
                and priority in {"high", "medium", "low"}
            ):
                st.session_state["independent_ai_debug"] = (
                    f"Ollama generated a valid recommendation for "
                    f"{company_name!r} on attempt {attempt}."
                )
                return result

            last_error = (
                "Ollama returned JSON, but its recommendation structure "
                "did not satisfy the required validation."
            )

        else:
            last_error = st.session_state.get(
                "independent_ai_debug",
                "Ollama returned no valid response.",
            )

        st.session_state["independent_ai_debug"] = (
            f"Recommendation attempt {attempt}/3 failed for "
            f"{company_name!r}: {last_error}"
        )

    st.session_state["independent_ai_debug"] = (
        f"Ollama could not generate a valid recommendation for "
        f"{company_name!r} after 3 attempts. Last error: {last_error}"
    )

    return None


def _generate_ai_company_recommendation(
    entity_value: str,
    group: pd.DataFrame,
    full_frame: pd.DataFrame,
    entity_column: str,
    previous_recommendations: list[dict] | None = None,
) -> str | None:
    """
    Generate one company-specific, evidence-grounded Ollama recommendation.

    Ollama remains the primary recommendation generator.
    Invalid responses receive one corrective retry.
    The existing Python fallback is used if both attempts fail.
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

    previous_context = (previous_recommendations or [])[-2:]

    # Collect numerical values actually present in the analytical evidence.
    allowed_numbers = set()

    def collect_numbers(value: Any) -> None:
        if isinstance(value, bool):
            return

        if isinstance(value, (int, float)):
            if math.isfinite(float(value)):
                allowed_numbers.add(round(float(value), 10))
            return

        if isinstance(value, dict):
            for nested_value in value.values():
                collect_numbers(nested_value)

        elif isinstance(value, (list, tuple)):
            for nested_value in value:
                collect_numbers(nested_value)

    collect_numbers(evidence)

    def correct_peer_comparison_direction(
        recommendation: str,
        evidence_columns: list[str],
    ) -> tuple[str, bool]:
        """Correct a generated direction using the computed peer evidence.

        This is a post-generation wording correction, not a rejection.
        It preserves the AI recommendation and prevents a directional
        contradiction from sending an otherwise valid response to fallback.
        """
        column_evidence = evidence.get("columns", {})
        matched_directions = []

        for column in evidence_columns:
            details = column_evidence.get(column, {})
            reference = details.get("peer_comparison_reference")

            if not isinstance(reference, dict):
                continue

            # Only correct a comparison when the recommendation identifies
            # the relevant source column, avoiding changes to unrelated claims.
            if str(column).casefold() not in recommendation.casefold():
                continue

            direction = reference.get("direction")
            if direction in {"higher", "lower", "equal"}:
                matched_directions.append(direction)

        # A recommendation should have one primary comparison. If multiple
        # distinct comparison directions are cited, do not guess which wording
        # belongs to which metric.
        unique_directions = set(matched_directions)
        if len(unique_directions) != 1:
            return recommendation, False

        direction = unique_directions.pop()
        if direction == "equal":
            return recommendation, False

        opposite_words = (
            {"higher": "lower", "above": "below", "greater": "less"}
            if direction == "lower"
            else {"lower": "higher", "below": "above", "less": "greater"}
        )

        corrected = recommendation
        for incorrect, correct in opposite_words.items():
            corrected = re.sub(
                rf"\b{incorrect}\b",
                correct,
                corrected,
                flags=re.IGNORECASE,
            )

        return corrected, corrected != recommendation

    def validate_response(result: dict) -> tuple[bool, str, str | None]:
        recommendation = result.get("recommendation")
        evidence_columns = result.get("evidence_columns")
        priority = result.get("priority")

        if not isinstance(recommendation, str):
            return False, "Recommendation must be a string.", None

        recommendation = recommendation.strip()

        if not recommendation:
            return False, "Recommendation is empty.", None

        if len(recommendation) > 2000:
            return False, "Recommendation exceeds 2000 characters.", None

        if not isinstance(evidence_columns, list) or not evidence_columns:
            return False, "Evidence columns are missing or empty.", None

        valid_columns = {
            str(column)
            for column in group.columns
            if str(column) != entity_column
        }

        invalid_columns = [
            column
            for column in evidence_columns
            if not isinstance(column, str)
            or column not in valid_columns
        ]

        if invalid_columns:
            return (
                False,
                f"Invalid evidence columns: {invalid_columns!r}",
                None,
            )

        if priority not in {"high", "medium", "low"}:
            return False, f"Invalid priority: {priority!r}", None

        # Reject numerical claims that cannot be found in the
        # current company's analytical evidence.
        number_pattern = re.compile(
            r"(?<![A-Za-z])[-+]?\d[\d,]*(?:\.\d+)?%?"
        )

        unsupported_numbers = []

        for match in number_pattern.findall(recommendation):
            normalized = match.replace(",", "").rstrip("%")

            try:
                number = round(float(normalized), 10)
            except ValueError:
                continue

            is_supported = any(
                math.isclose(
                    number,
                    allowed_number,
                    rel_tol=1e-6,
                    abs_tol=1e-6,
                )
                for allowed_number in allowed_numbers
            )

            if not is_supported:
                unsupported_numbers.append(match)

        if unsupported_numbers:
            return (
                False,
                "Unsupported numerical claims: "
                + ", ".join(unsupported_numbers),
                None,
            )

        return True, "Validation successful.", recommendation

    prompt = f"""
/no_think

You are an enterprise business intelligence recommendation engine.

Your responsibility is to analyse one company's actual records and
produce one high-quality, practical and evidence-grounded recommendation.

CURRENT COMPANY:
{entity_value}

COMPANY ANALYTICAL EVIDENCE:
{json.dumps(evidence, ensure_ascii=False, default=str)}

PREVIOUS COMPANY RECOMMENDATIONS:
{json.dumps(previous_context, ensure_ascii=False, default=str)}

IMPORTANT:
Previous recommendations are provided only to help avoid repetitive
wording. They are NOT evidence about the current company.

ANALYSIS REQUIREMENTS:

1. Analyse the current company's own analytical evidence.
2. Select the single most important actionable finding.
3. Consider missing values, numerical distributions, outliers,
   categorical patterns, record counts and peer comparisons where
   those measurements are available.
4. Use actual measured values from the supplied evidence.
5. Explain the practical significance of the selected finding.
6. Recommend specific actions that can be performed or investigated.
7. Explain which exact source columns support the recommendation.
8. Make the explanation specific to this company.
9. Similar underlying problems across companies may justify similar
   actions, but explain each company's own evidence separately.
10. When comparing a company with its peers, compare the SAME statistic.
    Use the supplied peer_comparison_reference: company_median against
    other_companies_median. Do not compare a company mean/average with a
    peer median.
11. Determine the direction from the two supplied values: if
    company_median is greater, say higher/above; if smaller, say
    lower/below; if equal, do not claim an advantage or disadvantage.
12. Do not contradict the supplied direction, even if another statistic
    such as the mean points in a different direction.

STRICT EVIDENCE RULES:

1. The supplied analytical evidence is the only factual source.
2. Never invent statistics, percentages, financial losses, targets,
   industry benchmarks, dates or external research.
3. Never claim that two metrics are correlated, interconnected,
   or causally related unless the supplied evidence establishes it.
4. Do not present a possible explanation as an established fact.
5. If a cause is unknown, recommend investigating it rather than
   claiming to know it.
6. Do not assume higher or lower values are automatically better.
7. Do not confuse missing data with poor business performance.
8. Do not recommend arbitrary mean, median or mode imputation.
9. Do not invent products, departments, customers or business processes.
10. Do not recommend actions unrelated to the measured evidence.
11. Do not invent monitoring frequencies or improvement targets.
12. If the evidence is insufficient to establish a business cause,
    explicitly preserve that uncertainty.

RECOMMENDATION QUALITY:

- Generate exactly ONE prioritized recommendation.
- Write 2 to 4 substantive sentences.
- Make the recommendation useful to a business decision-maker, not a generic summary.
- Identify the most important observed issue or opportunity for THIS company.
- Name the relevant source column(s) and describe the actual observed pattern.
- Give a concrete next action that a responsible business team can investigate or perform.
- Explain why that action follows from the supplied evidence.
- Separate observed facts from possible explanations; do not state an unverified cause.
- Prefer specific evidence over generic business language.
- For peer comparisons, use only the same-statistic median comparison
  explicitly supplied in peer_comparison_reference.
- State the company median and peer median only when both values are
  present, and describe their direction accurately.
- Avoid comparing a mean/average with a median.
- Avoid simply listing every analytical finding.
- Do not copy previous recommendations.
- Do not add unsupported claims merely to make the response detailed.

OUTPUT:
Return valid JSON only, using exactly this structure:

{{
    "recommendation": "One detailed, company-specific recommendation.",
    "evidence_columns": ["Exact source column names"],
    "priority": "high, medium, or low"
}}
"""

    generation_options = {
        "temperature": 0.1,
        "top_p": 0.9,
        "top_k": 40,
        "repeat_penalty": 1.05,
        "num_predict": 768,
        "num_ctx": 4096,
    }

    last_error = "No valid response received."

    # Initial generation + one corrective retry.
    for attempt in range(2):

        current_prompt = prompt

        if attempt == 1:
            current_prompt += f"""

    CORRECTIVE RETRY:

    Your previous response failed validation.

    Problem:
    {last_error}

    Regenerate the recommendation from the original company evidence.

    Correct the identified problem.
    If the validation error identifies an unsupported numerical claim,
    remove that number and state the finding qualitatively using the
    supplied evidence. Do not estimate or calculate a replacement value.
    For any peer comparison, use the explicit peer_comparison_reference
    and compare company_median only with other_companies_median. Follow
    its direction exactly; never compare a mean with a peer median.
    Keep the recommendation specific, actionable, and company-specific.
    Do not introduce unsupported facts.
    Return the exact required JSON structure.
    """

        result = _ollama_generate_json(
            current_prompt,
            OLLAMA_RECOMMENDATION_MODEL,
            generation_options=generation_options,
            timeout_seconds=150,
        )

        # Handle Ollama generation or connection failures.
        if not result:
            last_error = st.session_state.get(
                "independent_ai_debug",
                "Ollama returned no valid JSON response.",
            )

            # A timeout should not trigger another expensive generation attempt.
            if "timed out" in str(last_error).casefold():
                break

            # Allow one retry for other generation failures.
            continue

        # Validate a successfully received response.
        valid, reason, recommendation = validate_response(result)

        if not valid:
            last_error = reason
            continue

        # Enforce the mathematically computed peer direction after generation.
        # Do not reject or regenerate an otherwise valid AI recommendation.
        recommendation, direction_corrected = correct_peer_comparison_direction(
            recommendation,
            result.get("evidence_columns", []),
        )

        # Accept the validated recommendation.
        diagnostics[entity_value] = (
            "Ollama AI recommendation accepted. "
            f"Priority: {result.get('priority')}. "
            f"Evidence columns: {result.get('evidence_columns')}. "
            f"Company records analysed: {len(group)}. "
            f"Generation attempt: {attempt + 1}. "
            "Evidence-column and numerical-claim validation passed. "
            + (
                "Peer-comparison direction corrected from computed evidence."
                if direction_corrected
                else "Peer-comparison direction required no correction."
            )
        )

        st.session_state[
            "independent_ai_recommendation_debug"
        ] = diagnostics

        return recommendation

    fail(
        "Ollama recommendation generation failed after "
        f"2 attempts. Last validation/generation error: {last_error}"
    )

    return None


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




def _entity_header_semantic_score(column_name: str) -> tuple[float, str]:
    """Return a small, generic schema-label hint for entity-like headers.

    This uses broad semantic concepts, not any particular workbook's schema.
    It is only a ranking hint; the user must still confirm the field.
    """
    tokens = set(re.findall(r"[a-z0-9]+", str(column_name).casefold()))
    entity_terms = {
        "company", "organization", "organisation", "enterprise",
        "business", "firm", "corporation", "corporate", "legalentity",
    }
    identity_terms = {"name", "id", "identifier", "code", "number", "key"}

    matched = sorted(tokens & entity_terms)
    if matched:
        if tokens & identity_terms:
            return 1.0, f"Header contains entity concept(s): {', '.join(matched)}"
        return 0.85, f"Header contains entity concept(s): {', '.join(matched)}"
    return 0.0, "No explicit entity concept found in the header"


def _rank_python_entity_candidates(
    frame: pd.DataFrame,
) -> list[dict[str, Any]]:
    """Rank possible entity fields from structure plus generic header hints.

    Structural repetition is the main evidence. Generic semantic header
    hints break otherwise ambiguous ties (e.g. Company_Name vs Region).
    This is a suggestion, never proof of business meaning.
    """
    if frame.empty or len(frame.columns) == 0:
        return []

    profiles, eligible_columns = _build_entity_detection_profile(frame)
    profile_by_name = {item["column"]: item for item in profiles}
    ranked: list[dict[str, Any]] = []

    for column in sorted(eligible_columns, key=lambda value: (value.casefold(), value)):
        profile = profile_by_name.get(column)
        if not profile:
            continue

        populated = int(profile["non_null_count"])
        unique = int(profile["unique_count"])
        repeated_groups = int(profile["repeated_groups"])
        repeated_records = int(profile["records_in_repeated_groups"])
        record_count = int(profile["record_count"])

        if populated < 4 or unique < 2 or repeated_groups < 2:
            continue

        coverage = populated / max(record_count, 1)
        repeated_coverage = repeated_records / max(populated, 1)
        group_share = repeated_groups / max(unique, 1)
        structural_score = (
            0.40 * coverage
            + 0.40 * repeated_coverage
            + 0.20 * group_share
        )
        semantic_score, semantic_reason = _entity_header_semantic_score(column)
        # Structure remains the majority signal; the generic header hint only
        # resolves ties and helps avoid preferring Region/Product over Company.
        score = (0.65 * structural_score) + (0.35 * semantic_score)

        ranked.append({
            "column": column,
            "score": round(float(score), 4),
            "structural_score": round(float(structural_score), 4),
            "semantic_hint_score": round(float(semantic_score), 4),
            "unique_values": unique,
            "repeated_groups": repeated_groups,
            "populated_records": populated,
            "reason": (
                f"{unique} distinct values; {repeated_groups} groups contain "
                f"repeated records; {repeated_records}/{populated} populated "
                f"records belong to repeated groups. {semantic_reason}."
            ),
        })

    ranked.sort(key=lambda item: (-item["score"], item["column"].casefold(), item["column"]))
    st.session_state["independent_python_entity_candidates"] = ranked
    return ranked


def suggested_entity_column(frame: pd.DataFrame) -> str | None:
    """Try Ollama first; if it fails, rank candidates with Python.

    The returned field is only a suggestion. The UI still requires the user
    to explicitly select/confirm the grouping field.
    """
    if frame.empty or len(frame.columns) == 0:
        st.session_state["independent_python_entity_candidates"] = []
        return None

    # Primary path: always attempt AI first.
    ai_column = _ollama_detect_entity_column(frame)
    if ai_column is not None:
        st.session_state["independent_entity_suggestion_source"] = "Ollama AI"
        st.session_state["independent_ai_entity_detection"] = (
            f"Ollama suggested {ai_column!r}; Python structural validation passed. "
            "User confirmation is still required."
        )
        st.session_state["independent_python_entity_candidates"] = []
        return ai_column

    # Fallback path: run independently after the AI attempt has failed.
    ai_failure = st.session_state.get(
        "independent_ai_debug", "Ollama did not return a validated entity field."
    )
    try:
        ranked = _rank_python_entity_candidates(frame)
    except Exception as exc:
        ranked = []
        st.session_state["independent_python_entity_error"] = (
            f"{type(exc).__name__}: {exc}"
        )

    st.session_state["independent_entity_suggestion_source"] = "Python structural fallback"
    if ranked:
        best = ranked[0]
        st.session_state["independent_ai_entity_detection"] = (
            f"Ollama attempt failed ({ai_failure}). Python fallback ranked "
            f"{len(ranked)} structural candidate(s); top suggestion is "
            f"{best['column']!r}. This is not proof of company identity; "
            "confirm the field's meaning before proceeding."
        )
        return str(best["column"])

    st.session_state["independent_ai_entity_detection"] = (
        f"Ollama attempt failed ({ai_failure}). Python fallback found no "
        "structurally suitable repeated grouping field. You may still inspect "
        "and manually select a valid field."
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
    Generate one recommendation per identified company.

    Primary:
        Independent company-level analysis using Ollama.

    Fallback:
        Existing Python recommendation engine, only when AI generation
        is unavailable or its response fails validation.
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

    st.session_state[
        "independent_ai_recommendation_debug"
    ] = {}

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

        # Maintain previous recommendations so Ollama can distinguish
        # each company's explanation from the recommendations already made.
        previous_recommendations = []

        for label in unique_labels:
            mask = (
                entity_values.astype("string")
                .str.strip()
                .eq(label)
            )

            group = frame.loc[mask.fillna(False)]

            if group.empty:
                continue

            # PRIMARY: Ollama analyses this company's own complete records.
            recommendation = _generate_ai_company_recommendation(
                label,
                group,
                frame,
                entity_column,
                previous_recommendations,
            )

            if recommendation:
                st.session_state[
                    "independent_ai_recommendations_used"
                ] += 1

                recommendation_source = "Ollama"

            else:
                # FALLBACK: retain the existing Python recommendation engine.
                # Similarity between recommendations does NOT trigger fallback.
                recommendation = recommend_for_entity(
                    label,
                    group,
                    frame,
                    entity_column,
                )

                st.session_state[
                    "independent_python_fallback_recommendations"
                ] += 1

                recommendation_source = "Python fallback"

            rows.append({
                "Company Name": label,
                "Recommendation": recommendation,
            })

            # Make this recommendation available as context for subsequent
            # Ollama calls, regardless of whether it came from AI or fallback.
            previous_recommendations.append({
                "company": label,
                "recommendation": recommendation,
                "source": recommendation_source,
            })

        return pd.DataFrame(
            rows,
            columns=output_columns,
        )

    # No entity identified: do not invent a company name.
    return pd.DataFrame(columns=output_columns)




# ---------------------------------- Charts -------------------------------------


def _candidate_geo_columns(
    frame: pd.DataFrame,
    numeric: list[str],
) -> tuple[str, str] | None:
    """
    Identify coordinate columns using generic semantic evidence.
    Numeric ranges alone are never sufficient.
    """

    latitude_candidates = []
    longitude_candidates = []

    for column in frame.columns:
        tokens = set(
            re.findall(r"[a-z0-9]+", str(column).casefold())
        )

        if tokens & {"latitude", "lat"}:
            latitude_candidates.append(str(column))

        if tokens & {"longitude", "long", "lon", "lng"}:
            longitude_candidates.append(str(column))

    for lat_col in latitude_candidates:
        for lon_col in longitude_candidates:

            if lat_col == lon_col:
                continue

            lat = pd.to_numeric(frame[lat_col], errors="coerce")
            lon = pd.to_numeric(frame[lon_col], errors="coerce")

            valid = (
                lat.between(-90, 90)
                & lon.between(-180, 180)
            )

            if int(valid.sum()) >= 3:
                return lat_col, lon_col

    return None



@st.cache_data(show_spinner=False)
def _geocode_location_values(values: tuple[str, ...]) -> dict:
    """Resolve distinct geographic values into coordinates."""

    geolocator = Nominatim(
        user_agent="independent_company_intelligence"
    )

    reverse_lookup = RateLimiter(
        geolocator.geocode,
        min_delay_seconds=1.1,
        swallow_exceptions=True,
    )

    results = {}

    for value in values:
        try:
            result = reverse_lookup(
                value,
                exactly_one=True,
                addressdetails=True,
            )

            if result:
                address = result.raw.get("address", {})

                # Require geographic address evidence.
                location_parts = [
                    address.get("city"),
                    address.get("town"),
                    address.get("village"),
                    address.get("municipality"),
                    address.get("county"),
                    address.get("state"),
                    address.get("country"),
                ]

                location_parts = [
                    str(part) for part in location_parts if part
                ]

                if location_parts:
                    results[value] = {
                        "latitude": float(result.latitude),
                        "longitude": float(result.longitude),
                        "name": ", ".join(location_parts[:3]),
                    }

        except Exception:
            continue

    return results


def _discover_geographic_data(
    frame: pd.DataFrame,
    categorical: list[str],
) -> pd.DataFrame | None:
    """
    Discover a geographic field by testing its values.
    Does not rely on fixed column names or a particular dataset schema.
    """

    candidates = []

    for column in categorical:
        values = (
            frame[column]
            .dropna()
            .astype(str)
            .str.strip()
        )

        values = values[values.ne("")].drop_duplicates()

        if len(values) < 3:
            continue

        # Test a small sample before selecting a candidate.
        sample = tuple(values.head(3).tolist())

        resolved = _geocode_location_values(sample)

        success_ratio = len(resolved) / len(sample)

        if success_ratio >= 0.67:
            candidates.append(
                (success_ratio, len(resolved), column)
            )

    if not candidates:
        return None

    candidates.sort(
        key=lambda item: (item[0], item[1]),
        reverse=True,
    )

    selected_column = candidates[0][2]

    location_values = (
        frame[selected_column]
        .dropna()
        .astype(str)
        .str.strip()
    )

    unique_values = tuple(
        location_values[location_values.ne("")].drop_duplicates()
    )

    resolved_locations = _geocode_location_values(unique_values)

    if not resolved_locations:
        return None

    geo_data = frame.loc[
        frame[selected_column].notna()
    ].copy()

    geo_data["_Resolved_Location"] = (
        geo_data[selected_column].astype(str).str.strip()
    )

    geo_data["_Latitude"] = geo_data[
        "_Resolved_Location"
    ].map(
        lambda value: resolved_locations.get(value, {}).get("latitude")
    )

    geo_data["_Longitude"] = geo_data[
        "_Resolved_Location"
    ].map(
        lambda value: resolved_locations.get(value, {}).get("longitude")
    )

    geo_data["_Location_Name"] = geo_data[
        "_Resolved_Location"
    ].map(
        lambda value: resolved_locations.get(value, {}).get("name")
    )

    geo_data = geo_data.dropna(
        subset=["_Latitude", "_Longitude", "_Location_Name"]
    )

    return geo_data if len(geo_data) >= 3 else None



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

    
    # 7. Geographic distribution
    if chart_count < MAX_CHARTS_PER_SHEET:

        geo_data = None

        # First preference: genuine coordinate columns.
        geo_pair = _candidate_geo_columns(frame, numeric)

        if geo_pair:
            lat_col, lon_col = geo_pair

            geo_data = frame.copy()

            geo_data["_Latitude"] = pd.to_numeric(
                geo_data[lat_col], errors="coerce"
            )

            geo_data["_Longitude"] = pd.to_numeric(
                geo_data[lon_col], errors="coerce"
            )

            geo_data = geo_data[
                geo_data["_Latitude"].between(-90, 90)
                & geo_data["_Longitude"].between(-180, 180)
            ].copy()

            geo_data["_Location_Name"] = (
                geo_data["_Latitude"].round(4).astype(str)
                + ", "
                + geo_data["_Longitude"].round(4).astype(str)
            )

        # Second preference: discover and geocode descriptive locations.
        if geo_data is None or len(geo_data) < 3:
            geo_data = _discover_geographic_data(
                frame,
                categorical,
            )

        if geo_data is not None and len(geo_data) >= 3:

            
                # Preserve original dataset details in map tooltips.
                detail_columns = [
                    col for col in frame.columns
                    if col in geo_data.columns
                ]

                # Display one map marker per distinct location.
                map_data = geo_data.drop_duplicates(
                    subset=["_Latitude", "_Longitude"]
                ).copy()

                fig = px.scatter_geo(
                    geo_data,
                    lat="_Latitude",
                    lon="_Longitude",
                    hover_name="_Location_Name",
                    hover_data={
                        col: True for col in detail_columns
                    },
                    title="Geographic Distribution",
                    template="plotly_dark",
                )

                fig.update_traces(
                    mode="markers+text",
                    textposition="top center",
                    textfont=dict(
                        size=9,
                        color="#FFFFFF",
                    ),
                    marker=dict(
                        size=7,
                        color="#818CF8",
                        opacity=0.9,
                    ),
                )

                fig.update_geos(
                    projection_type="natural earth",
                    showland=True,
                    landcolor="#263449",
                    showocean=True,
                    oceancolor="#111827",
                    showcountries=True,
                    countrycolor="#64748B",
                    showcoastlines=True,
                    coastlinecolor="#64748B",
                    bgcolor="rgba(0,0,0,0)",
                )

                fig.update_layout(
                    title=dict(
                        text="Geographic Distribution",
                        font=dict(color="#FFFFFF", size=18),
                        x=0.02,
                    ),
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(color="#FFFFFF"),
                    height=550,
                    margin=dict(l=10, r=10, t=65, b=10),
                    showlegend=False,
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                )

                st.caption("Geographic Location Details")

                display_columns = [
                    col for col in geo_data.columns
                    if not str(col).startswith("_")
                ]

                st.dataframe(
                    geo_data[display_columns],
                    use_container_width=True,
                    hide_index=True,
                )

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

    ai_suggested_entity_column = suggested_entity_column(frame)

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

        st.write("Full runtime column profiles:")
        st.json(st.session_state.get("independent_ai_column_profiles", []))

        st.write("Python fallback candidates:")
        st.json(st.session_state.get("independent_python_entity_candidates", []))

        st.write("Suggestion source:")
        st.write(st.session_state.get("independent_entity_suggestion_source", "Not available"))

    if ai_suggested_entity_column:
        st.info(
            f"Suggested entity field ({st.session_state.get('independent_entity_suggestion_source', 'unknown source')}): "
            f"{ai_suggested_entity_column}. Confirm that this field genuinely "
            "identifies the organization whose records should be grouped."
        )
    else:
        st.warning(
            "Ollama was attempted but did not return a validated entity field, "
            "and Python did not find a suitable candidate. You may still inspect "
            "the column profiles and manually select a field if its meaning is known."
        )

    # Bind the confirmation widget to the uploaded workbook contents and sheet.
    # This prevents a previous workbook's selection from carrying into a new file.
    workbook_fingerprint = hashlib.sha256(uploaded.getvalue()).hexdigest()[:12]
    sheet_fingerprint = hashlib.sha256(str(selected_sheet).encode("utf-8")).hexdigest()[:8]
    confirmation_key = (
        f"independent_entity_column_confirmation_{workbook_fingerprint}_{sheet_fingerprint}"
    )
    column_options = ["-- Select entity field --"] + [str(c) for c in frame.columns]
    selected_entity_column = st.selectbox(
        "Confirm the company / organization grouping field",
        column_options,
        index=0,
        key=confirmation_key,
        help="Choose the field whose repeated values identify the same company or organization. Do not choose a customer, product, date, or transaction field unless that is genuinely the entity you want to analyse.",
    )

    entity_column = None
    if selected_entity_column != "-- Select entity field --":
        selected_values = clean_entity_values(
            frame[selected_entity_column]
        ).dropna().astype(str).str.strip()
        selected_values = selected_values[selected_values.ne("")]
        selected_counts = selected_values.value_counts()
        selected_repeated_groups = int((selected_counts >= 2).sum())

        if (
            len(selected_values) >= 4
            and selected_values.nunique() >= 2
            and selected_repeated_groups >= 2
        ):
            entity_column = selected_entity_column
            st.success(
                f"Confirmed grouping field: {entity_column} "
                f"({selected_values.nunique()} distinct values; "
                f"{selected_repeated_groups} repeated groups)."
            )
        else:
            st.error(
                "This field does not have enough repeated structure for "
                "reliable entity-level grouping. Select another field."
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
                f"{fallback_count} recommendation(s) were generated by Python "
                "after the AI attempt failed or returned invalid output."
            )

        fallback_errors = st.session_state.get("independent_python_fallback_errors", {})
        if fallback_errors:
            st.error(
                f"Python fallback also failed for {len(fallback_errors)} entity/entities. "
                "See diagnostics for the exact exceptions."
            )

        with st.expander("AI Recommendation Diagnostics"):

            st.caption(f"Running source file: {Path(__file__).resolve()}")

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

        # Automatically size the recommendation panel to its content.
        # No fixed height, so the last recommendation is never clipped.
        with st.container(border=True):

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
