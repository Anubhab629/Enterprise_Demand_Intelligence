"""
Enterprise Demand Intelligence Platform
GENERIC DATASET-DRIVEN RECOMMENDATION ENGINE

Consumes only upstream analytical and forecasting evidence.
No business schema, aliases, table names, column names, entity counts,
or hard-coded business actions are required.

Recommendation text is generated from runtime evidence through a
provider-neutral, configurable AI endpoint. No dataset-specific business
rules or manually authored business actions are embedded in this module.

The module:
- runs independently from VS Code
- is callable by main.py as a worker
- never imports main.py
- discovers upstream workbook contracts dynamically
- preserves the existing evidence/scoring architecture
- writes recommendation_results.xlsx and recommendation_report.json
"""
from __future__ import annotations

import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import math
import os
import re
import socket
import time
import sys
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import config

OUTPUT_DIR = Path(
    getattr(config, "OUTPUT_DATA_DIRECTORY", PROJECT_ROOT / "data" / "output")
).expanduser().resolve()

TOP_RECOMMENDATIONS = max(1, int(getattr(config, "TOP_RECOMMENDATIONS", 100)))
MIN_SIGNAL_STRENGTH = max(
    0.0,
    min(1.0, float(getattr(config, "RECOMMENDATION_MIN_SIGNAL_STRENGTH", 0.35))),
)
MIN_DIRECTIONAL_AGREEMENT = max(
    0.0,
    min(
        1.0,
        float(getattr(config, "RECOMMENDATION_MIN_DIRECTIONAL_AGREEMENT", 0.55)),
    ),
)
MIN_FORECAST_MOVE_PCT = max(
    0.0,
    float(getattr(config, "RECOMMENDATION_MIN_FORECAST_MOVEMENT_PCT", 5.0)),
)
MAX_SIGNAL_ROWS = max(
    1, int(getattr(config, "RECOMMENDATION_MAX_SIGNAL_ROWS", 50000))
)
MAX_CANDIDATE_ROWS = max(
    1, int(getattr(config, "RECOMMENDATION_MAX_CANDIDATE_ROWS", 50000))
)
MAX_FORECAST_ROWS = max(
    1, int(getattr(config, "RECOMMENDATION_MAX_FORECAST_ROWS", 50000))
)
MAX_SIGNALS_PER_TARGET = max(
    1, int(getattr(config, "RECOMMENDATION_MAX_SIGNALS_PER_TARGET", 5))
)
MAX_INPUT_FILES_TO_CHECK = max(
    1, int(getattr(config, "RECOMMENDATION_MAX_INPUT_FILES_TO_CHECK", 10))
)

# Provider-neutral AI configuration.
# These are environment/configuration controls, not dataset assumptions.
AI_ENABLED = str(os.getenv("EDI_AI_ENABLED", "false")).strip().lower() in {
    "1",
    "true",
    "yes",
    "on",
}
AI_API_URL = str(os.getenv("EDI_AI_API_URL", "")).strip()
AI_API_KEY = str(os.getenv("EDI_AI_API_KEY", "")).strip()
AI_MODEL = str(os.getenv("EDI_AI_MODEL", "")).strip()
AI_GENERATION_CONTRACT_VERSION = "recommendation_ai_contract_v7_dataset_independent_evidence"
AI_ENRICHMENT_ENABLED = str(os.getenv("EDI_AI_ENRICHMENT_ENABLED", "false")).strip().lower() in {
    "1", "true", "yes", "on",
}
try:
    AI_TIMEOUT_SECONDS = max(20, int(os.getenv("EDI_AI_TIMEOUT_SECONDS", "45")))
except ValueError:
    AI_TIMEOUT_SECONDS = 45

try:
    AI_TEMPERATURE = max(
        0.0, min(1.0, float(os.getenv("EDI_AI_TEMPERATURE", "0.0")))
    )
except ValueError:
    AI_TEMPERATURE = 0.2

try:
    AI_MAX_OUTPUT_TOKENS = max(128, int(os.getenv("EDI_AI_MAX_OUTPUT_TOKENS", "128")))
except ValueError:
    AI_MAX_OUTPUT_TOKENS = 256

try:
    AI_BATCH_SIZE = max(1, min(3, int(os.getenv("EDI_AI_BATCH_SIZE", "3"))))
except ValueError:
    AI_BATCH_SIZE = 6

try:
    AI_PARALLELISM = max(1, min(3, int(os.getenv("EDI_AI_PARALLELISM", "3"))))
except ValueError:
    AI_PARALLELISM = 3

try:
    AI_MAX_RETRIES = max(3, int(os.getenv("EDI_AI_MAX_RETRIES", "5")))
except ValueError:
    AI_MAX_RETRIES = 8

ANALYSIS_OUTPUT_FILE = OUTPUT_DIR / "analysis_results.xlsx"
FORECAST_OUTPUT_FILE = OUTPUT_DIR / "forecast_results.xlsx"
RECOMMENDATION_OUTPUT_FILE = OUTPUT_DIR / "recommendation_results.xlsx"
RECOMMENDATION_REPORT_FILE = OUTPUT_DIR / "recommendation_report.json"

# Runtime diagnostics populated by the current worker execution.
LAST_RUNTIME_STATS: dict[str, Any] = {}
LAST_AI_STATS: dict[str, Any] = {}

SIGNAL_SCHEMA = {
    "target_table", "target_time_column", "target_measure_column",
    "signal_table", "signal_time_column", "signal_measure_column",
    "lag_periods", "aligned_observations", "level_correlation",
    "change_correlation", "directional_agreement", "direction",
    "relationship_strength",
}
CANDIDATE_SCHEMA = {
    "table", "time_column", "measure_column", "candidate_score",
}
FORECAST_SCHEMA = {
    "table", "time_column", "value_column", "ForecastAverage",
    "RecentVsBaselineChangePct", "ForecastConfidence",
}
FORECAST_META = {
    "table", "time_column", "value_column", "ForecastStart", "ForecastEnd",
    "HistoryPoints", "RecentAverage", "BaselineAverage", "ForecastAverage",
    "ForecastTotal", "LowerForecastTotal", "UpperForecastTotal", "Volatility",
    "CoefficientOfVariation", "RecentVsBaselineChangePct", "ForecastConfidence",
    "FrequencyEvidence", "ForecastPoints", "ForecastHorizonDays", "candidate_score",
}


def _safe(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (str, int, bool)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if isinstance(value, pd.Timedelta):
        return value.total_seconds()
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    try:
        return value.item()
    except (AttributeError, ValueError):
        return str(value)


def _text(value: Any) -> str:
    safe = _safe(value)
    return "" if safe is None else str(safe).strip()


def _json(value: Any) -> str:
    return json.dumps(
        _safe(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def _id(value: Any) -> str:
    return hashlib.sha256(_json(value).encode("utf-8")).hexdigest()[:16]


def _set_score(value: Any) -> float:
    try:
        score = float(value)
    except (TypeError, ValueError):
        return 0.0
    if not math.isfinite(score):
        return 0.0
    if score <= 1.0:
        score *= 100.0
    return max(0.0, min(100.0, score))


def _existing_file(path: Path) -> Path | None:
    return path if path.is_file() else None


def _configured_path(name: str) -> Path | None:
    value = getattr(config, name, None)
    if not value:
        return None
    try:
        path = Path(value).expanduser().resolve()
    except (TypeError, ValueError):
        return None
    return _existing_file(path)


def resolve_analysis_workbook() -> Path:
    canonical = _existing_file(ANALYSIS_OUTPUT_FILE)
    if canonical is not None:
        return canonical

    for name in ("ANALYSIS_RESULTS_WORKBOOK", "ANALYTICAL_RESULTS_WORKBOOK"):
        configured = _configured_path(name)
        if configured is not None:
            return configured

    candidates = sorted(
        [p for p in OUTPUT_DIR.glob("*.xlsx") if p.is_file()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for path in candidates[:MAX_INPUT_FILES_TO_CHECK]:
        try:
            headers = read_workbook_headers(path)
        except Exception:
            continue
        if any(
            SIGNAL_SCHEMA.issubset(set(map(str, cols)))
            or CANDIDATE_SCHEMA.issubset(set(map(str, cols)))
            for cols in headers.values()
        ):
            return path

    raise FileNotFoundError(
        "No compatible analytical-results workbook was found. "
        "Run the analytical stage first."
    )


def resolve_forecast_workbook() -> Path:
    canonical = _existing_file(FORECAST_OUTPUT_FILE)
    if canonical is not None:
        return canonical

    for name in ("FORECAST_RESULTS_WORKBOOK", "FORECAST_WORKBOOK"):
        configured = _configured_path(name)
        if configured is not None:
            return configured

    candidates = sorted(
        [p for p in OUTPUT_DIR.glob("*.xlsx") if p.is_file()],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for path in candidates[:MAX_INPUT_FILES_TO_CHECK]:
        try:
            headers = read_workbook_headers(path)
        except Exception:
            continue
        if any(
            FORECAST_SCHEMA.issubset(set(map(str, cols)))
            for cols in headers.values()
        ):
            return path

    raise FileNotFoundError(
        "No compatible forecast-results workbook was found. "
        "Run the forecasting stage first."
    )


def resolve_output_workbook() -> Path:
    for name in ("RECOMMENDATION_WORKBOOK", "RECOMMENDATIONS_WORKBOOK"):
        configured = _configured_path(name)
        if configured is not None:
            return configured
    return RECOMMENDATION_OUTPUT_FILE


def read_workbook_headers(path: Path) -> dict[str, list[Any]]:
    workbook = pd.ExcelFile(path, engine="openpyxl")
    headers: dict[str, list[Any]] = {}
    for sheet_name in workbook.sheet_names:
        frame = pd.read_excel(workbook, sheet_name=sheet_name, nrows=0)
        headers[str(sheet_name)] = list(frame.columns)
    return headers


def _schema_score(columns: list[Any], required: set[str]) -> tuple[int, bool]:
    available = set(map(str, columns))
    overlap = len(required.intersection(available))
    return overlap, required.issubset(available)


def _best_sheet(headers: dict[str, list[Any]], required: set[str]) -> str | None:
    ranked: list[tuple[int, int, str]] = []
    for name, cols in headers.items():
        overlap, complete = _schema_score(cols, required)
        if overlap:
            ranked.append((1 if complete else 0, overlap, name))
    ranked.sort(key=lambda item: (-item[0], -item[1], item[2]))
    return ranked[0][2] if ranked else None


def _load_sheet(path: Path, name: str, limit: int) -> pd.DataFrame:
    workbook = pd.ExcelFile(path, engine="openpyxl")
    frame = pd.read_excel(workbook, sheet_name=name)
    if len(frame) > limit:
        raise RuntimeError(
            f"Evidence sheet '{name}' contains {len(frame):,} rows, exceeding "
            f"the configured technical safety limit of {limit:,}. "
            "Refusing to silently truncate analytical evidence."
        )
    return frame


def _empty_signal_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=sorted(SIGNAL_SCHEMA))


def _empty_candidate_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=sorted(CANDIDATE_SCHEMA))


def _empty_forecast_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=sorted(FORECAST_SCHEMA))


def load_analysis_evidence(
    path: Path,
) -> tuple[dict[str, list[Any]], pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    headers = read_workbook_headers(path)
    signal_sheet = _best_sheet(headers, SIGNAL_SCHEMA)
    candidate_sheet = _best_sheet(headers, CANDIDATE_SCHEMA)

    signals = (
        _load_sheet(path, signal_sheet, MAX_SIGNAL_ROWS)
        if signal_sheet is not None
        else _empty_signal_frame()
    )
    candidates = (
        _load_sheet(path, candidate_sheet, MAX_CANDIDATE_ROWS)
        if candidate_sheet is not None
        else _empty_candidate_frame()
    )

    metadata = {
        "signal_sheet": signal_sheet,
        "candidate_sheet": candidate_sheet,
        "analysis_sheet_count": len(headers),
        "analysis_contract_valid": signal_sheet is not None and candidate_sheet is not None,
    }
    return headers, signals, candidates, metadata


def load_forecast_evidence(
    path: Path,
) -> tuple[dict[str, list[Any]], pd.DataFrame, dict[str, Any]]:
    headers = read_workbook_headers(path)
    forecast_sheet = _best_sheet(headers, FORECAST_SCHEMA)
    forecasts = (
        _load_sheet(path, forecast_sheet, MAX_FORECAST_ROWS)
        if forecast_sheet is not None
        else _empty_forecast_frame()
    )
    metadata = {
        "forecast_sheet": forecast_sheet,
        "forecast_sheet_count": len(headers),
        "forecast_contract_valid": forecast_sheet is not None,
    }
    return headers, forecasts, metadata


def normalize_candidates(frame: pd.DataFrame) -> dict[tuple[str, str], float]:
    if frame.empty or not CANDIDATE_SCHEMA.issubset(set(map(str, frame.columns))):
        return {}

    work = frame.copy()
    work["candidate_score"] = pd.to_numeric(
        work["candidate_score"], errors="coerce"
    )
    work = work.dropna(subset=["candidate_score"])

    result: dict[tuple[str, str], float] = {}
    for _, row in work.sort_values(
        "candidate_score", ascending=False, kind="stable"
    ).iterrows():
        key = (_text(row["table"]), _text(row["measure_column"]))
        if not key[0] or not key[1]:
            continue
        result.setdefault(key, _set_score(row["candidate_score"]))
    return result


def normalize_signals(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or not SIGNAL_SCHEMA.issubset(set(map(str, frame.columns))):
        return _empty_signal_frame()

    work = frame.copy()
    numeric = [
        "lag_periods",
        "aligned_observations",
        "level_correlation",
        "change_correlation",
        "directional_agreement",
        "relationship_strength",
    ]
    for column in numeric:
        work[column] = pd.to_numeric(work[column], errors="coerce")

    for column in (
        "target_table",
        "target_measure_column",
        "signal_table",
        "signal_measure_column",
    ):
        work[column] = work[column].astype(str).str.strip()

    work = work[
        work["relationship_strength"].notna()
        & work["directional_agreement"].notna()
        & work["target_table"].ne("")
        & work["target_measure_column"].ne("")
        & work["signal_table"].ne("")
        & work["signal_measure_column"].ne("")
    ].copy()

    return work.reset_index(drop=True)


def normalize_forecasts(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty or not FORECAST_SCHEMA.issubset(set(map(str, frame.columns))):
        return pd.DataFrame()

    work = frame.copy()
    work["__change"] = pd.to_numeric(
        work["RecentVsBaselineChangePct"], errors="coerce"
    )
    work["__avg"] = pd.to_numeric(work["ForecastAverage"], errors="coerce")
    work["__table"] = work["table"].astype(str).str.strip()
    work["__measure"] = work["value_column"].astype(str).str.strip()

    work = work[
        work["__change"].notna()
        & work["__avg"].notna()
        & work["__table"].ne("")
        & work["__measure"].ne("")
    ].copy()

    rows: list[dict[str, Any]] = []
    for _, row in work.iterrows():
        scope = {
            str(column): _safe(value)
            for column, value in row.items()
            if str(column) not in FORECAST_META and not str(column).startswith("__") and _safe(value) is not None
        }

        rows.append(
            {
                "target_table": _text(row["__table"]),
                "target_time_column": _text(row["time_column"]),
                "target_measure_column": _text(row["__measure"]),
                "recent_change_pct": _safe(row["__change"]),
                "forecast_average": _safe(row["__avg"]),
                "forecast_total": _safe(row.get("ForecastTotal")),
                "recent_average": _safe(row.get("RecentAverage")),
                "baseline_average": _safe(row.get("BaselineAverage")),
                "forecast_confidence": _text(row.get("ForecastConfidence")),
                "scope": scope,
                "candidate_score": _safe(row.get("candidate_score")),
            }
        )

    return pd.DataFrame(rows)


def signal_support(row: pd.Series) -> float:
    try:
        strength = float(row.get("relationship_strength", 0.0))
    except (TypeError, ValueError):
        strength = 0.0

    try:
        agreement = float(row.get("directional_agreement", 0.0))
    except (TypeError, ValueError):
        agreement = 0.0

    try:
        observations = float(row.get("aligned_observations", 0.0) or 0.0)
    except (TypeError, ValueError):
        observations = 0.0

    strength = max(0.0, min(1.0, strength))
    agreement = max(0.0, min(1.0, agreement))
    observation_score = max(0.0, min(1.0, observations / 100.0))

    return max(
        0.0,
        min(
            100.0,
            (0.55 * strength + 0.30 * agreement + 0.15 * observation_score) * 100.0,
        ),
    )


def select_signals(
    frame: pd.DataFrame,
) -> dict[tuple[str, str], list[dict[str, Any]]]:
    if frame.empty:
        return {}

    work = normalize_signals(frame)
    if work.empty:
        return {}

    work["__support"] = work.apply(signal_support, axis=1)
    work = work[
        (work["relationship_strength"] >= MIN_SIGNAL_STRENGTH)
        & (work["directional_agreement"] >= MIN_DIRECTIONAL_AGREEMENT)
    ].copy()

    if work.empty:
        return {}

    work = work.sort_values(
        [
            "target_table",
            "target_measure_column",
            "__support",
            "aligned_observations",
        ],
        ascending=[True, True, False, False],
        kind="stable",
    )

    result: dict[tuple[str, str], list[dict[str, Any]]] = {}

    for _, row in work.iterrows():
        key = (
            _text(row["target_table"]),
            _text(row["target_measure_column"]),
        )
        bucket = result.setdefault(key, [])

        if len(bucket) >= MAX_SIGNALS_PER_TARGET:
            continue

        bucket.append(
            {
                "signal_table": _text(row["signal_table"]),
                "signal_time_column": _text(row["signal_time_column"]),
                "signal_measure_column": _text(row["signal_measure_column"]),
                "lag_periods": _safe(row["lag_periods"]),
                "aligned_observations": _safe(row["aligned_observations"]),
                "level_correlation": _safe(row["level_correlation"]),
                "change_correlation": _safe(row["change_correlation"]),
                "directional_agreement": _safe(row["directional_agreement"]),
                "direction": _text(row["direction"]),
                "relationship_strength": _safe(row["relationship_strength"]),
                "support_score": _safe(row["__support"]),
            }
        )

    return result


def _movement_score(change: Any) -> float:
    try:
        value = abs(float(change))
    except (TypeError, ValueError):
        return 0.0

    if not math.isfinite(value):
        return 0.0

    denominator = max(MIN_FORECAST_MOVE_PCT * 3.0, 1e-9)
    return max(0.0, min(1.0, value / denominator))


def _confidence_score(value: Any) -> float:
    text = _text(value).lower()
    if text == "high":
        return 1.0
    if text == "medium":
        return 0.65
    if text == "low":
        return 0.35
    return 0.50


def _priority_level(score: float) -> str:
    if score >= 75.0:
        return "High"
    if score >= 40.0:
        return "Medium"
    return "Low"


def _priority_score(
    forecast: dict[str, Any] | None,
    signals: list[dict[str, Any]],
    candidate_score: float,
) -> float:
    signal_component = 0.0

    if signals:
        signal_component = (
            max(
                float(item.get("support_score", 0.0) or 0.0)
                for item in signals
            )
            / 100.0
        )

    forecast_component = 0.0
    confidence_component = 0.50

    if forecast is not None:
        forecast_component = _movement_score(
            forecast.get("recent_change_pct")
        )
        confidence_component = _confidence_score(
            forecast.get("forecast_confidence")
        )

    candidate_component = max(
        0.0,
        min(1.0, _set_score(candidate_score) / 100.0),
    )

    score = (
        40.0 * signal_component
        + 30.0 * forecast_component
        + 20.0 * confidence_component
        + 10.0 * candidate_component
    )

    return max(0.0, min(100.0, score))


def _recommendation_payload(
    target_table: str,
    target_measure: str,
    scope: dict[str, Any],
    forecast: dict[str, Any] | None,
    signals: list[dict[str, Any]],
    candidate_score: float,
) -> dict[str, Any]:
    evidence_mode = "signal_and_forecast"

    if forecast is None:
        evidence_mode = "signal_only"
    elif not signals:
        evidence_mode = "forecast_only"

    return {
        "evidence_mode": evidence_mode,
        "target": {
            "table": target_table,
            "measure": target_measure,
            "scope": scope,
        },
        "candidate_score": _set_score(candidate_score),
        "forecast": forecast,
        "supporting_signals": signals,
    }


def _ai_is_configured() -> bool:
    if not (AI_ENABLED and AI_API_URL and AI_MODEL):
        return False
    # Native local Ollama /api/chat does not require an API key.
    if _is_ollama_native_endpoint():
        return True
    # Other provider-compatible endpoints require an explicit credential.
    return bool(AI_API_KEY)


def _recommendation_generation_context(
    target_table: str,
    target_measure: str,
    scope: dict[str, Any],
    forecast: dict[str, Any],
    signals: list[dict[str, Any]],
    candidate_score: float,
    priority_score: float,
    priority_level: str,
) -> dict[str, Any]:
    """
    Build a dataset-independent evidence projection for recommendation generation.

    Business identifiers, raw scope, ranking metadata, and schema names are
    deliberately excluded. Only generic analytical evidence is exposed.
    """
    del target_table
    del target_measure
    del scope
    del candidate_score
    del priority_score
    del priority_level

    forecast_evidence: dict[str, Any] = {
        "forecast_available": False,
        "recent_change_pct": None,
        "forecast_average": None,
        "forecast_total": None,
        "recent_average": None,
        "baseline_average": None,
        "movement_direction": "not_available",
        "movement_material": False,
        "movement_strength": None,
        "forecast_confidence": "",
        "time_column_available": False,
    }

    if isinstance(forecast, dict):
        forecast_evidence["forecast_available"] = True
        for field in (
            "recent_change_pct",
            "forecast_average",
            "forecast_total",
            "recent_average",
            "baseline_average",
        ):
            value = _safe(forecast.get(field))
            if value is not None:
                try:
                    forecast_evidence[field] = float(value)
                except (TypeError, ValueError):
                    forecast_evidence[field] = value

        change_value = forecast_evidence.get("recent_change_pct")
        try:
            numeric_change = float(change_value)
            if numeric_change > 0:
                forecast_evidence["movement_direction"] = "increasing"
            elif numeric_change < 0:
                forecast_evidence["movement_direction"] = "decreasing"
            else:
                forecast_evidence["movement_direction"] = "unchanged"
            forecast_evidence["movement_material"] = (
                abs(numeric_change) >= MIN_FORECAST_MOVE_PCT
            )
            forecast_evidence["movement_strength"] = abs(numeric_change)
        except (TypeError, ValueError):
            pass

        forecast_evidence["forecast_confidence"] = _text(
            forecast.get("forecast_confidence")
        )
        forecast_evidence["time_column_available"] = bool(
            _text(forecast.get("target_time_column"))
        )

    supporting_signal_count = 0
    aligned_signal_count = 0
    opposed_signal_count = 0
    unresolved_signal_count = 0
    relationship_strengths: list[float] = []
    directional_agreements: list[float] = []

    for signal in signals:
        if not isinstance(signal, dict):
            continue

        supporting_signal_count += 1
        direction = _text(signal.get("direction")).strip().lower()
        if direction in {"same", "aligned", "positive", "increasing", "decreasing"}:
            aligned_signal_count += 1
        elif direction in {"opposite", "opposed", "negative"}:
            opposed_signal_count += 1
        else:
            unresolved_signal_count += 1

        relationship = _safe(signal.get("relationship_strength"))
        if relationship is not None:
            try:
                relationship_strengths.append(float(relationship))
            except (TypeError, ValueError):
                pass

        agreement = _safe(signal.get("directional_agreement"))
        if agreement is not None:
            try:
                directional_agreements.append(float(agreement))
            except (TypeError, ValueError):
                pass

    return {
        "candidate_evidence_profile": {
            "evidence_mode": (
                "signal_and_forecast"
                if forecast_evidence["forecast_available"] and supporting_signal_count
                else (
                    "forecast_only"
                    if forecast_evidence["forecast_available"]
                    else "signal_only"
                    if supporting_signal_count
                    else "evidence_limited"
                )
            ),
        },
        "forecast_evidence": forecast_evidence,
        "signal_evidence": {
            "supporting_signal_count": supporting_signal_count,
            "aligned_signal_count": aligned_signal_count,
            "opposed_signal_count": opposed_signal_count,
            "unresolved_signal_count": unresolved_signal_count,
            "max_relationship_strength": (
                max(relationship_strengths) if relationship_strengths else None
            ),
            "max_directional_agreement": (
                max(directional_agreements) if directional_agreements else None
            ),
        },
    }




def _extract_ai_content(response_payload: Any) -> str:
    if not isinstance(response_payload, dict):
        return ""

    choices = response_payload.get("choices")
    if isinstance(choices, list) and choices:
        first = choices[0]
        if isinstance(first, dict):
            message = first.get("message")
            if isinstance(message, dict):
                content = message.get("content")
                if isinstance(content, str):
                    return content.strip()

            text = first.get("text")
            if isinstance(text, str):
                return text.strip()

    output = response_payload.get("output")
    if isinstance(output, str):
        return output.strip()

    return ""


def _clean_json_text(text: str) -> str:
    cleaned = text.strip()

    if cleaned.startswith("```"):
        lines = cleaned.splitlines()

        if lines and lines[0].strip().startswith("```"):
            lines = lines[1:]

        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]

        cleaned = "\n".join(lines).strip()

    return cleaned



def _empty_ai_result() -> dict[str, str]:
    return {
        "finding": "",
        "recommendation": "",
        "rationale": "",
        "expected_impact": "",
        "risk": "",
    }



def _extract_numeric_tokens(text: str) -> list[float]:
    """Extract ordinary numeric values from generated prose."""
    if not isinstance(text, str):
        return []
    tokens = re.findall(r"(?<![A-Za-z_])[-+]?\d+(?:,\d{3})*(?:\.\d+)?%?", text)
    values: list[float] = []
    for token in tokens:
        cleaned = token.replace(",", "").rstrip("%")
        try:
            value = float(cleaned)
        except ValueError:
            continue
        if math.isfinite(value):
            values.append(value)
    return values


def _collect_business_numbers(value: Any, key_name: str = "") -> tuple[list[float], list[float]]:
    """Collect numeric evidence and explicitly percentage-valued evidence."""
    numbers: list[float] = []
    percentages: list[float] = []
    if isinstance(value, dict):
        for key, item in value.items():
            key_text = str(key).lower()
            if key_text in {
                "table", "measure", "target_time_column", "time_column",
                "signal_table", "signal_time_column", "signal_measure_column",
                "confidence", "direction",
            }:
                continue
            child_numbers, child_percentages = _collect_business_numbers(item, key_text)
            numbers.extend(child_numbers)
            percentages.extend(child_percentages)
        return numbers, percentages
    if isinstance(value, (list, tuple)):
        for item in value:
            child_numbers, child_percentages = _collect_business_numbers(item, key_name)
            numbers.extend(child_numbers)
            percentages.extend(child_percentages)
        return numbers, percentages
    if isinstance(value, bool) or value is None:
        return numbers, percentages
    if isinstance(value, (int, float)):
        try:
            numeric = float(value)
        except (TypeError, ValueError):
            return numbers, percentages
        if math.isfinite(numeric):
            numbers.append(numeric)
            if key_name.endswith("_pct") or key_name.endswith("percentage"):
                percentages.append(numeric)
    return numbers, percentages


def _numeric_supported(generated_text: str, context: dict[str, Any]) -> bool:
    """Require every generated numeric value to exist with compatible units."""
    tokens = re.findall(
        r"(?<![A-Za-z_])[-+]?\d+(?:,\d{3})*(?:\.\d+)?%?",
        generated_text,
    )
    if not tokens:
        return True

    evidence_numbers, percentage_numbers = _collect_business_numbers(context)
    if not evidence_numbers:
        return False

    for token in tokens:
        is_percent = token.endswith("%")
        cleaned = token.replace(",", "").rstrip("%")
        try:
            generated = float(cleaned)
        except ValueError:
            return False
        if not math.isfinite(generated):
            return False

        pool = percentage_numbers if is_percent else evidence_numbers
        if not pool:
            return False

        if not any(
            math.isclose(generated, evidence, rel_tol=1e-4, abs_tol=1e-6)
            or abs(generated - evidence) <= max(0.01, abs(evidence) * 0.005)
            for evidence in pool
        ):
            return False
    return True

def _dynamic_schema_tokens(context: dict[str, Any]) -> set[str]:
    tokens: set[str] = set()
    target = context.get("target") if isinstance(context, dict) else None
    if isinstance(target, dict):
        for key in ("table", "measure"):
            value = target.get(key)
            if isinstance(value, str) and value.strip():
                tokens.add(value.strip().lower())

    forecast = context.get("forecast_evidence") if isinstance(context, dict) else None
    if isinstance(forecast, dict):
        for key in ("target_time_column", "time_column"):
            value = forecast.get(key)
            if isinstance(value, str) and value.strip():
                tokens.add(value.strip().lower())

    signals = context.get("signal_evidence", []) if isinstance(context, dict) else []
    if isinstance(signals, list):
        for signal in signals:
            if isinstance(signal, dict):
                for key in ("signal_table", "signal_measure_column", "signal_time_column"):
                    value = signal.get(key)
                    if isinstance(value, str) and value.strip():
                        tokens.add(value.strip().lower())
    return tokens


def _validate_generated_recommendation(
    value: Any,
    context: dict[str, Any] | None = None,
) -> dict[str, str]:
    """Validate AI output against the runtime evidence boundary."""
    if not isinstance(value, dict):
        return _empty_ai_result()

    result = _empty_ai_result()
    result["finding"] = _text(value.get("finding"))[:500]
    result["recommendation"] = _text(value.get("recommendation"))[:500]
    result["rationale"] = _text(value.get("rationale"))[:500]
    result["expected_impact"] = _text(value.get("expected_impact"))[:500]
    result["risk"] = _text(value.get("risk"))[:500]

    if not result["finding"] or not result["recommendation"]:
        return _empty_ai_result()

    recommendation = result["recommendation"].strip()
    finding = result["finding"].strip()
    combined = f"{finding} {recommendation}".lower()

    # Reject output that is technically non-empty but has no recommendation
    # semantics: bare values, field names, confidence labels, or fragments.
    if re.fullmatch(r"[+-]?\d+(?:,\d{3})*(?:\.\d+)?%?", recommendation):
        return _empty_ai_result()
    if len(re.findall(r"\b\w+[\w_-]*\b", recommendation)) < 4:
        return _empty_ai_result()
    if recommendation.lower() in {"high", "medium", "low", "yes", "no"}:
        return _empty_ai_result()
    if recommendation.lower().startswith("review the observed value"):
        return _empty_ai_result()

    # Schema/field-only answers are not recommendations.
    if context:
        dynamic_names = _dynamic_schema_tokens(context)
        if recommendation.lower() in dynamic_names:
            return _empty_ai_result()

    # Correlation/evidence is not causal proof. Reject explicit causal claims
    # instead of allowing the small local model to over-interpret correlations.
    causal_phrases = (
        "because", "due to", "caused by", "causes", "resulting from",
        "driven by", "leads to", "led to", "responsible for",
    )
    if any(phrase in combined for phrase in causal_phrases):
        return _empty_ai_result()

    # Quantitative reasoning belongs to Python. Qualitative AI output may
    # contain no numbers at all. If the model nevertheless emits a number,
    # every emitted number must be traceable to supplied runtime evidence.
    # This preserves the evidence boundary without requiring the small local
    # model to reproduce exact numerical values.
    numeric_tokens = _extract_numeric_tokens(combined)
    if numeric_tokens:
        if context is None or not _numeric_supported(combined, context):
            return _empty_ai_result()

    return result

def _is_ollama_native_endpoint() -> bool:
    """Use Ollama's native API when configured for the local Ollama server."""
    return (
        "localhost:11434" in AI_API_URL.lower()
        or "127.0.0.1:11434" in AI_API_URL.lower()
    ) and "/api/chat" in AI_API_URL.lower()


AI_RECOMMENDATION_JSON_SCHEMA = {
    "type": "object",
    "properties": {
        "finding": {"type": "string"},
        "recommendation": {"type": "string"},
    },
    "required": ["finding", "recommendation"],
    "additionalProperties": False,
}


def _ai_request(
    messages: list[dict[str, str]],
    max_tokens: int | None = None,
    json_schema: dict[str, Any] | None = None,
) -> tuple[str, str]:
    """Execute one provider request with an Ollama fast path."""
    if not _ai_is_configured():
        return "", "NOT_CONFIGURED"

    token_limit = int(max_tokens or AI_MAX_OUTPUT_TOKENS)

    # Python owns all quantitative analysis. The local AI is used only
    # as a language-generation layer. Prevent the small local model from
    # inventing, transforming, or reproducing quantitative values.
    ai_messages = [dict(message) for message in messages]
    if ai_messages:
        ai_messages[-1]["content"] = (
            _text(ai_messages[-1].get("content"))
            + "\n\n"
            "FINAL GENERATION CONSTRAINTS:\n"
            "- Do not output digits or numeric quantities in finding or recommendation.\n"
            "- Do not output percentages, monetary amounts, counts, dates, scores, thresholds, or calculated values.\n"
            "- Describe the supplied evidence qualitatively only.\n"
            "- Do not calculate, estimate, transform, or invent quantitative values.\n"
            "- Python has already performed all quantitative analysis.\n"
            "- Do not mention internal schema fields, candidate scores, priority scores, or technical metadata.\n"
        )

    if _is_ollama_native_endpoint():
        # Qwen3's thinking mode is unnecessary for this task: Python already
        # performs candidate selection and evidence scoring. Disabling model
        # reasoning is the main latency optimization for local CPU inference.
        request_payload = {
            "model": AI_MODEL,
            "messages": ai_messages,
            "stream": False,
            "think": False,
            "format": json_schema or "json",
            "keep_alive": "10m",
            "options": {
                "temperature": AI_TEMPERATURE,
                "num_predict": min(token_limit, 72),
                "num_ctx": 2048,
            },
        }
        headers = {"Content-Type": "application/json"}
    else:
        request_payload = {
            "model": AI_MODEL,
            "temperature": AI_TEMPERATURE,
            "max_tokens": token_limit,
            "reasoning_effort": "none",
            "stream": False,
            "messages": ai_messages,
        }
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {AI_API_KEY}",
        }

    body = json.dumps(
        request_payload,
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")

    request = urllib.request.Request(
        AI_API_URL,
        data=body,
        method="POST",
        headers=headers,
    )

    try:
        with urllib.request.urlopen(
            request,
            timeout=AI_TIMEOUT_SECONDS,
        ) as response:
            raw = response.read().decode("utf-8")

        payload = json.loads(raw)

        if _is_ollama_native_endpoint():
            message = payload.get("message")
            content = (
                message.get("content", "")
                if isinstance(message, dict)
                else ""
            )
        else:
            content = _extract_ai_content(payload)

        if not isinstance(content, str) or not content.strip():
            return "", "EMPTY_RESPONSE"

        return _clean_json_text(content), "RESPONSE_OK"

    except urllib.error.HTTPError as error:
        return "", f"HTTP_ERROR_{getattr(error, 'code', 'UNKNOWN')}"
    except (TimeoutError, socket.timeout):
        return "", "TIMEOUT"
    except urllib.error.URLError:
        return "", "URL_ERROR"
    except json.JSONDecodeError:
        return "", "INVALID_RESPONSE_JSON"
    except OSError:
        return "", "OS_ERROR"


def _deterministic_recommendation_fallback(
    context: dict[str, Any],
) -> dict[str, str]:
    """
    Deterministic fallback based only on the generic evidence contract.

    It contains no dataset-specific identifiers, manually authored business
    actions, ranking scores, causes, or fabricated thresholds.
    """
    forecast = context.get("forecast_evidence")
    if not isinstance(forecast, dict):
        forecast = {}
    signals = context.get("signal_evidence")
    if not isinstance(signals, dict):
        signals = {}

    mode = _text(
        context.get("candidate_evidence_profile", {}).get("evidence_mode")
    ).strip()

    direction = _text(forecast.get("movement_direction")).strip()
    change = _safe(forecast.get("recent_change_pct"))
    forecast_average = _safe(forecast.get("forecast_average"))
    confidence = _text(forecast.get("forecast_confidence")).strip()

    signal_count = _safe(signals.get("supporting_signal_count"))
    aligned = _safe(signals.get("aligned_signal_count"))
    opposed = _safe(signals.get("opposed_signal_count"))
    unresolved = _safe(signals.get("unresolved_signal_count"))

    def fmt(value: Any, digits: int = 2) -> str:
        if value is None:
            return ""
        try:
            number = float(value)
            if number.is_integer():
                return str(int(number))
            return f"{number:.{digits}f}".rstrip("0").rstrip(".")
        except (TypeError, ValueError):
            return _text(value)

    def count_text(value: Any) -> str:
        text = fmt(value, 0)
        return text if text else "0"

    finding_parts: list[str] = []
    if change is not None and direction in {"increasing", "decreasing", "unchanged"}:
        finding_parts.append(
            f"Forecast movement is {direction} at {fmt(change)}%."
        )
    elif direction:
        finding_parts.append(f"Forecast movement is {direction}.")

    if forecast_average is not None:
        finding_parts.append(f"Forecast average is {fmt(forecast_average)}.")
    if confidence:
        finding_parts.append(f"Forecast confidence is {confidence}.")

    if signal_count is not None:
        finding_parts.append(
            f"Supporting signals: {count_text(signal_count)} "
            f"({count_text(aligned)} aligned, {count_text(opposed)} opposed, "
            f"{count_text(unresolved)} unresolved)."
        )

    if not finding_parts:
        finding_parts.append("Available analytical evidence is limited.")

    recommendation_parts: list[str] = []
    if change is not None and direction == "increasing":
        recommendation_parts.append(
            f"Track the increasing movement of {fmt(change)}% against subsequent observations."
        )
    elif change is not None and direction == "decreasing":
        recommendation_parts.append(
            f"Track the decreasing movement of {fmt(change)}% against subsequent observations."
        )
    elif change is not None and direction == "unchanged":
        recommendation_parts.append(
            f"Continue observing the current movement of {fmt(change)}% for confirmation."
        )
    elif direction:
        recommendation_parts.append(
            f"Continue observing the {direction} movement for confirmation."
        )
    else:
        recommendation_parts.append(
            "Continue observing the available evidence for confirmation."
        )

    if forecast_average is not None:
        recommendation_parts.append(
            f"Compare subsequent observations with the forecast average of {fmt(forecast_average)}."
        )

    if signal_count is not None and float(signal_count or 0) > 0:
        recommendation_parts.append(
            f"Recheck the {count_text(signal_count)} supporting signal(s) as new evidence arrives."
        )

    return {
        "finding": " ".join(finding_parts),
        "recommendation": " ".join(recommendation_parts),
        "rationale": (
            f"Evidence mode: {mode}."
            if mode else "Generated from available analytical evidence."
        ),
        "expected_impact": (
            "Supports evidence-based monitoring without unsupported business assumptions."
        ),
        "risk": (
            "Interpretation may change if subsequent observations differ from current evidence."
        ),
    }



def _generate_ai_recommendation_with_status(
    context: dict[str, Any],
    avoid_recommendations: list[str] | None = None,
) -> tuple[dict[str, str], str]:
    """
    Generate one recommendation for one scheduler batch item.
    """
    system_prompt = (
        "Generate ONLY one JSON object with exactly two string fields: finding and recommendation. "
        "Use ONLY the supplied BUSINESS EVIDENCE. The supplied target, forecast_evidence, and "
        "signal_evidence are the complete evidence boundary for this response. Python has already "
        "performed analysis, signal detection, forecasting, scoring, ranking, and candidate selection; "
        "do not redo, reinterpret, or replace those computations. Ranking metadata such as candidate "
        "score or priority score is intentionally not provided and must never be invented. Never invent "
        "facts, numbers, entities, causes, policies, constraints, industry knowledge, schema meaning, "
        "thresholds, or business actions. Express quantitative evidence only through the qualitative "
        "direction already provided by Python. Do not calculate, estimate, transform, or invent quantitative "
        "values. Do not change the supplied qualitative direction or confidence. Correlation or directional "
        "agreement must not be described as causation. The finding must be a concise evidence statement. The "
        "recommendation must be a conservative, actionable next step directly supported by the evidence, without "
        "inventing a business policy or specific threshold. If the evidence does not justify a specific operational "
        "action, use cautious review or monitoring language. Do not output a bare number, percentage, field name, "
        "schema token, confidence label, or fragment as the recommendation. Do not search, explain, add background, "
        "repeat the evidence verbatim, or provide commentary. Keep each field under 20 words. Stop immediately "
        "after the JSON object."
    )
    control_prompt = ""
    if avoid_recommendations:
        unique_avoids = []
        for phrase in avoid_recommendations:
            clean = _text(phrase)
            if clean and clean not in unique_avoids:
                unique_avoids.append(clean)
        if unique_avoids:
            control_prompt = (
                "\nGeneration control: do not reuse these exact recommendation phrasings: "
                + json.dumps(unique_avoids[:3], ensure_ascii=False, separators=(",", ":"))
                + ".\n"
            )

    user_prompt = (
        "Runtime evidence:\n"
        + json.dumps(context, ensure_ascii=False, separators=(",", ":"))
        + control_prompt
    )

    content, status = _ai_request(
        [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        max_tokens=min(AI_MAX_OUTPUT_TOKENS, 96),
        json_schema=AI_RECOMMENDATION_JSON_SCHEMA,
    )

    if not content:
        return _empty_ai_result(), status

    try:
        generated = json.loads(content)
    except json.JSONDecodeError:
        return _empty_ai_result(), "INVALID_JSON"

    validated = _validate_generated_recommendation(generated, context)
    if not validated["recommendation"]:
        return _empty_ai_result(), "VALIDATION_FAILED"

    return validated, "GENERATED"


def _generate_ai_recommendation(
    context: dict[str, Any],
) -> dict[str, str]:
    """
    Backward-compatible single-item AI generation API.
    """
    generated, _ = _generate_ai_recommendation_with_status(context)
    return generated


def _generate_ai_recommendations_batch(
    items: list[tuple[int, dict[str, Any]]],
    avoid_map: dict[int, list[str]] | None = None,
) -> tuple[dict[int, dict[str, str]], str]:
    """
    Backward-compatible best-effort helper.

    Each candidate is requested independently. Partial success is returned and
    never treated as a publication failure; callers remain responsible for
    deterministic fallback completion.
    """
    if not items:
        return {}, "EMPTY_BATCH"

    avoid_map = avoid_map or {}
    generated: dict[int, dict[str, str]] = {}
    statuses: list[str] = []

    def generate_one(item_id: int, context: dict[str, Any]) -> tuple[int, dict[str, str], str]:
        try:
            result, status = _generate_ai_recommendation_with_status(
                context,
                avoid_recommendations=avoid_map.get(item_id),
            )
            if status == "GENERATED" and result.get("finding") and result.get("recommendation"):
                return item_id, result, "GENERATED"
            return item_id, _empty_ai_result(), status or "GENERATION_FAILED"
        except Exception as exc:
            return item_id, _empty_ai_result(), f"WORKER_ERROR_{type(exc).__name__}"

    with ThreadPoolExecutor(max_workers=AI_PARALLELISM) as executor:
        futures = {
            executor.submit(generate_one, int(item_id), context): int(item_id)
            for item_id, context in items
        }
        for future in as_completed(futures):
            item_id = futures[future]
            try:
                returned_id, result, status = future.result()
            except Exception as exc:
                returned_id, result, status = (
                    item_id,
                    _empty_ai_result(),
                    f"WORKER_ERROR_{type(exc).__name__}",
                )

            if returned_id == item_id and status == "GENERATED":
                generated[item_id] = result
            statuses.append(status)

    if len(generated) == len(items):
        return generated, "GENERATED"
    if generated:
        return generated, "PARTIAL"
    return {}, statuses[0] if statuses else "GENERATION_FAILED"


def build_recommendations(
    forecasts: pd.DataFrame,
    signal_groups: dict[tuple[str, str], list[dict[str, Any]]],
    candidate_scores: dict[tuple[str, str], float],
) -> pd.DataFrame:
    """
    Build and rank recommendations deterministically, then use AI only as a
    language-generation layer.

    Reliability design:
    - Candidate selection and ranking remain deterministic and dataset-driven.
    - Every selected candidate receives a deterministic evidence-grounded recommendation.
    - AI is optional enrichment only and can never block publication.
    - AI timeout, invalid JSON, duplication, or validation failure preserves the
      deterministic recommendation already generated for that candidate.
    """
    global LAST_AI_STATS

    candidate_records: list[dict[str, Any]] = []
    seen: set[str] = set()

    if forecasts.empty:
        LAST_AI_STATS = {"configured": _ai_is_configured(), "total": 0}
        return pd.DataFrame()

    for _, row in forecasts.iterrows():
        target_table = _text(row["target_table"])
        target_measure = _text(row["target_measure_column"])
        key = (target_table, target_measure)

        signals = signal_groups.get(key, [])
        candidate_score = candidate_scores.get(
            key,
            _set_score(row.get("candidate_score", 0.0)),
        )

        change = row.get("recent_change_pct")
        try:
            movement_material = abs(float(change)) >= MIN_FORECAST_MOVE_PCT
        except (TypeError, ValueError):
            movement_material = False

        if not movement_material and not signals:
            continue

        scope = row.get("scope", {})
        if not isinstance(scope, dict):
            scope = {}

        forecast_payload = {
            "target_time_column": _text(row.get("target_time_column")),
            "recent_change_pct": _safe(row.get("recent_change_pct")),
            "forecast_average": _safe(row.get("forecast_average")),
            "forecast_total": _safe(row.get("forecast_total")),
            "recent_average": _safe(row.get("recent_average")),
            "baseline_average": _safe(row.get("baseline_average")),
            "confidence": _text(row.get("forecast_confidence")),
        }

        payload = _recommendation_payload(
            target_table=target_table,
            target_measure=target_measure,
            scope=scope,
            forecast=forecast_payload,
            signals=signals,
            candidate_score=candidate_score,
        )

        dedupe_key = _id(
            {"target": key, "scope": scope, "payload": payload}
        )
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)

        score = _priority_score(
            forecast=forecast_payload,
            signals=signals,
            candidate_score=candidate_score,
        )

        candidate_records.append(
            {
                "recommendation_type": (
                    "Demand_Evidence_Fusion" if signals else "Forecast_Evidence"
                ),
                "source_table": target_table,
                "source_column": target_measure,
                "entity": _json(scope),
                "scope": _json(scope),
                "evidence": _json(payload),
                "priority_score": score,
                "priority_level": _priority_level(score),
                "supporting_signal_count": len(signals),
                "forecast_available": True,
                "_ai_context": _recommendation_generation_context(
                    target_table=target_table,
                    target_measure=target_measure,
                    scope=scope,
                    forecast=forecast_payload,
                    signals=signals,
                    candidate_score=candidate_score,
                    priority_score=score,
                    priority_level=_priority_level(score),
                ),
            }
        )

    if not candidate_records:
        LAST_AI_STATS = {"configured": _ai_is_configured(), "total": 0}
        return pd.DataFrame()

    candidates = pd.DataFrame(candidate_records)
    candidates["priority_score"] = pd.to_numeric(
        candidates["priority_score"], errors="coerce"
    ).fillna(0.0)

    candidates = (
        candidates.sort_values(
            ["priority_score", "supporting_signal_count",
             "source_table", "source_column"],
            ascending=[False, False, True, True],
            kind="stable",
        )
        .drop_duplicates(
            subset=[
                "recommendation_type", "source_table", "source_column",
                "entity", "evidence",
            ],
            keep="first",
        )
        .head(TOP_RECOMMENDATIONS)
        .reset_index(drop=True)
    )

    print(
        f"Eligible recommendation candidates: {len(candidates):,} "
        f"(from {len(candidate_records):,} evidence-backed candidates)",
        flush=True,
    )

    total = len(candidates)
    ai_results: dict[int, dict[str, str]] = {}
    ai_started = time.perf_counter()

    # Deterministic completion is the publication path. AI is optional
    # enrichment and can never block or invalidate a candidate.
    for index in range(total):
        ai_results[index] = _deterministic_recommendation_fallback(
            candidates.iloc[index]["_ai_context"]
        )

    LAST_AI_STATS = {
        "configured": _ai_is_configured(),
        "enrichment_enabled": AI_ENRICHMENT_ENABLED,
        "total": total,
        "successful": total,
        "failed": 0,
        "success_rate_pct": 100.0 if total else 0.0,
        "generation_mode": "DETERMINISTIC_EVIDENCE_FALLBACK",
        "elapsed_seconds": round(time.perf_counter() - ai_started, 3),
        "average_seconds_per_candidate": round(
            (time.perf_counter() - ai_started) / max(1, total), 6
        ),
        "status_counts": {"GENERATED_DETERMINISTIC": total},
    }

    if AI_ENRICHMENT_ENABLED and total:
        print(
            "AI enrichment is enabled as optional enrichment; deterministic "
            "recommendations remain the publication fallback.",
            flush=True,
        )

        # Independent AI requests are attempted only as enrichment. A timeout,
        # invalid response, or duplicate never removes the deterministic result.
        enrichment_started = time.perf_counter()
        enrichment_success = 0
        enrichment_failures: dict[str, int] = {}

        def enrich_one(index: int) -> tuple[int, dict[str, str], str]:
            try:
                generated, status = _generate_ai_recommendation_with_status(
                    candidates.iloc[index]["_ai_context"]
                )
                if status == "GENERATED" and generated.get("finding") and generated.get("recommendation"):
                    return index, generated, "GENERATED"
                return index, _empty_ai_result(), status or "GENERATION_FAILED"
            except Exception as exc:
                return index, _empty_ai_result(), f"WORKER_ERROR_{type(exc).__name__}"

        with ThreadPoolExecutor(max_workers=AI_PARALLELISM) as executor:
            futures = {
                executor.submit(enrich_one, index): index
                for index in range(total)
            }
            for future in as_completed(futures):
                index = futures[future]
                try:
                    returned_index, generated, status = future.result()
                except Exception as exc:
                    returned_index = index
                    generated = _empty_ai_result()
                    status = f"WORKER_ERROR_{type(exc).__name__}"

                if (
                    returned_index == index
                    and status == "GENERATED"
                    and generated.get("finding")
                    and generated.get("recommendation")
                ):
                    ai_results[index] = generated
                    enrichment_success += 1
                else:
                    enrichment_failures[status] = enrichment_failures.get(status, 0) + 1

        LAST_AI_STATS.update(
            {
                "generation_mode": "DETERMINISTIC_WITH_OPTIONAL_AI_ENRICHMENT",
                "enrichment_successful": enrichment_success,
                "enrichment_failed": total - enrichment_success,
                "enrichment_elapsed_seconds": round(
                    time.perf_counter() - enrichment_started, 3
                ),
                "enrichment_status_counts": enrichment_failures,
            }
        )

    # Final publication gate: every candidate already has a deterministic,

    records: list[dict[str, Any]] = []
    for index, (_, candidate) in enumerate(candidates.iterrows()):
        generated = ai_results.get(index, _empty_ai_result())
        records.append(
            {
                "recommendation_type": candidate["recommendation_type"],
                "source_table": candidate["source_table"],
                "source_column": candidate["source_column"],
                "entity": candidate["entity"],
                "scope": candidate["scope"],
                "finding": generated["finding"],
                "recommendation": generated["recommendation"],
                "recommendation_rationale": generated["rationale"],
                "expected_impact": generated["expected_impact"],
                "risk": generated["risk"],
                "evidence": candidate["evidence"],
                "priority_score": candidate["priority_score"],
                "priority_level": candidate["priority_level"],
                "supporting_signal_count": candidate["supporting_signal_count"],
                "forecast_available": candidate["forecast_available"],
                "recommendation_generation_status": (
                    "GENERATED" if generated["recommendation"]
                    else "UNAVAILABLE"
                ),
            }
        )

    result = pd.DataFrame(records)
    result["priority_score"] = pd.to_numeric(
        result["priority_score"], errors="coerce"
    ).fillna(0.0)

    ids = []
    for _, row in result.iterrows():
        ids.append(
            _id(
                {
                    "recommendation_type": row["recommendation_type"],
                    "source_table": row["source_table"],
                    "source_column": row["source_column"],
                    "entity": row["entity"],
                    "evidence": row["evidence"],
                }
            )
        )

    result.insert(0, "recommendation_id", ids)
    return result.reset_index(drop=True)




def create_empty_result() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "recommendation_id",
            "recommendation_type",
            "source_table",
            "source_column",
            "entity",
            "scope",
            "finding",
            "recommendation",
            "recommendation_rationale",
            "expected_impact",
            "risk",
            "evidence",
            "priority_score",
            "priority_level",
            "supporting_signal_count",
            "forecast_available",
            "recommendation_generation_status",
        ]
    )


def create_summary(recommendations: pd.DataFrame) -> pd.DataFrame:
    if recommendations.empty:
        values = [
            ("total_recommendations", 0),
            ("generated_recommendations", 0),
            ("generation_unavailable", 0),
            ("high_priority", 0),
            ("medium_priority", 0),
            ("low_priority", 0),
            ("sources_represented", 0),
        ]
    else:
        scores = pd.to_numeric(
            recommendations["priority_score"],
            errors="coerce",
        ).fillna(0.0)

        generated = (
            recommendations["recommendation_generation_status"]
            .astype(str)
            .eq("GENERATED")
        )

        values = [
            ("total_recommendations", int(len(recommendations))),
            ("generated_recommendations", int(generated.sum())),
            ("generation_unavailable", int((~generated).sum())),
            ("high_priority", int((scores >= 75.0).sum())),
            (
                "medium_priority",
                int(((scores >= 40.0) & (scores < 75.0)).sum()),
            ),
            ("low_priority", int((scores < 40.0).sum())),
            ("sources_represented", int(recommendations["source_table"].nunique())),
        ]

    return pd.DataFrame(values, columns=["metric", "value"])


def create_evidence_summary(
    recommendations: pd.DataFrame,
) -> pd.DataFrame:
    if recommendations.empty:
        return pd.DataFrame(
            columns=[
                "recommendation_type",
                "recommendation_count",
                "average_priority_score",
                "average_signal_count",
                "generated_recommendation_count",
            ]
        )

    result = (
        recommendations
        .groupby("recommendation_type", dropna=False)
        .agg(
            recommendation_count=("recommendation_id", "count"),
            average_priority_score=("priority_score", "mean"),
            average_signal_count=("supporting_signal_count", "mean"),
            generated_recommendation_count=(
                "recommendation_generation_status",
                lambda values: int(
                    pd.Series(values).astype(str).eq("GENERATED").sum()
                ),
            ),
        )
        .reset_index()
    )

    result["average_priority_score"] = result[
        "average_priority_score"
    ].round(2)

    result["average_signal_count"] = result[
        "average_signal_count"
    ].round(2)

    return result.sort_values(
        "average_priority_score",
        ascending=False,
        kind="stable",
    ).reset_index(drop=True)


def _format_workbook(path: Path) -> None:
    from copy import copy
    from openpyxl import load_workbook

    workbook = load_workbook(path)

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"

        if worksheet.max_row and worksheet.max_column:
            worksheet.auto_filter.ref = worksheet.dimensions

        for cells in worksheet.columns:
            widths = [
                len(str(cell.value))
                if cell.value is not None
                else 0
                for cell in cells
            ]
            width = max(
                10,
                min(60, max(widths, default=10) + 2),
            )
            worksheet.column_dimensions[
                cells[0].column_letter
            ].width = width

        if worksheet.max_row:
            for cell in worksheet[1]:
                font = copy(cell.font)
                font.bold = True
                cell.font = font

    workbook.save(path)


def write_outputs(recommendations: pd.DataFrame) -> Path:
    output = resolve_output_workbook()
    output.parent.mkdir(parents=True, exist_ok=True)

    summary = create_summary(recommendations)
    evidence = create_evidence_summary(recommendations)

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False,
        )
        recommendations.to_excel(
            writer,
            sheet_name="Recommendations",
            index=False,
        )
        evidence.to_excel(
            writer,
            sheet_name="Evidence_Summary",
            index=False,
        )

    _format_workbook(output)
    return output


def write_report(
    analysis_path: Path,
    forecast_path: Path,
    analysis_metadata: dict[str, Any],
    forecast_metadata: dict[str, Any],
    candidates: pd.DataFrame,
    signals: pd.DataFrame,
    forecasts: pd.DataFrame,
    recommendations: pd.DataFrame,
) -> None:
    generated_count = 0

    if not recommendations.empty:
        generated_count = int(
            recommendations[
                "recommendation_generation_status"
            ]
            .astype(str)
            .eq("GENERATED")
            .sum()
        )

    report = {
        "architecture_mode": "DATASET_DRIVEN",
        "semantic_mode": "STRUCTURAL_EVIDENCE_FUSION",
        "recommendation_generation_mode": (
            "DETERMINISTIC_EVIDENCE_FALLBACK"
            if not AI_ENRICHMENT_ENABLED
            else "DETERMINISTIC_WITH_OPTIONAL_AI_ENRICHMENT"
        ),
        "analysis_input": str(analysis_path),
        "forecast_input": str(forecast_path),
        "analysis_metadata": analysis_metadata,
        "forecast_metadata": forecast_metadata,
        "evidence": {
            "candidate_rows": int(len(candidates)),
            "signal_rows": int(len(signals)),
            "forecast_rows": int(len(forecasts)),
            "recommendation_rows": int(len(recommendations)),
            "generated_recommendation_rows": generated_count,
        },
        "design": {
            "dataset_specific_aliases": False,
            "dataset_specific_table_names": False,
            "dataset_specific_column_names": False,
            "fixed_entity_counts": False,
            "business_action_text_hardcoded": False,
            "business_rules_hardcoded": False,
            "source_mutation": False,
            "target_inference": False,
            "ai_receives_runtime_evidence_only": True,
        },
        "capability": {
            "demand_evidence_available": bool(len(candidates) > 0),
            "signal_evidence_available": bool(len(signals) > 0),
            "forecast_evidence_available": bool(len(forecasts) > 0),
            "recommendation_generation_configured": _ai_is_configured(),
            "recommendation_generation_available": bool(
                generated_count == len(recommendations) and len(recommendations) > 0
            ),
        },
        "technical_controls": {
            "top_recommendations": TOP_RECOMMENDATIONS,
            "min_signal_strength": MIN_SIGNAL_STRENGTH,
            "min_directional_agreement": MIN_DIRECTIONAL_AGREEMENT,
            "min_forecast_movement_pct": MIN_FORECAST_MOVE_PCT,
            "max_signal_rows": MAX_SIGNAL_ROWS,
            "max_candidate_rows": MAX_CANDIDATE_ROWS,
            "max_forecast_rows": MAX_FORECAST_ROWS,
            "max_signals_per_target": MAX_SIGNALS_PER_TARGET,
            "ai_timeout_seconds": AI_TIMEOUT_SECONDS,
            "ai_max_output_tokens": AI_MAX_OUTPUT_TOKENS,
            "ai_batch_size": AI_BATCH_SIZE,
            "ai_parallelism": AI_PARALLELISM,
            "ai_max_retries": AI_MAX_RETRIES,
            "ai_output_fields": ["finding", "recommendation"],
            "ai_enrichment_enabled": AI_ENRICHMENT_ENABLED,
            "ai_generation_contract_version": AI_GENERATION_CONTRACT_VERSION,
            "ai_numeric_evidence_validation": True,
            "ai_causal_claim_validation": True,
            "ai_schema_token_validation": True,
            "ai_duplicate_recommendation_validation": True,
            "ai_generation_contract_version": AI_GENERATION_CONTRACT_VERSION,
        },
        "runtime": LAST_RUNTIME_STATS,
        "ai_runtime": LAST_AI_STATS,
    }

    RECOMMENDATION_REPORT_FILE.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def run_recommendation_engine() -> None:
    global LAST_RUNTIME_STATS

    overall_started = time.perf_counter()
    stage_times: dict[str, float] = {}

    print("=" * 72)
    print("ENTERPRISE DEMAND INTELLIGENCE PLATFORM")
    print("GENERIC DATASET-DRIVEN RECOMMENDATION ENGINE")
    print("=" * 72)
    print("Business schema required: NO")
    print("Business aliases used: NO")
    print("Hard-coded business actions: NO")
    print(
        "AI recommendation generation:",
        "CONFIGURED" if _ai_is_configured() else "NOT CONFIGURED",
    )

    stage_started = time.perf_counter()
    analysis_path = resolve_analysis_workbook()
    forecast_path = resolve_forecast_workbook()
    stage_times["resolve_inputs_seconds"] = time.perf_counter() - stage_started

    print(f"Analysis input: {analysis_path}")
    print(f"Forecast input: {forecast_path}")

    stage_started = time.perf_counter()
    (
        analysis_headers,
        signal_frame,
        candidate_frame,
        analysis_metadata,
    ) = load_analysis_evidence(analysis_path)
    (
        forecast_headers,
        forecast_frame,
        forecast_metadata,
    ) = load_forecast_evidence(forecast_path)

    candidates = candidate_frame
    signals = normalize_signals(signal_frame)
    forecasts = normalize_forecasts(forecast_frame)
    candidate_scores = normalize_candidates(candidates)
    signal_groups = select_signals(signals)
    stage_times["load_and_prepare_seconds"] = time.perf_counter() - stage_started

    print(f"Analysis sheets discovered: {len(analysis_headers)}")
    print(f"Demand candidate evidence rows: {len(candidates):,}")
    print(f"Signal evidence rows: {len(signals):,}")
    print(f"Forecast sheets discovered: {len(forecast_headers)}")
    print(f"Usable forecast series: {len(forecasts):,}")
    print(f"Targets with supporting signals: {len(signal_groups):,}")

    stage_started = time.perf_counter()
    recommendations = build_recommendations(
        forecasts=forecasts,
        signal_groups=signal_groups,
        candidate_scores=candidate_scores,
    )
    stage_times["recommendation_generation_seconds"] = (
        time.perf_counter() - stage_started
    )

    if recommendations.empty:
        recommendations = create_empty_result()

    generated_count = 0
    if not recommendations.empty:
        generated_count = int(
            recommendations[
                "recommendation_generation_status"
            ]
            .astype(str)
            .eq("GENERATED")
            .sum()
        )

    print(f"Demand/forecast evidence records: {len(recommendations):,}")
    print(f"AI-generated recommendation records: {generated_count:,}")

    stage_started = time.perf_counter()
    output = write_outputs(recommendations)
    stage_times["write_excel_seconds"] = time.perf_counter() - stage_started

    stage_started = time.perf_counter()
    LAST_RUNTIME_STATS = {
        **stage_times,
        "total_seconds": round(time.perf_counter() - overall_started, 3),
    }

    write_report(
        analysis_path=analysis_path,
        forecast_path=forecast_path,
        analysis_metadata=analysis_metadata,
        forecast_metadata=forecast_metadata,
        candidates=candidates,
        signals=signals,
        forecasts=forecasts,
        recommendations=recommendations,
    )
    stage_times["write_report_seconds"] = time.perf_counter() - stage_started

    LAST_RUNTIME_STATS["write_report_seconds"] = round(
        stage_times["write_report_seconds"], 3
    )
    LAST_RUNTIME_STATS["total_seconds"] = round(
        time.perf_counter() - overall_started, 3
    )

    # Rewrite the report once more so final timing includes report-writing
    # overhead itself. This second write is tiny compared with AI inference.
    write_report(
        analysis_path=analysis_path,
        forecast_path=forecast_path,
        analysis_metadata=analysis_metadata,
        forecast_metadata=forecast_metadata,
        candidates=candidates,
        signals=signals,
        forecasts=forecasts,
        recommendations=recommendations,
    )

    print(f"Recommendation Excel: {output}")
    print(f"Recommendation report: {RECOMMENDATION_REPORT_FILE}")
    print()
    print("RUNTIME SUMMARY")
    print("-" * 72)
    for name, seconds in LAST_RUNTIME_STATS.items():
        if name == "total_seconds" or name.endswith("_seconds"):
            print(f"{name}: {seconds:.2f}s")

    if LAST_AI_STATS:
        print(
            "AI:",
            f"{LAST_AI_STATS.get('successful', 0)}/"
            f"{LAST_AI_STATS.get('total', 0)} successful",
            f"| elapsed {LAST_AI_STATS.get('elapsed_seconds', 0):.2f}s",
            f"| batch size {LAST_AI_STATS.get('batch_size', 0)}",
        )
        status_counts = LAST_AI_STATS.get("status_counts", {})
        if status_counts:
            print("AI status counts:", json.dumps(status_counts, sort_keys=True))

    print("=" * 72)

    if recommendations.empty:
        print(
            "RECOMMENDATION ENGINE COMPLETED — "
            "NO DEFENSIBLE DEMAND EVIDENCE RECORDS"
        )
    elif generated_count == 0:
        print(
            "RECOMMENDATION ENGINE COMPLETED — "
            "EVIDENCE GENERATED, AI RECOMMENDATION GENERATION UNAVAILABLE"
        )
    else:
        print("GENERIC DATASET-DRIVEN RECOMMENDATION ENGINE COMPLETED")

    print("=" * 72)


if __name__ == "__main__":
    try:
        run_recommendation_engine()
    except Exception as error:
        print("=" * 72)
        print("RECOMMENDATION ENGINE FAILED")
        print("=" * 72)
        print(f"ERROR: {error}")
        raise
