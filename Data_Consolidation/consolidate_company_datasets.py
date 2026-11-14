"""
DATASET-INDEPENDENT COMPANY DATA CONSOLIDATION UTILITY

Python: 3.11+

Required for semantic matching:
    sentence-transformers (downloads all-MiniLM-L6-v2 on first run)

INPUT:
    Data_Consolidation/input/

OUTPUT:
    Data_Consolidation/output/Consolidated_Company_Dataset.xlsx

OUTPUT WORKSHEET:
    All_Consolidated_Data

Supported formats:
    JSON, CSV, XLSX, XLS, PARQUET

CORE DESIGN:
    1. Discover every supported input file recursively.
    2. Read every record from every discovered dataset.
    3. Dynamically discover columns and conservatively harmonize only high-confidence equivalent headers.
    4. Append every record into one master worksheet.
    5. Preserve source provenance.
    6. Flatten nested JSON objects.
    7. Preserve nested arrays as JSON text in cells.
    8. Never create separate entity/company worksheets.
    9. Never modify original input files.
    10. Safely publish the output workbook.
    11. Never depend on a fixed synonym/alias dictionary for column matching.
    12. Keep the output limited to the requested Excel workbook; no sidecar CSV files.
    13. Use native worksheet AutoFilter (no fragile structured Table XML).
    14. Freeze only the header row so no vertical freeze divider appears.
"""

from __future__ import annotations

import json
import math
import os
import tempfile
from collections import OrderedDict
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

import pandas as pd
import re
import warnings

from datetime import date, datetime, time, timezone
from decimal import Decimal, InvalidOperation
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


# ============================================================
# 1. CENTRAL CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

OUTPUT_FILE = (
    OUTPUT_DIR / "Consolidated_Company_Dataset.xlsx"
)

OUTPUT_SHEET = "All_Consolidated_Data"

SUPPORTED_FORMATS = {
    ".json",
    ".csv",
    ".xlsx",
    ".xls",
    ".parquet"
}

FIELD_NOT_PRESENT = (
    "This field is not present in this company's dataset"
)

VALUE_NOT_PROVIDED = (
    "Not provided in the original record"
)

MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLUMNS = 16_384
MAX_CELL_LENGTH = 32_767

# ============================================================
# DATA QUALITY AND DATE CONFIGURATION
# ============================================================

DATE_DISPLAY_FORMAT = "yyyy-mm-dd"

MIN_DATE_PROFILE_RATIO = 0.80

SEMANTIC_MODEL_REQUIRED = False

METADATA_COLUMNS = [
    "__meta_company_source",
    "__meta_source_file",
    "__meta_source_path",
    "__meta_source_format",
    "__meta_source_worksheet_or_collection",
    "__meta_source_record_number",
    "__meta_record_granularity",
]

METADATA_COLUMNS = [
    "__meta_company_source",
    "__meta_source_file",
    "__meta_source_path",
    "__meta_source_format",
    "__meta_source_worksheet_or_collection",
    "__meta_source_record_number",
    "__meta_record_granularity",
]


# ============================================================
# 2. EXCEL FORMATTING
# ============================================================

BODY_FONT = Font(
    name="Aptos",
    size=10
)

BODY_ALIGNMENT = Alignment(
    horizontal="left",
    vertical="top"
)


# ============================================================
# 3. GENERAL VALUE HANDLING
# ============================================================

def is_missing(value: Any) -> bool:
    """Identify null and NaN values safely."""

    if value is None:
        return True

    if isinstance(value, (dict, list, tuple, set)):
        return False

    try:
        result = pd.isna(value)

        if isinstance(result, bool):
            return result

        if hasattr(result, "item"):
            return bool(result.item())

    except Exception:
        pass

    return False


def convert_value(value: Any) -> Any:
    """Convert values into safe Python representations."""

    if is_missing(value):
        return None

    if isinstance(value, dict):
        return {
            str(key): convert_value(item)
            for key, item in value.items()
        }

    if isinstance(value, (list, tuple)):
        return [
            convert_value(item)
            for item in value
        ]

    if isinstance(value, set):
        return [
            convert_value(item)
            for item in sorted(value, key=str)
        ]

    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    if isinstance(value, (datetime, date, time)):
        return value

    if hasattr(value, "item") and callable(value.item):
        try:
            return convert_value(value.item())
        except Exception:
            pass

    if isinstance(value, float) and not math.isfinite(value):
        return None

    if isinstance(value, (str, int, float, bool)):
        return value

    return str(value)


# ============================================================
# 4. NESTED DATA HANDLING
# ============================================================

def flatten_record(
    record: Any,
    prefix: str = ""
) -> OrderedDict:

    """
    Flatten nested dictionaries.

    Example:
        {
            "employee": {
                "name": "Amit",
                "address": {
                    "city": "Kolkata"
                }
            }
        }

    Becomes:
        employee.name
        employee.address.city

    Nested arrays are preserved as JSON text in one cell.
    No additional worksheets are generated.
    """

    result = OrderedDict()

    if isinstance(record, dict):

        for key, value in record.items():

            key = str(key).strip() or "unnamed_field"

            column = (
                f"{prefix}.{key}"
                if prefix
                else key
            )

            if isinstance(value, dict):

                nested = flatten_record(
                    value,
                    column
                )

                result.update(nested)

            elif isinstance(value, (list, tuple, set)):

                result[column] = json.dumps(
                    convert_value(value),
                    ensure_ascii=False,
                    default=str
                )

            else:

                result[column] = convert_value(value)

    else:

        result["value"] = convert_value(record)

    return result


# ============================================================
# 5. COMPANY AND SOURCE IDENTIFICATION
# ============================================================

def get_company_source(file_path: Path) -> str:

    """
    Use the first folder under input/ as the company/source label.

    If a file is directly inside input/, use its filename.

    This is only provenance metadata.
    No company-specific assumptions are made.
    """

    try:

        relative = file_path.relative_to(INPUT_DIR)

        if len(relative.parts) > 1:
            return relative.parts[0]

    except ValueError:
        pass

    return file_path.stem


def get_relative_path(file_path: Path) -> str:

    try:

        return str(
            file_path.relative_to(INPUT_DIR)
        ).replace("\\", "/")

    except ValueError:

        return file_path.name


# ============================================================
# 6. JSON READER
# ============================================================

def read_json(file_path: Path) -> list:

    """
    Supports:

    A. JSON list:
       [
           {"id": 1},
           {"id": 2}
       ]

    B. JSON object containing multiple collections:
       {
           "employees": [...],
           "customers": [...]
       }

    C. Single JSON object:
       {"id": 1, "name": "Amit"}

    D. Primitive JSON values.

    Each top-level collection becomes records in the SAME
    consolidated worksheet.

    Shared metadata fields are included in each collection record.
    """

    with open(
        file_path,
        "r",
        encoding="utf-8-sig"
    ) as file:

        data = json.load(file)

    output = []

    if isinstance(data, list):

        for record in data:

            output.append(
                (file_path.stem, record)
            )

        return output

    if isinstance(data, dict):

        collections = {}

        shared_fields = {}

        for key, value in data.items():

            if (
                isinstance(value, list)
                and all(
                    isinstance(item, dict)
                    for item in value
                )
            ):

                collections[key] = value

            else:

                shared_fields[key] = value

        if collections:

            for collection_name, records in collections.items():

                for record in records:

                    merged_record = dict(shared_fields)

                    merged_record.update(record)

                    output.append(
                        (str(collection_name), merged_record)
                    )

            return output

        return [
            (file_path.stem, data)
        ]

    return [
        (file_path.stem, {"value": data})
    ]


# ============================================================
# 7. CSV READER
# ============================================================


def read_csv(file_path: Path) -> list:

    """
    Read CSV files while allowing pandas to infer numeric types.

    Integer and decimal values should remain numeric.
    Missing-value descriptions remain text.
    """

    try:

        dataframe = pd.read_csv(
            file_path,
            encoding="utf-8-sig",
            keep_default_na=False
        )

    except UnicodeDecodeError:

        dataframe = pd.read_csv(
            file_path,
            encoding="latin-1",
            keep_default_na=False
        )

    records = dataframe.to_dict(
        orient="records"
    )

    return [
        (file_path.stem, record)
        for record in records
    ]



# ============================================================
# 8. EXCEL READER
# ============================================================

def read_excel(file_path: Path) -> list:

    """
    Reads all worksheets from XLSX and XLS files.

    Every worksheet's records are appended into the SAME
    final worksheet.
    """

    sheets = pd.read_excel(
        file_path,
        sheet_name=None,
        dtype=object
    )

    output = []

    for sheet_name, dataframe in sheets.items():

        records = dataframe.to_dict(
            orient="records"
        )

        for record in records:

            output.append(
                (str(sheet_name), record)
            )

    return output


# ============================================================
# 9. PARQUET READER
# ============================================================

def read_parquet(file_path: Path) -> list:

    dataframe = pd.read_parquet(file_path)

    records = dataframe.to_dict(
        orient="records"
    )

    return [
        (file_path.stem, record)
        for record in records
    ]


# ============================================================
# 10. UNIVERSAL FILE READER
# ============================================================

def read_source_file(file_path: Path) -> list:

    extension = file_path.suffix.lower()

    if extension == ".json":
        return read_json(file_path)

    if extension == ".csv":
        return read_csv(file_path)

    if extension in {".xlsx", ".xls"}:
        return read_excel(file_path)

    if extension == ".parquet":
        return read_parquet(file_path)

    return []


# ============================================================
# 11. INPUT FILE DISCOVERY
# ============================================================

def discover_input_files() -> list:

    INPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    files = []

    for path in INPUT_DIR.rglob("*"):

        if not path.is_file():
            continue

        if path.suffix.lower() not in SUPPORTED_FORMATS:
            continue

        if path.name.startswith(("~$", ".")):
            continue

        files.append(path)

    return sorted(
        files,
        key=lambda item: str(item).lower()
    )


# ============================================================
# 11A. CONSERVATIVE SEMANTIC COLUMN HARMONIZATION
# ============================================================

# Semantic similarity is only a candidate generator. A merge is accepted
# only when the headers are strongly similar, profiles are compatible,
# the fields never coexist in the same source schema, and the match is
# mutual-best with a sufficient margin. Uncertain columns stay separate.
SEMANTIC_MODEL_NAME = "sentence-transformers/all-MiniLM-L6-v2"
SEMANTIC_MATCH_THRESHOLD = 0.84
SEMANTIC_MATCH_MARGIN = 0.05
SEMANTIC_MODEL_REQUIRED = True


def _normalized_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).casefold())


def _value_family(values: list[Any]) -> str:
    observed = [v for v in values if not is_missing(v)][:200]
    if not observed:
        return "unknown"
    if all(isinstance(v, bool) for v in observed):
        return "boolean"
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in observed):
        return "numeric"
    if all(isinstance(v, (datetime, date, time, pd.Timestamp)) for v in observed):
        return "datetime"
    if all(isinstance(v, str) for v in observed):
        # Do not mistake ordinary numeric-looking text identifiers for numbers.
        parsed = pd.to_datetime(pd.Series(observed[:40]), errors="coerce", format="mixed")
        if len(observed) >= 4 and parsed.notna().mean() >= 0.9:
            return "datetime_text"
        return "text"
    return "mixed"


def _compatible_column_profiles(left_values: list[Any], right_values: list[Any]) -> bool:
    left_family = _value_family(left_values)
    right_family = _value_family(right_values)
    if "unknown" in {left_family, right_family} or left_family == right_family:
        return True
    # Numeric and textual columns are not merged automatically. This is
    # deliberately conservative because text may contain IDs, units or codes.
    return False


def _headers_cooccur(left: str, right: str, schemas: list[set]) -> bool:
    """Two distinct fields in one source schema must never be collapsed."""
    return any(left in schema and right in schema for schema in schemas)

# ============================================================
# GENERIC SOURCE-SCOPED DATE NORMALIZATION
# ============================================================

def is_date_column(column_name: str) -> bool:
    """
    Detect explicit date/time field names without using
    company-specific aliases.
    """

    tokens = set(
        re.findall(r"[a-z0-9]+", str(column_name).lower())
    )

    compact = "".join(tokens)

    return (
        bool(tokens.intersection({
            "date",
            "datetime",
            "timestamp",
        }))
        or compact.endswith((
            "date",
            "datetime",
            "timestamp",
        ))
    )


def parse_iso_date(value: str):
    """Parse ISO year-first date and datetime strings."""

    if not re.match(
        r"^\d{4}-\d{2}-\d{2}(?:$|[T ])",
        value.strip()
    ):
        return None

    try:
        parsed = datetime.fromisoformat(
            value.strip().replace("Z", "+00:00")
        )

        if parsed.tzinfo is not None:
            parsed = (
                parsed.astimezone(timezone.utc)
                .replace(tzinfo=None)
            )

        return parsed

    except ValueError:
        return None


def numeric_date_parts(value: str):
    """
    Extract numeric date components.

    Supports:
        DD/MM/YYYY
        MM/DD/YYYY
        DD-MM-YYYY
        YYYY is mandatory.

    Does not independently guess ambiguous date order.
    """

    match = re.fullmatch(
        r"\s*(\d{1,2})([/.-])(\d{1,2})\2(\d{4})\s*",
        value.strip()
    )

    if not match:
        return None

    return (
        int(match.group(1)),
        int(match.group(3)),
        int(match.group(4)),
    )


def detect_date_convention(values: list) -> str | None:
    """
    Infer DMY or MDY only from unambiguous observations
    belonging to the same source field.
    """

    evidence = set()

    for value in values:

        if not isinstance(value, str):
            continue

        parts = numeric_date_parts(value)

        if not parts:
            continue

        first, second, _ = parts

        if first > 12 and 1 <= second <= 12:
            evidence.add("DMY")

        elif second > 12 and 1 <= first <= 12:
            evidence.add("MDY")

    if len(evidence) == 1:
        return next(iter(evidence))

    return None


def parse_numeric_date(
    value: str,
    convention: str | None
):
    """Parse numeric dates only when their order is established."""

    parts = numeric_date_parts(value)

    if not parts:
        return None

    first, second, year = parts

    if first > 12 and second <= 12:
        day, month = first, second

    elif second > 12 and first <= 12:
        month, day = first, second

    elif convention == "DMY":
        day, month = first, second

    elif convention == "MDY":
        month, day = first, second

    else:
        return None

    try:
        return datetime(year, month, day)

    except ValueError:
        return None


def normalize_source_dates(
    rows: list[dict],
    schemas: list[set],
):
    """
    Normalize date values independently for each source file,
    worksheet/collection and original field.

    Ambiguous dates without reliable evidence remain unchanged.
    """

    if len(rows) != len(schemas):
        raise ValueError(
            "Date normalization stopped: row/schema mismatch."
        )

    source_groups = defaultdict(list)

    for index, row in enumerate(rows):

        source_key = (
            row.get("__meta_source_path"),
            row.get("__meta_source_worksheet_or_collection"),
        )

        source_groups[source_key].append(index)

    messages = []

    for source_key, indexes in source_groups.items():

        columns = set()

        for index in indexes:
            columns.update(schemas[index])

        for column in columns:

            values = [
                rows[index].get(column)
                for index in indexes
                if column in schemas[index]
                and not is_missing(rows[index].get(column))
            ]

            if not values:
                continue

            native_dates = sum(
                isinstance(
                    value,
                    (datetime, date, pd.Timestamp)
                )
                for value in values
            )

            shaped_dates = 0

            for value in values:

                if isinstance(
                    value,
                    (datetime, date, pd.Timestamp)
                ):
                    shaped_dates += 1

                elif isinstance(value, str):

                    if (
                        parse_iso_date(value) is not None
                        or numeric_date_parts(value) is not None
                    ):
                        shaped_dates += 1

            detected = (
                is_date_column(column)
                or native_dates > 0
                or (
                    shaped_dates > 0
                    and shaped_dates / len(values)
                    >= MIN_DATE_PROFILE_RATIO
                )
            )

            if not detected:
                continue

            convention = detect_date_convention(values)

            for index in indexes:

                if column not in schemas[index]:
                    continue

                value = rows[index].get(column)

                if is_missing(value):
                    continue

                if isinstance(value, pd.Timestamp):

                    rows[index][column] = value.to_pydatetime()

                elif isinstance(value, datetime):

                    if value.tzinfo is not None:

                        rows[index][column] = (
                            value.astimezone(timezone.utc)
                            .replace(tzinfo=None)
                        )

                elif isinstance(value, date):

                    rows[index][column] = datetime.combine(
                        value,
                        datetime.min.time()
                    )

                elif isinstance(value, str):

                    parsed = parse_iso_date(value)

                    if parsed is None:
                        parsed = parse_numeric_date(
                            value,
                            convention
                        )

                    if parsed is not None:
                        rows[index][column] = parsed

                    elif numeric_date_parts(value):
                        messages.append(
                            f"Ambiguous date retained as text: "
                            f"{source_key}, {column}, {value}"
                        )

    return list(dict.fromkeys(messages))

def build_semantic_column_map(
    rows: list[dict],
    schemas: list[set],
) -> tuple[dict[str, str], list[str], list[dict[str, Any]]]:
    """Build a high-precision, dataset-derived column map and audit trail."""
    source_columns: list[str] = []
    for schema in schemas:
        for column in sorted(schema, key=lambda x: str(x).casefold()):
            if column not in source_columns and column not in METADATA_COLUMNS:
                source_columns.append(column)

    identity_map = {column: column for column in source_columns}
    audit: list[dict[str, Any]] = []
    warnings: list[str] = []
    if len(source_columns) < 2:
        return identity_map, warnings, audit

    values_by_column = {column: [] for column in source_columns}
    for row, schema in zip(rows, schemas):
        for column in schema:
            if column in values_by_column and column in row:
                value = row[column]
                if not is_missing(value) and len(values_by_column[column]) < 250:
                    values_by_column[column].append(value)

    # Identical normalized headers are safe only if they do not coexist in
    # one schema (e.g. punctuation/case variants from different files).
    candidates: dict[tuple[str, str], float] = {}
    for i, left in enumerate(source_columns):
        for right in source_columns[i + 1:]:
            if _headers_cooccur(left, right, schemas):
                audit.append({
                    "Source_Column_A": left, "Source_Column_B": right,
                    "Similarity": None, "Decision": "REJECTED",
                    "Reason": "Both fields occur in the same source schema; kept separate.",
                    "Canonical_Column": "",
                })
                continue
            if not _compatible_column_profiles(values_by_column[left], values_by_column[right]):
                audit.append({
                    "Source_Column_A": left, "Source_Column_B": right,
                    "Similarity": None, "Decision": "REJECTED",
                    "Reason": "Observed value-type profiles are incompatible.",
                    "Canonical_Column": "",
                })
                continue
            if _normalized_header(left) == _normalized_header(right):
                candidates[(left, right)] = 1.0

    # Only embed headers, never sample values. Values can describe different
    # companies/regions/products and can distort the meaning of a field name.
    try:
        from sentence_transformers import SentenceTransformer
        from sentence_transformers.util import cos_sim
        from huggingface_hub import snapshot_download

        # Prefer the local model cache first. Once downloaded, normal runs do
        # not need an internet connection or Hugging Face Hub checks. If the
        # cache is empty, fall back to the initial online download.
        try:
            model_path = snapshot_download(
                repo_id=SEMANTIC_MODEL_NAME,
                local_files_only=True,
            )
            print("Using cached semantic model (offline-capable mode).")
        except Exception:
            print("Semantic model is not cached; downloading it from Hugging Face...")
            model_path = snapshot_download(
                repo_id=SEMANTIC_MODEL_NAME,
                local_files_only=False,
            )

        model = SentenceTransformer(model_path)
        header_texts = [f"Database column name: {name.replace('_', ' ').replace('-', ' ')}" for name in source_columns]
        embeddings = model.encode(
            header_texts, convert_to_tensor=True,
            normalize_embeddings=True, show_progress_bar=False,
        )
        similarity_matrix = cos_sim(embeddings, embeddings).cpu().numpy()
    except Exception as exc:
        if SEMANTIC_MODEL_REQUIRED:
            raise RuntimeError(
                "Semantic matching could not load the cached language model. "
                "Run once with internet access to download it, then retry. "
                f"Underlying error: {type(exc).__name__}: {exc}"
            ) from exc
        similarity_matrix = None
        warnings.append("Semantic model unavailable; exact normalized headers only were harmonized.")

    # Similarity scores and constraints for every pair.
    pair_scores: dict[tuple[str, str], float] = dict(candidates)
    for i, left in enumerate(source_columns):
        for j in range(i + 1, len(source_columns)):
            right = source_columns[j]
            if _headers_cooccur(left, right, schemas):
                continue
            if not _compatible_column_profiles(values_by_column[left], values_by_column[right]):
                continue
            score = float(similarity_matrix[i][j]) if similarity_matrix is not None else 0.0
            pair_scores[(left, right)] = max(pair_scores.get((left, right), 0.0), score)

    def score_between(a: str, b: str) -> float:
        key = (a, b) if source_columns.index(a) < source_columns.index(b) else (b, a)
        return pair_scores.get(key, -1.0)

    # Candidate must be the unique mutual-best match and exceed a high
    # threshold with a margin over each side's second-best candidate.
    accepted: list[tuple[str, str, float]] = []
    for (left, right), score in pair_scores.items():
        if score < SEMANTIC_MATCH_THRESHOLD:
            continue
        left_scores = sorted((v for (a, b), v in pair_scores.items() if a == left or b == left), reverse=True)
        right_scores = sorted((v for (a, b), v in pair_scores.items() if a == right or b == right), reverse=True)
        left_best = max(((b if a == left else a, v) for (a, b), v in pair_scores.items() if a == left or b == left), key=lambda x: x[1], default=(None, -1))[0]
        right_best = max(((b if a == right else a, v) for (a, b), v in pair_scores.items() if a == right or b == right), key=lambda x: x[1], default=(None, -1))[0]
        left_second = left_scores[1] if len(left_scores) > 1 else -1.0
        right_second = right_scores[1] if len(right_scores) > 1 else -1.0
        if left_best == right and right_best == left and score - left_second >= SEMANTIC_MATCH_MARGIN and score - right_second >= SEMANTIC_MATCH_MARGIN:
            accepted.append((left, right, score))
        else:
            audit.append({
                "Source_Column_A": left, "Source_Column_B": right,
                "Similarity": round(score, 4), "Decision": "REVIEW / KEPT SEPARATE",
                "Reason": "Not a sufficiently unambiguous mutual-best match.",
                "Canonical_Column": "",
            })

    # Complete-link clustering: every pair in a proposed group must pass
    # threshold and no two members may co-occur in one source schema.
    clusters: list[list[str]] = [[column] for column in source_columns]
    accepted.sort(key=lambda item: item[2], reverse=True)
    for left, right, score in accepted:
        left_cluster = next(c for c in clusters if left in c)
        right_cluster = next(c for c in clusters if right in c)
        if left_cluster is right_cluster:
            continue
        proposed = left_cluster + right_cluster
        valid = True
        for i in range(len(proposed)):
            for j in range(i + 1, len(proposed)):
                a, b = proposed[i], proposed[j]
                if _headers_cooccur(a, b, schemas) or score_between(a, b) < SEMANTIC_MATCH_THRESHOLD:
                    valid = False
                    break
            if not valid:
                break
        if valid:
            left_cluster.extend(right_cluster)
            clusters.remove(right_cluster)

    mapping: dict[str, str] = {}
    merged_count = 0
    for members in clusters:
        # Prefer the most informative source-derived header; never invent a
        # domain-specific label. Stable tie-break makes runs reproducible.
        canonical = sorted(
            members,
            key=lambda name: (len(re.sub(r"[_\W]+", "", str(name))), str(name).casefold()),
            reverse=True,
        )[0]
        if len(members) > 1:
            merged_count += 1
            for i, member_a in enumerate(members):
                for member_b in members[i + 1:]:
                    audit.append({
                        "Source_Column_A": member_a,
                        "Source_Column_B": member_b,
                        "Similarity": round(score_between(member_a, member_b), 4),
                        "Decision": "MERGED",
                        "Reason": "High-confidence semantic match; compatible profile; no same-schema collision.",
                        "Canonical_Column": canonical,
                    })
        mapping.update({member: canonical for member in members})
        if len(members) == 1:
            audit.append({
                "Source_Column_A": members[0],
                "Source_Column_B": "",
                "Similarity": None,
                "Decision": "KEPT SEPARATE",
                "Reason": "No sufficiently confident, unambiguous equivalent field was established.",
                "Canonical_Column": members[0],
            })

    if similarity_matrix is not None:
        warnings.append(
            f"Semantic model: {SEMANTIC_MODEL_NAME}; high-confidence merged groups: {merged_count}; "
            "ambiguous or incompatible fields were kept separate."
        )
    return mapping, warnings, audit


def apply_semantic_column_map(
    rows: list[dict],
    schemas: list[set],
    mapping: dict[str, str],
) -> tuple[list[dict], list[set]]:
    """Apply only validated mappings; collision guard is a final safety net."""
    if len(rows) != len(schemas):
        raise ValueError(
            "Semantic mapping stopped: row count and schema count differ. "
            "No records will be silently skipped."
        )

    remapped_rows: list[dict] = []
    remapped_schemas: list[set] = []
    for row, schema in zip(rows, schemas):
        new_row: dict[str, Any] = {}
        new_schema: set[str] = set()
        for original_column in sorted(schema, key=lambda item: str(item).casefold()):
            canonical = mapping.get(original_column, original_column)
            new_schema.add(canonical)
            if original_column not in row:
                continue
            value = row[original_column]
            if canonical not in new_row or is_missing(new_row[canonical]):
                new_row[canonical] = value
            elif not is_missing(value) and new_row[canonical] != value:
                # Never silently overwrite a value. This should only occur
                # for a source with duplicated/normalised headers.
                alternate = f"{canonical}__unmerged_source_field__{original_column}"
                new_row[alternate] = value
                new_schema.add(alternate)
        for metadata in METADATA_COLUMNS:
            if metadata in row:
                new_row[metadata] = row[metadata]
        remapped_rows.append(new_row)
        remapped_schemas.append(new_schema)
    return remapped_rows, remapped_schemas




# ============================================================
# DATASET GRANULARITY CLASSIFICATION
# ============================================================

def classify_record_granularity(source_entity_name: str) -> str:
    """
    Flag likely summary/aggregate sources using their source name.

    This is a review hint only. It does NOT prove that a record is
    aggregated, nor does it remove or alter any source record.
    """
    name = str(source_entity_name).casefold()
    tokens = set(re.findall(r"[a-z0-9]+", name))

    summary_terms = {
        "summary", "summaries", "aggregate", "aggregated",
        "aggregation", "rollup", "rollups"
    }

    if tokens.intersection(summary_terms):
        return "summary_candidate"

    return "unspecified"


# ============================================================
# CONSERVATIVE DATE DETECTION AND PARSING
# ============================================================

def is_date_column(column_name: str) -> bool:
    """
    Only treat clearly date/time-named columns as date columns.
    Ordinary fields such as 'period', 'age', or 'created_at'
    are deliberately not inferred as dates.
    """
    name = str(column_name).casefold()
    tokens = set(re.findall(r"[a-z0-9]+", name))

    if tokens.intersection({"date", "datetime", "timestamp"}):
        return True

    compact = re.sub(r"[^a-z0-9]+", "", name)
    return (
        compact.endswith("date")
        or compact.endswith("datetime")
        or compact.endswith("timestamp")
    )


def parse_unambiguous_date(value):
    """
    Convert only clearly interpretable date values.

    Supported:
      - Native Python/pandas date values
      - ISO dates and ISO datetimes
      - Day-first numeric dates only when the day/month order
        is unambiguous (one component is greater than 12)

    Ambiguous values such as 04/05/2024 remain unchanged.
    """
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()

    if isinstance(value, datetime):
        return value

    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())

    if not isinstance(value, str):
        return value

    text_value = value.strip()
    if not text_value:
        return value

    # ISO date or ISO datetime only.
    if re.match(r"^\d{4}-\d{2}-\d{2}(?:$|[T ])", text_value):
        try:
            normalized = text_value.replace("Z", "+00:00")
            parsed = datetime.fromisoformat(normalized)

            # Excel does not support timezone-aware datetime values.
            if parsed.tzinfo is not None:
                parsed = parsed.replace(tzinfo=None)

            return parsed
        except ValueError:
            return value

    # Numeric day-first dates are accepted only when unambiguous.
    match = re.fullmatch(
        r"(\d{1,2})([/.-])(\d{1,2})\2(\d{4})",
        text_value
    )

    if match:
        first = int(match.group(1))
        second = int(match.group(3))
        year = int(match.group(4))

        if first > 12 and 1 <= second <= 12:
            day, month = first, second
        elif second > 12 and 1 <= first <= 12:
            month, day = first, second
        else:
            return value

        try:
            return datetime(year, month, day)
        except ValueError:
            return value

    return value


# ============================================================
# 12. RECORD COLLECTION AND SCHEMA DISCOVERY
# ============================================================

def collect_all_records(
    source_files: list
):

    """
    Central consolidation engine.

    Every source record becomes one output row.

    The union of columns is discovered dynamically.

    Source schema is tracked separately for each source entity,
    allowing accurate missing-field markers.
    """

    all_rows = []

    all_schemas = []

    union_columns = OrderedDict()

    errors = []
    schema_audit = []

    for file_path in source_files:

        print(
            f"Processing: {get_relative_path(file_path)}"
        )

        try:

            source_records = read_source_file(file_path)

            source_format = (
                file_path.suffix.lstrip(".").upper()
            )

            # Group records by their source entity only for schema
            # discovery. They are NOT written to separate worksheets.

            grouped_records = {}

            for entity_name, record in source_records:

                grouped_records.setdefault(
                    entity_name,
                    []
                ).append(record)

            for entity_name, records in grouped_records.items():

                flattened_records = []

                entity_schema = set()

                for record in records:

                    flat = flatten_record(record)

                    flattened_records.append(flat)

                    entity_schema.update(
                        flat.keys()
                    )

                # Add every discovered source column to the global union.

                for column in entity_schema:

                    union_columns.setdefault(
                        column,
                        None
                    )

                # Append each source record to the master dataset.

                for record_number, flat in enumerate(
                    flattened_records,
                    start=1
                ):

                    row = {}

                    # Provenance fields.

                    row["__meta_company_source"] = (
                        get_company_source(file_path)
                    )

                    row["__meta_source_file"] = (
                        file_path.name
                    )

                    row["__meta_source_path"] = (
                        get_relative_path(file_path)
                    )

                    row["__meta_source_format"] = (
                        source_format
                    )

                    row["__meta_source_worksheet_or_collection"] = (
                        str(entity_name)
                    )

                    row["__meta_source_record_number"] = (
                        record_number
                    )

                    row["__meta_record_granularity"] = (
                        classify_record_granularity(entity_name)
                    )

                    row.update(flat)

                    all_rows.append(row)

                    all_schemas.append(
                        set(entity_schema)
                    )

        except Exception as error:

            errors.append(
                f"{get_relative_path(file_path)}: "
                f"{type(error).__name__}: {error}"
            )
    # Normalize date values before semantic schema mapping.
    # Date conventions are inferred independently for each source field.

    date_warnings = normalize_source_dates(
        all_rows,
        all_schemas,
    )

    for warning in date_warnings:
        print(f"DATE WARNING: {warning}")

    # Discover semantically equivalent columns across every source schema.
    # This replaces simple union-by-header with data-driven canonicalisation.
    if all_rows:
        semantic_map, semantic_warnings, schema_audit = build_semantic_column_map(
            all_rows,
            all_schemas,
        )

        all_rows, all_schemas = apply_semantic_column_map(
            all_rows,
            all_schemas,
            semantic_map,
        )

        for warning in semantic_warnings:
            print(f"\nSchema harmonization: {warning}")

    # Build the final union from canonicalized schemas.
    final_columns = list(METADATA_COLUMNS)
    canonical_union = OrderedDict()

    for schema in all_schemas:
        for column in schema:
            canonical_union.setdefault(column, None)

    for column in canonical_union:
        if column not in final_columns:
            final_columns.append(column)

    return (
        all_rows,
        all_schemas,
        final_columns,
        errors,
        schema_audit
    )


# ============================================================
# 13. MISSING VALUE HANDLING
# ============================================================

def build_output_matrix(
    rows: list,
    schemas: list,
    columns: list
) -> list:

    """
    Creates a rectangular matrix.

    If a field does not exist in the source entity's schema:
        FIELD_NOT_PRESENT

    If the field exists in that schema but the individual
    record has no value:
        VALUE_NOT_PROVIDED
    """

    matrix = []

    for row, source_schema in zip(rows, schemas):

        output_row = []

        for column in columns:

            value = row.get(column)

            if column in METADATA_COLUMNS:

                if is_missing(value):

                    value = VALUE_NOT_PROVIDED

                output_row.append(value)

                continue

            if column not in source_schema:

                output_row.append(
                    FIELD_NOT_PRESENT
                )

                continue

            if is_missing(value):

                output_row.append(
                    VALUE_NOT_PROVIDED
                )

            else:

                output_row.append(value)

        matrix.append(output_row)

    return matrix


# ============================================================
# 14. EXCEL VALUE VALIDATION
# ============================================================


import re
from decimal import Decimal, InvalidOperation


def prepare_excel_value(
    value: Any,
    column_name: str
) -> Any:

    """
    Preserve the correct data type of each individual Excel cell.

    - Missing-value descriptions remain text.
    - Integer strings become integers.
    - Decimal strings become numeric decimals (float).
    - Existing numeric values remain numeric.
    - Identifier-like fields are protected.
    - Leading-zero identifiers remain text.
    """

    if is_missing(value):
        return VALUE_NOT_PROVIDED

    # Never convert descriptive missing-value markers.
    if isinstance(value, str):

        if value in {
            FIELD_NOT_PRESENT,
            VALUE_NOT_PROVIDED
        }:
            return value

    # Protect identifiers and fields where numeric conversion
    # could destroy meaningful formatting or leading zeros.
    protected_keywords = {
        "id",
        "code",
        "phone",
        "mobile",
        "telephone",
        "postal",
        "postcode",
        "zipcode",
        "zip",
        "pin",
        "account",
        "registration",
        "reference"
    }

    column_tokens = set(
        re.findall(
            r"[a-z0-9]+",
            str(column_name).lower()
        )
    )

    if column_tokens.intersection(protected_keywords):
        return value

    # Date conversion is restricted to clearly date/time-named fields.
    # Ambiguous date strings remain text; no locale guessing is performed.
    if is_date_column(column_name):
        return parse_unambiguous_date(value)

    # Convert numeric strings without changing ordinary text.
    if isinstance(value, str):

        text_value = value.strip()

        if not text_value:
            return value

        # Preserve strings with leading zeros.
        if re.fullmatch(r"[+-]?0\d+", text_value):
            return value

        # Support standard integers and decimal values.
        if re.fullmatch(
            r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)",
            text_value
        ):

            try:

                number = Decimal(text_value)

                if number == number.to_integral_value():
                    return int(number)

                return float(number)

            except (InvalidOperation, ValueError, OverflowError):
                return value

        return value

    # Preserve native integer, float, and boolean values.
    if isinstance(value, bool):
        return value

    if isinstance(value, int):
        return value

    if isinstance(value, float):

        if not math.isfinite(value):
            return VALUE_NOT_PROVIDED

        return value

    if isinstance(value, (datetime, date, time)):
        return value

    if isinstance(value, (dict, list, tuple, set)):

        return json.dumps(
            convert_value(value),
            ensure_ascii=False,
            default=str
        )

    return str(value)



# ============================================================
# 15. EXCEL DIMENSION VALIDATION
# ============================================================

def validate_dimensions(
    record_count: int,
    column_count: int
):

    if record_count + 1 > MAX_EXCEL_ROWS:

        raise ValueError(
            "The number of records exceeds Excel's "
            "worksheet row limit. No data was discarded."
        )

    if column_count > MAX_EXCEL_COLUMNS:

        raise ValueError(
            "The number of columns exceeds Excel's "
            "worksheet column limit. No columns were discarded."
        )


# ============================================================
# 16. WORKBOOK GENERATION
# ============================================================

def write_workbook(
    columns: list,
    matrix: list,
    temporary_path: Path
):
    """
    Generate exactly one worksheet.

    No individual company or entity worksheets are created.
    """

    validate_dimensions(
        len(matrix),
        len(columns)
    )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = OUTPUT_SHEET

    # Write the header.
    worksheet.append(columns)

    # Append all consolidated records.
    for row in matrix:
        worksheet.append(
            [
                prepare_excel_value(
                    value,
                    columns[index]
                )
                for index, value in enumerate(row)
            ]
        )

    # Freeze only the header row.
    worksheet.freeze_panes = "A2"

    # Apply native AutoFilter to the complete dataset.
    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions

    worksheet.sheet_view.showGridLines = False

    # Apply body formatting.
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = BODY_ALIGNMENT

    # Apply a consistent display format to actual date/datetime cells.
    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, (datetime, date)):
                cell.number_format = DATE_DISPLAY_FORMAT

    # Calculate dynamic column widths.
    sample_end = min(
        worksheet.max_row,
        201
    )

    for column_index, column_name in enumerate(
        columns,
        start=1
    ):
        letter = get_column_letter(column_index)

        width = min(
            max(len(str(column_name)), 14),
            36
        )

        for row_index in range(2, sample_end + 1):
            value = worksheet.cell(
                row=row_index,
                column=column_index
            ).value

            if value is not None:
                width = max(
                    width,
                    min(len(str(value)), 36)
                )

        worksheet.column_dimensions[letter].width = min(
            max(width + 2, 14),
            40
        )

    # Save the temporary workbook.
    workbook.save(temporary_path)
    workbook.close()


# ============================================================
# 17. GENERATED WORKBOOK VALIDATION
# ============================================================

def validate_workbook(
    file_path: Path,
    expected_records: int,
    expected_columns: int
):

    workbook = load_workbook(
        file_path,
        read_only=False,
        data_only=False
    )

    try:

        # Exactly one worksheet must exist.

        if workbook.sheetnames != [OUTPUT_SHEET]:

            raise RuntimeError(
                "Output workbook must contain exactly "
                "one worksheet."
            )

        worksheet = workbook[OUTPUT_SHEET]

        if worksheet.freeze_panes != "A2":
            raise RuntimeError(
                "Freeze-pane validation failed: expected A2 so only the "
                "header row remains frozen."
            )

        if worksheet.auto_filter.ref != worksheet.dimensions:
            raise RuntimeError(
                "AutoFilter validation failed: filter range does not cover "
                "the complete output dataset."
            )

        expected_rows = expected_records + 1

        if worksheet.max_row != expected_rows:

            raise RuntimeError(
                f"Row validation failed. "
                f"Expected {expected_rows}, "
                f"found {worksheet.max_row}."
            )

        if worksheet.max_column != expected_columns:

            raise RuntimeError(
                "Column validation failed."
            )

        headers = [
            cell.value
            for cell in next(
                worksheet.iter_rows(
                    min_row=1,
                    max_row=1
                )
            )
        ]

        if len(headers) != expected_columns or any(header is None or str(header).strip() == "" for header in headers):
            raise RuntimeError("Header validation failed: missing or blank column name.")

        if len(set(headers)) != len(headers):
            raise RuntimeError("Header validation failed: duplicate column names detected.")

        required_metadata = set(METADATA_COLUMNS)
        if not required_metadata.issubset(set(headers)):
            raise RuntimeError("Metadata validation failed: one or more required metadata columns are missing.")

        for row_number, row in enumerate(
            worksheet.iter_rows(min_row=2),
            start=2
        ):
            if len(row) != expected_columns:
                raise RuntimeError(
                    f"Column count validation failed in row {row_number}."
                )

            for cell in row:
                if cell.value is None or cell.value == "":
                    raise RuntimeError(
                        f"Blank or incomplete cell detected at "
                        f"{cell.coordinate}. The workbook will not be published."
                    )

                if isinstance(cell.value, (datetime, date)):
                    if cell.number_format != DATE_DISPLAY_FORMAT:
                        raise RuntimeError(
                            f"Date-format validation failed at "
                            f"{cell.coordinate}. Expected "
                            f"{DATE_DISPLAY_FORMAT}, found {cell.number_format}."
                        )

    finally:

        workbook.close()


# ============================================================
# 18. SAFE OUTPUT PUBLICATION
# ============================================================

def publish_output(
    columns: list,
    matrix: list
):

    """
    Save to a temporary file in the output folder.

    Validate the temporary workbook.

    Replace the previous output only after successful validation.

    If generation or replacement fails, the previous valid
    workbook remains untouched.
    """

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    temporary_handle = tempfile.NamedTemporaryFile(
        prefix="consolidation_temp_",
        suffix=".xlsx",
        dir=OUTPUT_DIR,
        delete=False
    )

    temporary_path = Path(
        temporary_handle.name
    )

    temporary_handle.close()

    try:

        write_workbook(
            columns,
            matrix,
            temporary_path
        )

        validate_workbook(
            temporary_path,
            len(matrix),
            len(columns)
        )

        os.replace(
            temporary_path,
            OUTPUT_FILE
        )

    finally:

        if temporary_path.exists():

            try:

                temporary_path.unlink()

            except OSError:

                pass


# ============================================================
# 19. MAIN EXECUTION
# ============================================================

def main():

    print("=" * 65)

    print(
        "DATASET-INDEPENDENT COMPANY DATA CONSOLIDATION"
    )

    print("=" * 65)

    print(
        f"Input folder: {INPUT_DIR}"
    )

    print(
        f"Output file: {OUTPUT_FILE}"
    )

    print(
        f"Output worksheet: {OUTPUT_SHEET}"
    )

    print()

    # Discover files.

    source_files = discover_input_files()

    if not source_files:

        print(
            "No supported input files were found."
        )

        print(
            "Add JSON, CSV, XLSX, XLS or Parquet files "
            "inside the input folder."
        )

        return

    print(
        f"Discovered {len(source_files)} input files."
    )

    # Read and consolidate.

    print(
        "\nStarting dataset consolidation...\n"
    )

    (
        rows,
        schemas,
        columns,
        errors,
        schema_audit
    ) = collect_all_records(
        source_files
    )

    # Display processing errors.

    if errors:

        print(
            "\nWARNING: Some files could not be processed:"
        )

        for error in errors:

            print(
                f"  - {error}"
            )

    if errors:
        print(
            "\nConsolidation stopped because at least one input file could not be fully processed."
        )
        print("The output workbook was not generated or replaced; no partial result was published.")
        return

    if not rows:

        print(
            "\nNo records were extracted. "
            "The previous output was not replaced."
        )

        return

    # Build the master matrix.

    matrix = build_output_matrix(
        rows,
        schemas,
        columns
    )

    # Summary.

    print("\n" + "-" * 65)

    print(
        f"Input files discovered : {len(source_files):,}"
    )

    print(
        f"Records consolidated  : {len(rows):,}"
    )

    print(
        f"Data columns discovered: "
        f"{len(columns) - len(METADATA_COLUMNS):,}"
    )

    print(
        f"Metadata columns      : {len(METADATA_COLUMNS):,}"
    )

    print(
        f"Total output columns  : {len(columns):,}"
    )

    print("-" * 65)

    # Generate output.

    try:

        publish_output(
            columns,
            matrix
        )

    except PermissionError:

        print(
            "\nERROR: The output workbook is locked."
        )

        print(
            "Close Consolidated_Company_Dataset.xlsx "
            "in Excel and run the script again."
        )

        print(
            "The previous valid output has not been replaced."
        )

        return

    except Exception as error:

        print(
            f"\nOUTPUT ERROR: {type(error).__name__}: {error}"
        )

        print(
            "The previous valid output has not been replaced."
        )

        return

    # Schema matching decisions are summarized in the console.
    # No separate CSV audit file is generated.

    # Success.

    print("\n" + "=" * 65)

    print(
        "CONSOLIDATION COMPLETED SUCCESSFULLY"
    )

    print("=" * 65)

    print(
        f"Workbook: {OUTPUT_FILE}"
    )

    print(
        f"Worksheet: {OUTPUT_SHEET}"
    )

    print(
        f"Total consolidated records: {len(rows):,}"
    )

    print(
        f"Total columns: {len(columns):,}"
    )

    if errors:

        print(
            f"\nCompleted with {len(errors)} input file-processing errors."
        )

    else:

        print(
            "\nAll discovered input files were processed successfully."
        )

    print("\nImportant:")

    print(
        "- Exactly one output worksheet was generated."
    )

    print(
        "- All source records were appended into one dataset."
    )

    print(
        "- Columns were discovered dynamically."
    )

    print(
        "- Nested arrays were preserved as JSON text."
    )

    print(
        "- Original input files were not modified."
    )


# ============================================================
# 20. PROGRAM ENTRY POINT
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        print(
            "\nOperation cancelled by user."
        )

    except Exception as error:

        print(
            f"\nUNEXPECTED ERROR: {type(error).__name__}: {error}"
        )

        raise
