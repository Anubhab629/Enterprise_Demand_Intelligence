"""Power BI PBIP/PBIR/TMDL integration for the Enterprise Demand Intelligence Platform.

Production report generator.

Design contract
---------------
* The Power BI layer is the final BI/decision-support presentation layer.
* Source business schema is never hardcoded.
* Pipeline-produced analytical evidence contracts are discovered at runtime.
* A single immutable TableDefinitionRegistry is used for TMDL emission, measures,
  report bindings and validation.
* Current-run workbooks are staged under powerbi/data/<RUN_ID>.
* The PBIP project is generated under
  powerbi/Demand_Intelligence_Report_<RUN_ID>/.
* The proven Fix 06 Excel -> Power Query M -> TMDL import path is preserved.
* Relationships are not materialized unless a complete structural safety test
  succeeds.
* Organizational-boundary discovery never becomes operational RLS by itself.
* Report pages are capability-driven but follow the locked 8-page architecture:
  Executive Intelligence, Data Landscape, Demand Intelligence, Forecast Outlook,
  Recommendations & Actions, Signal Intelligence (when supported),
  Data Quality & Readiness, Record Explorer.
* Analytical pages use KPI -> insight -> supporting insight -> evidence/watchlist.
  Record Explorer is the only broad row-level surface.

The implementation deliberately uses only PBIR/TMDL structures that are supported
by the current working scaffold and validates the generated project before it is
returned to main.py. Actual Power BI Desktop open/refresh/save remains an external
acceptance gate.
"""

from __future__ import annotations

import datetime as dt
import ast
import json
import math
import os
import re
import shutil
import subprocess
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable

try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None

import config
import env


# ---------------------------------------------------------------------------
# PBIR/TMDL format contracts
# ---------------------------------------------------------------------------

POWER_BI_SCHEMA_ROOT = (
    "https://developer.microsoft.com/json-schemas/fabric/item/report"
)
PBIP_SCHEMA = (
    "https://developer.microsoft.com/json-schemas/fabric/pbip/"
    "pbipProperties/1.0.0/schema.json"
)
PBISM_SCHEMA = (
    "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/"
    "definitionProperties/1.0.0/schema.json"
)
PBIR_SCHEMA = (
    "https://developer.microsoft.com/json-schemas/fabric/item/report/"
    "definitionProperties/2.0.0/schema.json"
)
VISUAL_SCHEMA = (
    f"{POWER_BI_SCHEMA_ROOT}/definition/visualContainer/2.9.0/schema.json"
)
PAGE_SCHEMA = (
    f"{POWER_BI_SCHEMA_ROOT}/definition/page/2.1.0/schema.json"
)
REPORT_SCHEMA = (
    f"{POWER_BI_SCHEMA_ROOT}/definition/report/3.3.0/schema.json"
)
PBIR_VERSION_SCHEMA = (
    f"{POWER_BI_SCHEMA_ROOT}/definition/versionMetadata/1.0.0/schema.json"
)
PAGES_METADATA_SCHEMA = (
    f"{POWER_BI_SCHEMA_ROOT}/definition/pagesMetadata/1.1.0/schema.json"
)

PAGE_WIDTH = 1280
PAGE_HEIGHT = 720
MAX_SAMPLE_ROWS_FOR_SCHEMA = 200
MAX_REPORT_VISUALS_PER_PAGE = 12
MAX_RELATIONSHIP_CANDIDATES = max(
    1, int(getattr(config, "MAX_RELATIONSHIP_PAIRS", 5000))
)
MAX_POWERBI_TABLES = max(
    1, int(getattr(config, "MAX_DISCOVERED_TABLES", 1000))
)
MAX_POWERBI_COLUMNS_PER_TABLE = max(
    1, int(getattr(config, "MAX_DISCOVERED_COLUMNS_PER_TABLE", 1000))
)

# Semantic meaning is carried only by explicit worker-output contracts.\n

# ---------------------------------------------------------------------------
# Small primitives
# ---------------------------------------------------------------------------

def utc_now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sanitize_identifier(value: Any, fallback: str = "Table") -> str:
    text = re.sub(r"[^A-Za-z0-9_]+", "_", str(value).strip())
    text = re.sub(r"_+", "_", text).strip("_") or fallback
    if text[0].isdigit():
        text = f"T_{text}"
    return text[:120]


def unique_identifier(base: str, used: set[str]) -> str:
    candidate = base
    counter = 2
    while candidate in used:
        candidate = f"{base}_{counter}"
        counter += 1
    used.add(candidate)
    return candidate


def tmdl_name(value: str) -> str:
    return "'" + str(value).replace("'", "''") + "'"

def dax_column_ref(table_name: str, column_name: str) -> str:
    """Build a fully-qualified DAX column reference."""
    safe_column = str(column_name).replace("]", "]]")
    return f"{tmdl_name(table_name)}[{safe_column}]"


def tmdl_string(value: Any) -> str:
    return '"' + str(value).replace('"', '""') + '"'


def json_write(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )


def nonempty(value: Any) -> bool:
    return value is not None and str(value).strip() != ""


def normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def infer_powerbi_type(values: Iterable[Any]) -> str:
    values = [v for v in values if nonempty(v)]
    if not values:
        return "string"
    if all(isinstance(v, bool) for v in values):
        return "boolean"
    if all(isinstance(v, int) and not isinstance(v, bool) for v in values):
        return "int64"
    if all(
        isinstance(v, (int, float)) and not isinstance(v, bool)
        for v in values
    ):
        return "double"
    if all(isinstance(v, (dt.datetime, dt.date)) for v in values):
        return "dateTime"
    return "string"


# ---------------------------------------------------------------------------
# Workbook/schema discovery
# ---------------------------------------------------------------------------

def load_workbook_schema(path: Path) -> list[dict[str, Any]]:
    if openpyxl is None:
        raise RuntimeError("openpyxl is required for Power BI project generation.")
    if not path.is_file() or path.stat().st_size == 0:
        raise FileNotFoundError(f"Power BI workbook missing or empty: {path}")

    workbook = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=True,
    )
    result: list[dict[str, Any]] = []
    try:
        if len(workbook.sheetnames) > MAX_POWERBI_TABLES:
            raise RuntimeError(
                f"{path.name} contains {len(workbook.sheetnames)} sheets; "
                f"technical limit is {MAX_POWERBI_TABLES}."
            )

        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            row_iter = worksheet.iter_rows(
                min_row=1,
                max_row=1 + MAX_SAMPLE_ROWS_FOR_SCHEMA,
                values_only=True,
            )
            try:
                header = next(row_iter)
            except StopIteration:
                continue

            headers: list[str] = []
            seen_headers: set[str] = set()
            for index, raw in enumerate(header):
                base = (
                    str(raw).strip()
                    if raw is not None and str(raw).strip()
                    else f"Column_{index + 1}"
                )
                header_name = unique_identifier(base, seen_headers)
                headers.append(header_name)

            if len(headers) > MAX_POWERBI_COLUMNS_PER_TABLE:
                raise RuntimeError(
                    f"{path.name}/{sheet_name} contains {len(headers)} columns; "
                    f"technical limit is {MAX_POWERBI_COLUMNS_PER_TABLE}."
                )

            samples = list(row_iter)
            columns: list[dict[str, Any]] = []
            for index, source_name in enumerate(headers):
                values = [
                    row[index] if index < len(row) else None
                    for row in samples
                ]
                clean_values = [
                    v for v in values if nonempty(v)
                ]
                distinct = {str(v) for v in clean_values}
                columns.append(
                    {
                        "source_name": source_name,
                        "powerbi_type": infer_powerbi_type(values),
                        "sample_values": values,
                        "sample_non_null": len(clean_values),
                        "sample_distinct": len(distinct),
                        "sample_cardinality": (
                            len(distinct) / max(1, len(clean_values))
                        ),
                    }
                )

            # worksheet.max_row can include formatted empty rows. It is still
            # preferable to inventing a count from the sample. The source
            # workbook is itself a pipeline artifact and row_count is metadata,
            # not a business assumption.
            row_count = max(0, int(worksheet.max_row or 1) - 1)
            result.append(
                {
                    "sheet_name": str(sheet_name),
                    "columns": columns,
                    "row_count": row_count,
                }
            )
    finally:
        workbook.close()

    return result


def excel_m_expression(
    source_file: Path,
    sheet_name: str,
    columns: list[dict[str, Any]],
) -> str:
    path_text = source_file.resolve().as_posix().replace('"', '""')
    item_text = sheet_name.replace('"', '""')

    type_map = {
        "string": "type text",
        "int64": "Int64.Type",
        "double": "type number",
        "boolean": "type logical",
        "dateTime": "type datetime",
    }

    transforms: list[str] = []
    for column in columns:
        source_name = str(column["source_name"]).replace('"', '""')
        powerbi_type = str(column["powerbi_type"])
        m_type = type_map.get(powerbi_type)
        if m_type is None:
            raise RuntimeError(
                f"Unsupported Power Query type mapping: {powerbi_type!r}"
            )
        transforms.append(f'{{"{source_name}", {m_type}}}')

    lines = [
        "let",
        f'    Source = Excel.Workbook(File.Contents("{path_text}"), null, true),',
        f'    SheetData = Source{{[Item="{item_text}",Kind="Sheet"]}}[Data],',
        "    PromotedHeaders = Table.PromoteHeaders(",
        "        SheetData, [PromoteAllScalars=true]),",
        "    TypedColumns = Table.TransformColumnTypes(",
        "        PromotedHeaders,",
        "        {",
    ]

    lines.extend(
        f"            {item}{',' if index < len(transforms) - 1 else ''}"
        for index, item in enumerate(transforms)
    )

    lines.extend(
        [
            "        }",
            "    )",
            "in",
            "    TypedColumns",
        ]
    )

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Canonical model registry
# ---------------------------------------------------------------------------

@dataclass
class TableDefinition:
    table_name: str
    source_file: Path
    sheet_name: str
    columns: list[dict[str, Any]]
    row_count: int
    tmdl_filename: str

    @property
    def artifact_group(self) -> str:
        return self.source_file.stem.lower()


@dataclass
class EvidenceProfile:
    table: TableDefinition
    temporal: list[dict[str, Any]] = field(default_factory=list)
    numeric: list[dict[str, Any]] = field(default_factory=list)
    categorical: list[dict[str, Any]] = field(default_factory=list)
    text: list[dict[str, Any]] = field(default_factory=list)
    role_candidates: dict[str, list[dict[str, Any]]] = field(
        default_factory=dict
    )


class TableDefinitionRegistry:
    """Single source of truth for every generated table and report binding."""

    def __init__(self, tables: list[TableDefinition]) -> None:
        self.tables = tables
        self.by_name = {t.table_name: t for t in tables}
        self.by_filename = {t.tmdl_filename: t for t in tables}

    @classmethod
    def discover(cls, staged_files: list[Path]) -> "TableDefinitionRegistry":
        tables: list[TableDefinition] = []
        used_tables: set[str] = set()

        for source in staged_files:
            for sheet in load_workbook_schema(source):
                base = sanitize_identifier(
                    f"{source.stem}_{sheet['sheet_name']}",
                    "Table",
                )
                table_name = unique_identifier(base, used_tables)

                used_columns: set[str] = set()
                columns: list[dict[str, Any]] = []
                for column in sheet["columns"]:
                    safe_name = unique_identifier(
                        sanitize_identifier(
                            column["source_name"],
                            "Column",
                        ),
                        used_columns,
                    )
                    columns.append(
                        {
                            **column,
                            "safe_name": safe_name,
                        }
                    )

                tmdl_filename = (
                    sanitize_identifier(table_name, "Table") + ".tmdl"
                )
                tables.append(
                    TableDefinition(
                        table_name=table_name,
                        source_file=source,
                        sheet_name=sheet["sheet_name"],
                        columns=columns,
                        row_count=int(sheet["row_count"]),
                        tmdl_filename=tmdl_filename,
                    )
                )

        if len(tables) > MAX_POWERBI_TABLES:
            raise RuntimeError(
                f"Power BI model would contain {len(tables)} tables; "
                f"technical limit is {MAX_POWERBI_TABLES}."
            )

        return cls(tables)

    def table(self, name: str) -> TableDefinition:
        try:
            return self.by_name[name]
        except KeyError as exc:
            raise RuntimeError(
                f"Report binding references unknown model table: {name}"
            ) from exc

    def column(self, table_name: str, safe_name: str) -> dict[str, Any]:
        table = self.table(table_name)
        for column in table.columns:
            if column["safe_name"] == safe_name:
                return column
        raise RuntimeError(
            f"Report binding references unknown field: "
            f"{table_name}.{safe_name}"
        )

    def tmdl_path(self, semantic_definition: Path, table_name: str) -> Path:
        table = self.table(table_name)
        path = semantic_definition / "tables" / table.tmdl_filename
        if not path.is_file():
            raise RuntimeError(
                f"Canonical table registry references missing TMDL: {path}"
            )
        return path

    def artifact_tables(self, artifact_stem: str) -> list[TableDefinition]:
        stem = artifact_stem.lower()
        return [t for t in self.tables if t.artifact_group == stem]


# ---------------------------------------------------------------------------
# Evidence discovery
# ---------------------------------------------------------------------------


def build_evidence_profiles(
    registry: TableDefinitionRegistry,
) -> list[EvidenceProfile]:
    """Build type/cardinality profiles for generic row exploration only.

    No business meaning is inferred from source column names. Analytical
    semantics come exclusively from explicit worker evidence contracts.
    """
    profiles: list[EvidenceProfile] = []
    for table in registry.tables:
        profile = EvidenceProfile(table=table)
        for column in table.columns:
            kind = column["powerbi_type"]
            if kind == "dateTime":
                profile.temporal.append(column)
            elif kind in {"int64", "double"}:
                profile.numeric.append(column)
            elif kind == "string":
                distinct = int(column.get("sample_distinct", 0))
                if 2 <= distinct <= 50:
                    profile.categorical.append(column)
                non_null = int(column.get("sample_non_null", 0))
                if non_null:
                    avg_len = sum(
                        len(str(v))
                        for v in column.get("sample_values", [])
                        if nonempty(v)
                    ) / max(1, non_null)
                    if avg_len >= 40:
                        profile.text.append(column)
        profiles.append(profile)
    return profiles


def profile_for(
    profiles: list[EvidenceProfile],
    table: TableDefinition,
) -> EvidenceProfile:
    for profile in profiles:
        if profile.table.table_name == table.table_name:
            return profile
    raise RuntimeError(f"Missing evidence profile for {table.table_name}")



def contract_table(
    registry: TableDefinitionRegistry,
    artifact_stem: str,
    sheet_name: str,
    required_columns: tuple[str, ...] = (),
) -> TableDefinition | None:
    """Resolve a worker-owned evidence table by its explicit output contract.

    These names belong to the Intelligence Core output contract, not to the
    uploaded/source dataset schema. Report semantics are never inferred from
    source column names or heuristic role tokens.
    """
    matches = [
        table
        for table in registry.artifact_tables(artifact_stem)
        if table.sheet_name == sheet_name
    ]
    if not matches:
        return None
    table = matches[0]
    if required_columns:
        available = {column["source_name"] for column in table.columns}
        if not set(required_columns).issubset(available):
            return None
    return table


def contract_column(
    table: TableDefinition,
    source_name: str,
) -> dict[str, Any]:
    for column in table.columns:
        if column["source_name"] == source_name:
            return column
    raise RuntimeError(
        f"Worker evidence contract field is missing: "
        f"{table.table_name}.{source_name}"
    )


def contract_columns(
    table: TableDefinition,
    source_names: tuple[str, ...],
) -> list[dict[str, Any]]:
    return [contract_column(table, name) for name in source_names]


def worker_contract_available(
    registry: TableDefinitionRegistry,
    artifact_stem: str,
    sheet_name: str,
    required_columns: tuple[str, ...],
) -> bool:
    return contract_table(
        registry, artifact_stem, sheet_name, required_columns
    ) is not None


def best_temporal(profile: EvidenceProfile) -> dict[str, Any] | None:
    return profile.temporal[0] if profile.temporal else None


def select_dominant_table(
    registry: TableDefinitionRegistry,
    artifact_stems: tuple[str, ...],
    sheet_tokens: tuple[str, ...] = (),
) -> TableDefinition | None:
    candidates: list[TableDefinition] = []
    for stem in artifact_stems:
        candidates.extend(registry.artifact_tables(stem))

    if not candidates:
        return None

    if sheet_tokens:
        token_candidates = [
            t
            for t in candidates
            if any(
                token in normalize_header(t.sheet_name)
                for token in sheet_tokens
            )
        ]
        if token_candidates:
            candidates = token_candidates

    candidates.sort(
        key=lambda t: (
            int(t.row_count),
            len(t.columns),
            t.sheet_name,
        ),
        reverse=True,
    )
    return candidates[0]


# ---------------------------------------------------------------------------
# TMDL generation
# ---------------------------------------------------------------------------

def build_tmdl_table(table: TableDefinition) -> str:
    lines = [f"table {tmdl_name(table.table_name)}"]

    for column in table.columns:
        kind = column["powerbi_type"]
        summarize = (
            "none"
            if kind in {"string", "boolean", "dateTime"}
            else "sum"
        )
        lines.extend(
            [
                f"    column {tmdl_name(column['safe_name'])}",
                f"        dataType: {kind}",
                f"        sourceColumn: {tmdl_string(column['source_name'])}",
                f"        summarizeBy: {summarize}",
                "",
            ]
        )

    partition_name = sanitize_identifier(
        f"{table.table_name}_Partition",
        "Partition",
    )
    lines.extend(
        [
            f"    partition {tmdl_name(partition_name)} = m",
            "        mode: import",
            "        source =",
        ]
    )
    lines.extend(
        "            " + line
        for line in excel_m_expression(
            table.source_file,
            table.sheet_name,
            table.columns,
        ).splitlines()
    )    

    return "\n".join(lines).rstrip() + "\n"


def write_tmdl_registry(
    semantic_definition: Path,
    registry: TableDefinitionRegistry,
) -> None:
    tables_dir = semantic_definition / "tables"
    tables_dir.mkdir(parents=True, exist_ok=True)

    # Remove stale table definitions inside this newly created project only.
    for path in tables_dir.glob("*.tmdl"):
        path.unlink()

    # Emit each canonical table definition directly to its final path and
    # immediately verify that the file is still present/readable. Power BI
    # workspaces may live under OneDrive on Windows; transient synchronization
    # can otherwise race with file creation. Retry only the affected file;
    # do not rebuild the project or alter the registry.
    for table in registry.tables:
        target = tables_dir / table.tmdl_filename
        payload = build_tmdl_table(table)
        last_error: Exception | None = None
        for attempt in range(3):
            try:
                target.write_text(payload, encoding="utf-8")
                if target.is_file() and target.stat().st_size > 0:
                    break
            except OSError as exc:
                last_error = exc

            if attempt < 2:
                time.sleep(0.15 * (attempt + 1))
        else:
            detail = f": {last_error}" if last_error else ""
            raise RuntimeError(
                f"TMDL emission failed for {table.table_name}: {target}{detail}"
            )

    # Canonical registry/on-disk consistency gate. This is intentionally before
    # any report binding or measure generation so a missing TMDL can never
    # surface later as an opaque field-registry/FileNotFoundError.
    expected = {t.tmdl_filename for t in registry.tables}
    actual = {p.name for p in tables_dir.glob("*.tmdl")}
    missing = sorted(expected - actual)
    unexpected = sorted(actual - expected)
    if missing or unexpected:
        raise RuntimeError(
            "TMDL registry emission contract failed. "
            f"Missing={missing}; Unexpected={unexpected}"
        )

    # A relationships.tmdl file is emitted only when the registry contains
    # validated materialized relationships. The current conservative model
    # normally has zero materialized relationships.
    relationship_file = semantic_definition / "relationships.tmdl"
    if relationship_file.exists():
        relationship_file.unlink()


# ---------------------------------------------------------------------------
# Relationship and security discovery
# ---------------------------------------------------------------------------

def discover_relationship_candidates(
    registry: TableDefinitionRegistry,
) -> list[dict[str, Any]]:
    """Discover candidates conservatively without automatically materializing them.

    A candidate is not a relationship. We document it only after testing:
    candidate key -> uniqueness -> coverage. Since the evidence workbooks are
    intentionally independent analytical outputs, direction/ambiguity cannot
    safely be inferred from shared names alone, so no relationship is
    materialized by this generator.
    """
    occurrences: dict[str, list[dict[str, Any]]] = {}

    for table in registry.tables:
        for column in table.columns:
            if column["powerbi_type"] == "dateTime":
                continue
            values = {
                str(v)
                for v in column.get("sample_values", [])
                if nonempty(v)
            }
            if len(values) < 2:
                continue
            occurrences.setdefault(
                normalize_header(column["source_name"]),
                [],
            ).append(
                {
                    "table": table,
                    "column": column,
                    "values": values,
                }
            )

    candidates: list[dict[str, Any]] = []

    for key, entries in occurrences.items():
        for i in range(len(entries)):
            for j in range(i + 1, len(entries)):
                left = entries[i]
                right = entries[j]
                left_values = left["values"]
                right_values = right["values"]
                overlap = len(left_values & right_values) / max(
                    1,
                    min(len(left_values), len(right_values)),
                )
                if overlap <= 0:
                    continue

                left_unique = len(left_values) == len(
                    [
                        v
                        for v in left["column"].get("sample_values", [])
                        if nonempty(v)
                    ]
                )
                right_unique = len(right_values) == len(
                    [
                        v
                        for v in right["column"].get("sample_values", [])
                        if nonempty(v)
                    ]
                )

                candidates.append(
                    {
                        "normalized_candidate_key": key,
                        "left_table": left["table"].table_name,
                        "right_table": right["table"].table_name,
                        "left_column": left["column"]["safe_name"],
                        "right_column": right["column"]["safe_name"],
                        "uniqueness": {
                            "left": left_unique,
                            "right": right_unique,
                        },
                        "sample_coverage": round(float(overlap), 6),
                        "cardinality": (
                            "one_to_one_candidate"
                            if left_unique and right_unique
                            else "many_side_candidate"
                        ),
                        "direction": "not_materialized",
                        "ambiguity": "not_resolved",
                        "materialized": False,
                        "reason": (
                            "analytical evidence tables are independent "
                            "unless a complete relational contract is validated"
                        ),
                    }
                )
                if len(candidates) >= MAX_RELATIONSHIP_CANDIDATES:
                    return candidates

    return candidates


def discover_access_boundary(
    registry: TableDefinitionRegistry,
) -> dict[str, Any] | None:
    getter = getattr(env, "get_max_access_boundary_cardinality", None)
    if callable(getter):
        limit = max(2, int(getter()))
    else:
        fallback = getattr(
            env,
            "get_max_companies_per_dataset",
            lambda: 100,
        )
        limit = max(2, int(fallback()))

    candidates: list[dict[str, Any]] = []

    for table in registry.tables:
        if table.row_count <= 1 or table.row_count > limit:
            continue
        for column in table.columns:
            if column["powerbi_type"] not in {"string", "int64"}:
                continue
            values = [
                v
                for v in column.get("sample_values", [])
                if nonempty(v)
            ]
            distinct = {str(v) for v in values}
            if len(distinct) < min(5, table.row_count):
                continue
            if len(distinct) != len(values):
                continue

            candidates.append(
                {
                    "table": table,
                    "column": column,
                    "values": values,
                }
            )

    candidates.sort(
        key=lambda item: (
            item["table"].row_count,
            item["table"].table_name,
            item["column"]["source_name"],
        )
    )

    for candidate in candidates:
        values = {str(v) for v in candidate["values"]}
        related_tables = 0

        for table in registry.tables:
            if table is candidate["table"]:
                continue
            for column in table.columns:
                if (
                    normalize_header(column["source_name"])
                    != normalize_header(
                        candidate["column"]["source_name"]
                    )
                ):
                    continue
                other = {
                    str(v)
                    for v in column.get("sample_values", [])
                    if nonempty(v)
                }
                if values & other:
                    related_tables += 1
                    break

        if related_tables >= 2:
            return {
                "table": candidate["table"].table_name,
                "column": candidate["column"]["safe_name"],
                "related_table_count": related_tables,
                "operational_identity_mapping": False,
            }

    return None


# ---------------------------------------------------------------------------
# Semantic model generation
# ---------------------------------------------------------------------------

def generate_model(
    project_root: Path,
    registry: TableDefinitionRegistry,
) -> dict[str, Any]:
    # LOCKED FIX 06 baseline child artifact names.
    # Keep these short: the parent already contains the unique Run ID.
    # Do not repeat the Run ID in child names on Windows/OneDrive.
    semantic_name = "Model.SemanticModel"
    report_name = "Report.Report"

    semantic_root = project_root / semantic_name
    semantic_definition = semantic_root / "definition"
    tables_directory = semantic_definition / "tables"
    report_root = project_root / report_name
    report_definition = report_root / "definition"

    tables_directory.mkdir(parents=True, exist_ok=True)
    (report_definition / "pages").mkdir(
        parents=True,
        exist_ok=True,
    )
    (report_root / "StaticResources").mkdir(
        parents=True,
        exist_ok=True,
    )

    # PBIP points to the report folder. The report definition then points
    # relatively to the sibling semantic model folder.
    json_write(
        project_root / f"{project_root.name}.pbip",
        {
            "$schema": PBIP_SCHEMA,
            "version": "1.0",
            "artifacts": [
                {
                    "report": {
                        "path": report_name,
                    }
                }
            ],
        },
    )

    json_write(
        semantic_root / "definition.pbism",
        {
            "$schema": PBISM_SCHEMA,
            "version": "4.2",
            "settings": {
                "qnaEnabled": True,
            },
        },
    )

    semantic_definition.mkdir(parents=True, exist_ok=True)

    (semantic_definition / "database.tmdl").write_text(
        f"database {sanitize_identifier(project_root.name, 'Model')}\n"
        "    compatibilityLevel: 1567\n",
        encoding="utf-8",
    )
    (semantic_definition / "model.tmdl").write_text(
        "model Model\n"
        "    culture: en-US\n"
        "    defaultPowerBIDataSourceVersion: powerBI_V3\n",
        encoding="utf-8",
    )

    write_tmdl_registry(semantic_definition, registry)

    relationship_candidates = discover_relationship_candidates(registry)
    boundary = discover_access_boundary(registry)

    if boundary:
        security = {
            "status": "BOUNDARY_DISCOVERED_MAPPING_REQUIRED",
            "table": boundary["table"],
            "column": boundary["column"],
            "related_table_count": boundary["related_table_count"],
            "operational_identity_mapping_established": False,
            "operational_rls": False,
            "access_default": "deny",
        }
    else:
        security = {
            "status": "FAIL_SAFE_BOUNDARY_NOT_ESTABLISHED",
            "operational_identity_mapping_established": False,
            "operational_rls": False,
            "access_default": "deny",
        }

    return {
        "semantic_name": semantic_name,
        "report_name": report_name,
        "semantic_definition": str(
            semantic_definition.resolve()
        ),
        "table_count": len(registry.tables),
        "relationship_count": 0,
        "relationship_candidates": len(
            relationship_candidates
        ),
        "relationship_policy": (
            "CANDIDATE_UNIQUENESS_COVERAGE_CARDINALITY_DIRECTION_"
            "AMBIGUITY_REQUIRED_BEFORE_MATERIALIZATION"
        ),
        "security": security,
        "registry": [
            {
                "table_name": t.table_name,
                "tmdl_filename": t.tmdl_filename,
                "source_file": str(t.source_file.resolve()),
                "sheet_name": t.sheet_name,
                "row_count": t.row_count,
                "columns": [
                    {
                        "source_name": c["source_name"],
                        "safe_name": c["safe_name"],
                        "powerbi_type": c["powerbi_type"],
                    }
                    for c in t.columns
                ],
            }
            for t in registry.tables
        ],
    }


# ---------------------------------------------------------------------------
# PBIR expressions and visuals
# ---------------------------------------------------------------------------

def column_expr(entity: str, prop: str) -> dict[str, Any]:
    return {
        "Column": {
            "Expression": {
                "SourceRef": {
                    "Entity": entity
                }
            },
            "Property": prop,
        }
    }


def measure_expr(entity: str, prop: str) -> dict[str, Any]:
    return {
        "Measure": {
            "Expression": {
                "SourceRef": {
                    "Entity": entity
                }
            },
            "Property": prop,
        }
    }


def aggregation_expr(
    entity: str,
    prop: str,
    function: int = 0,
) -> dict[str, Any]:
    return {
        "Aggregation": {
            "Expression": column_expr(entity, prop),
            "Function": function,
        }
    }


def projection(
    expression: dict[str, Any],
    entity: str,
    property_name: str,
    *,
    active: bool = True,
) -> dict[str, Any]:
    result = {
        "field": expression,
        "queryRef": f"{entity}.{property_name}",
        "nativeQueryRef": property_name,
    }
    if active:
        result["active"] = True
    return result


def literal_value(value: str) -> dict[str, Any]:
    return {
        "expr": {
            "Literal": {
                "Value": value,
            }
        }
    }


def title_objects(title: str) -> dict[str, Any]:
    return {
        "visualContainerObjects": {
            "title": [
                {
                    "properties": {
                        "show": literal_value("true"),
                        "text": literal_value(
                            "'" + title.replace("'", "''") + "'"
                        ),
                        "fontSize": literal_value("13D"),
                    }
                }
            ],
            "subTitle": [
                {
                    "properties": {
                        "show": literal_value("false")
                    }
                }
            ],
            "background": [
                {
                    "properties": {
                        "show": literal_value("true"),
                        "transparency": literal_value("0D"),
                    }
                }
            ],
            "border": [
                {
                    "properties": {
                        "show": literal_value("true"),
                        "radius": literal_value("6D"),
                        "width": literal_value("1D"),
                    }
                }
            ],
            "visualHeader": [
                {
                    "properties": {
                        "show": literal_value("false")
                    }
                }
            ],
        }
    }


def write_visual(
    path: Path,
    *,
    name: str,
    visual_type: str,
    position: dict[str, Any],
    query_state: dict[str, Any] | None = None,
    title: str | None = None,
    objects: dict[str, Any] | None = None,
    visual_container_objects: dict[str, Any] | None = None,
) -> None:
    payload: dict[str, Any] = {
        "$schema": VISUAL_SCHEMA,
        "name": name,
        "position": position,
        "visual": {
            "visualType": visual_type,
        },
    }

    visual = payload["visual"]

    if query_state is not None:
        visual["query"] = {
            "queryState": query_state
        }

    if objects:
        visual["objects"] = objects

    if title:
        vco = title_objects(title)["visualContainerObjects"]
        if visual_container_objects:
            for key, value in visual_container_objects.items():
                vco[key] = value
        visual["visualContainerObjects"] = vco
    elif visual_container_objects:
        visual["visualContainerObjects"] = visual_container_objects

    path.parent.mkdir(parents=True, exist_ok=True)
    json_write(path, payload)


def make_card(
    table: TableDefinition,
    measure_name: str,
    title: str,
    position: dict[str, Any],
) -> dict[str, Any]:
    return {
        "kind": "cardVisual",
        "position": position,
        "query_state": {
            "Data": {
                "projections": [
                    projection(
                        measure_expr(
                            table.table_name,
                            measure_name,
                        ),
                        table.table_name,
                        measure_name,
                    )
                ]
            }
        },
        "title": title,
    }


def make_chart(
    table: TableDefinition,
    category: dict[str, Any],
    numeric: dict[str, Any],
    title: str,
    position: dict[str, Any],
    visual_type: str = "barChart",
) -> dict[str, Any]:
    return {
        "kind": visual_type,
        "position": position,
        "query_state": {
            "Category": {
                "projections": [
                    projection(
                        column_expr(
                            table.table_name,
                            category["safe_name"],
                        ),
                        table.table_name,
                        category["safe_name"],
                    )
                ]
            },
            "Y": {
                "projections": [
                    projection(
                        aggregation_expr(
                            table.table_name,
                            numeric["safe_name"],
                            0,
                        ),
                        table.table_name,
                        numeric["safe_name"],
                        active=False,
                    )
                ]
            },
        },
        "title": title,
    }


def ensure_table_definition_file(
    table: TableDefinition,
    semantic_definition: Path,
) -> Path:
    """Ensure the canonical TMDL exists immediately before any mutation.

    This is a narrow recovery guard for Windows/OneDrive file churn. It never
    changes the registry or invents a different filename; the registry remains
    the sole authority.
    """
    path = semantic_definition / "tables" / table.tmdl_filename
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.is_file() and path.stat().st_size > 0:
        return path

    payload = build_tmdl_table(table)
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
            temporary.write_text(payload, encoding="utf-8")
            os.replace(temporary, path)
            if path.is_file() and path.stat().st_size > 0:
                return path
        except OSError as exc:
            last_error = exc
        finally:
            try:
                if 'temporary' in locals() and temporary.exists():
                    temporary.unlink()
            except OSError:
                pass
        if attempt < 2:
            time.sleep(0.15 * (attempt + 1))

    detail = f": {last_error}" if last_error else ""
    raise RuntimeError(f"Canonical TMDL could not be restored: {path}{detail}")


def ensure_table_count_measure(
    table: TableDefinition,
    semantic_definition: Path,
    measure_name: str,
) -> str:
    """Create a table-local COUNTROWS measure for a distribution visual."""
    path = ensure_table_definition_file(table, semantic_definition)
    append_measure(
        path,
        measure_name,
        f"COUNTROWS({tmdl_name(table.table_name)})",
    )
    return measure_name


def ensure_table_measure(
    table: TableDefinition,
    semantic_definition: Path,
    measure_name: str,
    expression: str,
) -> str:
    path = ensure_table_definition_file(table, semantic_definition)
    append_measure(path, measure_name, expression)
    return measure_name


def ensure_table_sum_measure(
    table: TableDefinition,
    semantic_definition: Path,
    measure_name: str,
    column: dict[str, Any],
) -> str:
    return ensure_table_measure(
        table,
        semantic_definition,
        measure_name,
        f"SUM({dax_column_ref(table.table_name, column['safe_name'])})",
    )


def ensure_table_average_measure(
    table: TableDefinition,
    semantic_definition: Path,
    measure_name: str,
    column: dict[str, Any],
    *,
    format_string: str = "0.00",
) -> str:
    path = ensure_table_definition_file(table, semantic_definition)
    append_measure(
        path,
        measure_name,
        f"AVERAGE({dax_column_ref(table.table_name, column['safe_name'])})",
        format_string=format_string,
    )
    return measure_name



def make_measure_chart(
    table: TableDefinition,
    category: dict[str, Any],
    measure_name: str,
    title: str,
    position: dict[str, Any],
    visual_type: str = "barChart",
) -> dict[str, Any]:
    """Bind a category to an explicit worker-evidence measure."""
    return {
        "kind": visual_type,
        "position": position,
        "query_state": {
            "Category": {
                "projections": [
                    projection(
                        column_expr(table.table_name, category["safe_name"]),
                        table.table_name,
                        category["safe_name"],
                    )
                ]
            },
            "Y": {
                "projections": [
                    projection(
                        measure_expr(table.table_name, measure_name),
                        table.table_name,
                        measure_name,
                        active=False,
                    )
                ]
            },
        },
        "title": title,
    }


def make_count_chart(
    table: TableDefinition,
    category: dict[str, Any],
    title: str,
    position: dict[str, Any],
    measure_name: str,
    visual_type: str = "barChart",
) -> dict[str, Any]:
    return {
        "kind": visual_type,
        "position": position,
        "query_state": {
            "Category": {
                "projections": [
                    projection(
                        column_expr(
                            table.table_name,
                            category["safe_name"],
                        ),
                        table.table_name,
                        category["safe_name"],
                    )
                ]
            },
            "Y": {
                "projections": [
                    projection(
                        measure_expr(
                            table.table_name,
                            measure_name,
                        ),
                        table.table_name,
                        measure_name,
                        active=False,
                    )
                ]
            },
        },
        "title": title,
    }


def make_table_visual(
    table: TableDefinition,
    columns: list[dict[str, Any]],
    title: str,
    position: dict[str, Any],
) -> dict[str, Any]:
    selected = columns if title == "Readiness Structure Matrix" else columns[:8]
    return {
        "kind": "tableEx",
        "position": position,
        "query_state": {
            "Values": {
                "projections": [
                    projection(
                        column_expr(
                            table.table_name,
                            c["safe_name"],
                        ),
                        table.table_name,
                        c["safe_name"],
                    )
                    for c in selected
                ]
            }
        },
        "title": title,
    }


def make_slicer(
    table: TableDefinition,
    column: dict[str, Any],
    title: str,
    position: dict[str, Any],
) -> dict[str, Any]:
    return {
        "kind": "slicer",
        "position": position,
        "query_state": {
            "Values": {
                "projections": [
                    projection(
                        column_expr(
                            table.table_name,
                            column["safe_name"],
                        ),
                        table.table_name,
                        column["safe_name"],
                    )
                ]
            }
        },
        "title": title,
        "slicer_column": column,
    }


def make_textbox(
    text: str,
    position: dict[str, Any],
    *,
    font_size: int = 18,
    textbox_height: int | None = None,
) -> dict[str, Any]:
    resolved_position = dict(position)

    if textbox_height is not None:
        resolved_position["height"] = int(textbox_height)

    return {
        "kind": "textbox",
        "position": resolved_position,
        "text": text,
        "font_size": int(font_size),
    }


def make_subtitle(
    text: str,
    position: dict[str, Any],
) -> dict[str, Any]:
    return make_textbox(text, position, font_size=11)


# ---------------------------------------------------------------------------
# Measure generation
# ---------------------------------------------------------------------------

def append_measure(
    tmdl_path: Path,
    name: str,
    expression: str,
    format_string: str = "#,##0",
) -> None:
    text = tmdl_path.read_text(encoding="utf-8").rstrip() + "\n"
    if re.search(
        rf"^\s*measure\s+{re.escape(tmdl_name(name))}\s*=",
        text,
        flags=re.MULTILINE,
    ):
        return

    text += (
        f"    measure {tmdl_name(name)} = {expression}\n"
        f"        formatString: {format_string}\n\n"
    )
    tmdl_path.write_text(text, encoding="utf-8")



def create_kpi_measures(
    semantic_definition: Path,
    registry: TableDefinitionRegistry,
) -> dict[str, Any]:
    """Create KPIs from explicit model/evidence populations.

    Every analytical KPI is tied to a documented worker-output contract.
    Heterogeneous worksheet row counts are never summed together.
    """
    analytical = registry.artifact_tables("analytical_dataset")
    if not analytical:
        raise RuntimeError("No analytical dataset was staged for Power BI.")

    primary = max(
        analytical,
        key=lambda t: (t.row_count, len(t.columns), t.table_name),
    )
    source_records = " + ".join(
        f"COUNTROWS({tmdl_name(t.table_name)})" for t in analytical
    )

    ensure_table_measure(
        primary, semantic_definition, "Source Records", source_records
    )

    overview = contract_table(
        registry, "analysis_results", "Table_Overview",
        ("table", "row_count", "column_count", "numeric_column_count",
         "temporal_column_count", "categorical_column_count", "duplicate_row_count"),
    )
    candidate = contract_table(
        registry, "analysis_results", "Demand_Candidate_Evidence",
        ("candidate_score", "measure_column"),
    )
    signal = contract_table(
        registry, "analysis_results", "Demand_Signal_Evidence",
        ("relationship_strength", "direction"),
    )
    forecasts = contract_table(
        registry, "forecast_results", "Forecasts",
        ("table", "value_column", "ForecastAverage", "ForecastConfidence",
         "DemandTrend", "RecentVsBaselineChangePct", "ForecastHorizonDays"),
    )
    recommendations = contract_table(
        registry, "recommendation_results", "Recommendations",
        ("recommendation_id", "recommendation_type", "finding", "recommendation",
         "evidence", "priority_level", "priority_score"),
    )

    bindings: dict[str, Any] = {
        "__measure_table__": primary.table_name,
        "__semantic_definition__": str(semantic_definition.resolve()),
        "Source Records": (primary.table_name, "Source Records"),
        "Discovered Structures": (
            overview.table_name,
            ensure_table_count_measure(
                overview, semantic_definition, "Discovered Structure Count"
            ),
        ) if overview else None,
        "Demand Candidates": (
            candidate.table_name,
            ensure_table_count_measure(
                candidate, semantic_definition, "Demand Candidate Count"
            ),
        ) if candidate else None,
        "Signal Relationships": (
            signal.table_name,
            ensure_table_count_measure(
                signal, semantic_definition, "Signal Relationship Count"
            ),
        ) if signal else None,
        "Forecast Series": (
            forecasts.table_name,
            ensure_table_count_measure(
                forecasts, semantic_definition, "Forecast Series Count"
            ),
        ) if forecasts else None,
        "Recommendations": (
            recommendations.table_name,
            ensure_table_count_measure(
                recommendations, semantic_definition, "Recommendation Count"
            ),
        ) if recommendations else None,
    }

    # Explicit evidence measures used by charts. These are averages/counts over
    # the worker contract itself; no generic SUM is applied to relationship
    # coefficients or candidate scores.
    if candidate:
        candidate_score = contract_column(candidate, "candidate_score")
        ensure_table_average_measure(
            candidate, semantic_definition, "Average Candidate Score", candidate_score
        )
        bindings["__candidate_score_measure__"] = (
            candidate.table_name, "Average Candidate Score"
        )

    if signal:
        relationship_strength = contract_column(signal, "relationship_strength")
        ensure_table_average_measure(
            signal, semantic_definition, "Average Relationship Strength",
            relationship_strength,
        )
        bindings["__relationship_strength_measure__"] = (
            signal.table_name, "Average Relationship Strength"
        )
        threshold = float(getattr(config, "DEMAND_STRONG_RELATIONSHIP_THRESHOLD", 0.7))
        ensure_table_measure(
            signal,
            semantic_definition,
            "Strong Signal Relationships",
            f"CALCULATE(COUNTROWS({tmdl_name(signal.table_name)}), "
            f"{dax_column_ref(signal.table_name, relationship_strength['safe_name'])} >= {threshold})",
        )
        bindings["__strong_relationship_measure__"] = (
            signal.table_name, "Strong Signal Relationships"
        )

    if forecasts:
        horizon = contract_column(forecasts, "ForecastHorizonDays")
        change = contract_column(forecasts, "RecentVsBaselineChangePct")
        ensure_table_average_measure(
            forecasts, semantic_definition, "Average Forecast Horizon", horizon,
            format_string="0"
        )
        ensure_table_average_measure(
            forecasts, semantic_definition, "Average Recent vs Baseline Change", change,
            format_string="0.00"
        )
        bindings["__forecast_horizon_measure__"] = (
            forecasts.table_name, "Average Forecast Horizon"
        )
        bindings["__forecast_change_measure__"] = (
            forecasts.table_name, "Average Recent vs Baseline Change"
        )
    if recommendations and any(
        c["source_name"] == "priority_score" for c in recommendations.columns
    ):
        ensure_table_average_measure(
            recommendations, semantic_definition, "Average Priority Score",
            contract_column(recommendations, "priority_score"),
        )
        bindings["__priority_score_measure__"] = (
            recommendations.table_name, "Average Priority Score"
        )

    return bindings




# ---------------------------------------------------------------------------
# Report layout and content contracts
# ---------------------------------------------------------------------------

PAGE_CONTRACTS = (
    (
        "Executive Intelligence",
        "What is the overall intelligence picture from this run?",
    ),
    (
        "Data Landscape",
        "What data and analytical capabilities were discovered?",
    ),
    (
        "Demand Intelligence",
        "What demand patterns and signal relationships were discovered?",
    ),
    (
        "Forecast Outlook",
        "What does the forecast indicate about future demand?",
    ),
    (
        "Recommendations & Actions",
        "What evidence-backed findings/actions were produced?",
    ),
    (
        "Signal Intelligence",
        "What market/supporting signals show evidence of relationship with demand?",
    ),
    (
        "Data Quality & Readiness",
        "Is the analytical foundation usable and ready?",
    ),
    (
        "Record Explorer",
        "Row-level investigation for the current run.",
    ),
)



def evidence_columns_for_display(
    profile: EvidenceProfile,
    maximum: int = 8,
) -> list[dict[str, Any]]:
    """Select compact row-explorer fields using only structural properties."""
    columns = list(profile.table.columns)

    def rank(c: dict[str, Any]) -> tuple[int, int, float, str]:
        kind = c["powerbi_type"]
        type_rank = {"dateTime": 0, "string": 1, "int64": 2, "double": 3, "boolean": 4}.get(kind, 5)
        avg_len = (
            sum(
                len(str(v))
                for v in c.get("sample_values", [])
                if nonempty(v)
            ) / max(1, int(c.get("sample_non_null", 0)))
        )
        # Prefer bounded categorical/date fields and short values; avoid large
        # free-form text columns in the first visible set.
        cardinality = int(c.get("sample_distinct", 0))
        text_penalty = 1 if avg_len > 120 else 0
        return (text_penalty, type_rank, cardinality, c["source_name"])

    columns.sort(key=rank)
    return columns[:maximum]



def create_page_visuals(
    page_name: str,
    page_question: str,
    registry: TableDefinitionRegistry,
    profiles: list[EvidenceProfile],
    measures: dict[str, Any],
) -> list[dict[str, Any]]:
    """Build the locked eight-page report with one consistent decision-support grid.

    Report semantics are sourced only from explicit worker-output contracts.
    Structural inference is used only for the row-level Record Explorer.
    """
    semantic_definition = Path(measures["__semantic_definition__"])
    visuals: list[dict[str, Any]] = []

    # Common 1280x720 grid: 24px outer margin, 12px gutter.
    CARD_Y = 125
    CARD_H = 92
    CARD_W = 190
    CARD_X = [24, 226, 428, 630, 832, 1034]
    PRIMARY_Y = 225
    PRIMARY_H = 250
    EVIDENCE_Y = 495
    EVIDENCE_H = 205
    LEFT_W = 590
    RIGHT_X = 642
    RIGHT_W = 590

    visuals.extend([
        make_textbox(
            page_name,
            {"x": 24, "y": 8, "z": 3000, "height": 28, "width": 1180, "tabOrder": 0},
            font_size=18,
            textbox_height=77,
        ),
        make_subtitle(
            f"{page_question}  \u2022  Current run",
            {"x": 24, "y": 86, "z": 3001, "height": 32, "width": 1180, "tabOrder": 1},
        ),
    ])

    def add_card(
        label: str,
        index: int,
        *,
        display_label: str | None = None,
        table_binding: tuple[str, str] | None = None,
    ) -> None:
        binding = table_binding or measures.get(label)
        if not binding or index >= len(CARD_X):
            return
        table_name, measure_name = binding
        visuals.append(make_card(
            registry.table(table_name),
            measure_name,
            display_label or label,
            {"x": CARD_X[index], "y": CARD_Y, "z": 1000 + index, "height": CARD_H, "width": CARD_W, "tabOrder": 100 + index},
        ))

    def add_count(table: TableDefinition, label: str) -> str:
        return ensure_table_count_measure(table, semantic_definition, label)

    def unavailable(message: str, x: int, y: int, width: int, height: int, z: int) -> None:
        visuals.append(make_textbox(
            message,
            {"x": x, "y": y, "z": z, "height": height, "width": width, "tabOrder": z},
            font_size=12,
        ))

    # ------------------------------------------------------------------
    # 01 Executive Intelligence
    # ------------------------------------------------------------------
    if page_name == "Executive Intelligence":
        for idx, (binding, label) in enumerate((
            ("Source Records", "Source Rows"),
            ("Discovered Structures", "Structures"),
            ("Demand Candidates", "Demand Candidates"),
            ("Signal Relationships", "Signal Links"),
            ("Forecast Series", "Forecast Series"),
            ("Recommendations", "Recommendations"),
        )):
            add_card(binding, idx, display_label=label)

        signal = contract_table(registry, "analysis_results", "Demand_Signal_Evidence",
            ("target_table", "target_measure_column", "signal_table", "signal_measure_column",
             "direction", "relationship_strength"))
        forecast = contract_table(registry, "forecast_results", "Forecasts",
            ("table", "value_column", "ForecastAverage", "ForecastConfidence", "DemandTrend",
             "RecentVsBaselineChangePct"))
        rec = contract_table(registry, "recommendation_results", "Recommendations",
            ("recommendation_type", "finding", "recommendation", "priority_level", "priority_score"))

        if signal and "__relationship_strength_measure__" in measures:
            visuals.append(make_measure_chart(
                signal, contract_column(signal, "direction"),
                measures["__relationship_strength_measure__"][1],
                "Average Relationship Strength by Direction",
                {"x": 24, "y": PRIMARY_Y, "z": 2000, "height": PRIMARY_H, "width": LEFT_W, "tabOrder": 2000},
                "columnChart"))
        else:
            unavailable("No worker-produced demandâ€“signal relationship evidence is available.", 24, PRIMARY_Y, LEFT_W, PRIMARY_H, 2000)

        if forecast:
            visuals.append(make_count_chart(
                forecast, contract_column(forecast, "DemandTrend"),
                "Forecast Trend Distribution",
                {"x": RIGHT_X, "y": PRIMARY_Y, "z": 2001, "height": PRIMARY_H, "width": RIGHT_W, "tabOrder": 2001},
                add_count(forecast, "Executive Forecast Trend Count"), "columnChart"))
        else:
            unavailable("Forecast evidence is unavailable for this run.", RIGHT_X, PRIMARY_Y, RIGHT_W, PRIMARY_H, 2001)


    # ------------------------------------------------------------------
    # 02 Data Landscape
    # ------------------------------------------------------------------
    elif page_name == "Data Landscape":
        overview = contract_table(registry, "analysis_results", "Table_Overview",
            ("table", "row_count", "column_count", "numeric_column_count",
             "temporal_column_count", "categorical_column_count", "duplicate_row_count"))
        if overview:
            definitions = [
                ("Structures", f"COUNTROWS({tmdl_name(overview.table_name)})", "Structures"),
                ("Records", f"SUM({dax_column_ref(overview.table_name, contract_column(overview, 'row_count')['safe_name'])})", "Records"),
                ("Columns", f"SUM({dax_column_ref(overview.table_name, contract_column(overview, 'column_count')['safe_name'])})", "Columns"),
                ("Numeric Fields", f"SUM({dax_column_ref(overview.table_name, contract_column(overview, 'numeric_column_count')['safe_name'])})", "Numeric"),
                ("Temporal Fields", f"SUM({dax_column_ref(overview.table_name, contract_column(overview, 'temporal_column_count')['safe_name'])})", "Temporal"),
                ("Categorical Fields", f"SUM({dax_column_ref(overview.table_name, contract_column(overview, 'categorical_column_count')['safe_name'])})", "Categorical"),
            ]
            for idx, (label, expression, display) in enumerate(definitions):
                m = ensure_table_measure(overview, semantic_definition, f"Landscape {label}", expression)
                add_card(label, idx, display_label=display, table_binding=(overview.table_name, m))

            visuals.append(make_chart(overview, contract_column(overview, "table"), contract_column(overview, "row_count"),
                "Records by Discovered Structure",
                {"x": 24, "y": PRIMARY_Y, "z": 2000, "height": PRIMARY_H, "width": LEFT_W, "tabOrder": 2000}, "barChart"))
            visuals.append(make_chart(overview, contract_column(overview, "table"), contract_column(overview, "numeric_column_count"),
                "Numeric Capability by Structure",
                {"x": RIGHT_X, "y": PRIMARY_Y, "z": 2001, "height": PRIMARY_H, "width": RIGHT_W, "tabOrder": 2001}, "barChart"))
            visuals.append(make_table_visual(overview, contract_columns(overview, (
                "table", "row_count", "column_count", "numeric_column_count", "temporal_column_count",
                "categorical_column_count", "duplicate_row_count")),
                "Discovered Structure Matrix",
                {"x": 24, "y": EVIDENCE_Y, "z": 2100, "height": EVIDENCE_H, "width": 1208, "tabOrder": 2100}))
        else:
            unavailable("Discovered structure evidence is unavailable for this run.", 24, PRIMARY_Y, 900, 100, 2000)

    # ------------------------------------------------------------------
    # 03 Demand Intelligence
    # ------------------------------------------------------------------
    elif page_name == "Demand Intelligence":
        candidates = contract_table(registry, "analysis_results", "Demand_Candidate_Evidence",
            ("table", "time_column", "measure_column", "observation_count", "temporal_coverage", "candidate_score", "evidence_type"))
        signal = contract_table(registry, "analysis_results", "Demand_Signal_Evidence",
            ("target_table", "target_measure_column", "signal_table", "signal_measure_column", "lag_periods",
             "direction", "relationship_strength", "aligned_observations"))
        if candidates:
            add_card("Demand Candidates", 0, display_label="Demand Candidates")
            if "__candidate_score_measure__" in measures:
                visuals.append(make_measure_chart(candidates, contract_column(candidates, "measure_column"),
                    measures["__candidate_score_measure__"][1], "Average Candidate Evidence Score",
                    {"x": 24, "y": PRIMARY_Y, "z": 2000, "height": PRIMARY_H, "width": LEFT_W, "tabOrder": 2000}, "barChart"))
            else:
                visuals.append(make_count_chart(candidates, contract_column(candidates, "measure_column"),
                    "Demand Candidate Distribution",
                    {"x": 24, "y": PRIMARY_Y, "z": 2000, "height": PRIMARY_H, "width": LEFT_W, "tabOrder": 2000},
                    add_count(candidates, "Demand Candidate Count by Measure"), "barChart"))
            if signal and "__relationship_strength_measure__" in measures:
                visuals.append(make_measure_chart(signal, contract_column(signal, "direction"),
                    measures["__relationship_strength_measure__"][1], "Relationship Strength by Direction",
                    {"x": RIGHT_X, "y": PRIMARY_Y, "z": 2001, "height": PRIMARY_H, "width": RIGHT_W, "tabOrder": 2001}, "columnChart"))
            elif signal:
                visuals.append(make_count_chart(signal, contract_column(signal, "direction"), "Relationship Evidence by Direction",
                    {"x": RIGHT_X, "y": PRIMARY_Y, "z": 2001, "height": PRIMARY_H, "width": RIGHT_W, "tabOrder": 2001},
                    add_count(signal, "Demand Signal Direction Count"), "columnChart"))
            else:
                unavailable("No demandâ€“signal relationship evidence is available.", RIGHT_X, PRIMARY_Y, RIGHT_W, PRIMARY_H, 2001)
        elif signal:
            add_card("Signal Relationships", 0, display_label="Signal Links")
            visuals.append(make_measure_chart(signal, contract_column(signal, "direction"), measures["__relationship_strength_measure__"][1],
                "Relationship Strength by Direction",
                {"x": 24, "y": PRIMARY_Y, "z": 2000, "height": PRIMARY_H, "width": LEFT_W, "tabOrder": 2000}, "columnChart"))
        else:
            unavailable("No demand-sensing evidence was published for this run.", 24, PRIMARY_Y, 1000, 100, 2000)

    # ------------------------------------------------------------------
    # 04 Forecast Outlook
    # ------------------------------------------------------------------
    elif page_name == "Forecast Outlook":
        forecast = contract_table(registry, "forecast_results", "Forecasts",
            ("table", "value_column", "ForecastAverage", "ForecastConfidence", "DemandTrend",
             "RecentVsBaselineChangePct", "ForecastHorizonDays"))
        if forecast:
            add_card("Forecast Series", 0, display_label="Forecast Series")
            if "__forecast_horizon_measure__" in measures:
                add_card(
                    "Forecast Series",
                    1,
                    display_label="Avg Horizon",
                    table_binding=measures["__forecast_horizon_measure__"]
                )
            if "__forecast_change_measure__" in measures:
                add_card("Forecast Series", 2, display_label="Avg Recent Change", table_binding=measures["__forecast_change_measure__"])
            visuals.append(make_count_chart(forecast, contract_column(forecast, "DemandTrend"), "Forecast Trend Distribution",
                {"x": 24, "y": PRIMARY_Y, "z": 2000, "height": PRIMARY_H, "width": LEFT_W, "tabOrder": 2000},
                add_count(forecast, "Forecast Trend Count"), "columnChart"))
            visuals.append(make_count_chart(forecast, contract_column(forecast, "ForecastConfidence"), "Forecast Confidence Distribution",
                {"x": RIGHT_X, "y": PRIMARY_Y, "z": 2001, "height": PRIMARY_H, "width": RIGHT_W, "tabOrder": 2001},
                add_count(forecast, "Forecast Confidence Count"), "barChart"))
        else:
            unavailable("Forecast capability is unavailable for this run. No forecast evidence was published.", 24, PRIMARY_Y, 1000, 100, 2000)

    # ------------------------------------------------------------------
    # 05 Recommendations & Actions
    # ------------------------------------------------------------------
    elif page_name == "Recommendations & Actions":
        rec = contract_table(registry, "recommendation_results", "Recommendations",
            ("recommendation_type", "finding", "recommendation", "evidence",
             "priority_level", "priority_score", "supporting_signal_count",
             "forecast_available", "entity_attribute", "entity_value",
             "entity_context"))

        if rec:
            # KPI cards
            recommendations_binding = measures.get("Recommendations")
            if recommendations_binding:
                table_name, measure_name = recommendations_binding
                visuals.append(
                    make_card(
                        registry.table(table_name),
                        measure_name,
                        "Recommendations",
                        {
                            "x": 24,
                            "y": 165,
                            "z": 1100,
                            "height": 92,
                            "width": 190,
                            "tabOrder": 110,
                        },
                    )
                )

            priority_binding = measures.get("__priority_score_measure__")
            if priority_binding:
                table_name, measure_name = priority_binding
                visuals.append(
                    make_card(
                        registry.table(table_name),
                        measure_name,
                        "Avg Priority",
                        {
                            "x": 226,
                            "y": 165,
                            "z": 1101,
                            "height": 92,
                            "width": 190,
                            "tabOrder": 111,
                        },
                    )
                )

            # Runtime-discovered recommendation entity filters.
            # These fields are created by the presentation adapter and are
            # deliberately not tied to company_id or any current dataset key.
            entity_attribute = next(
                (
                    column
                    for column in rec.columns
                    if column["source_name"] == "entity_attribute"
                ),
                None,
            )
            entity_value = next(
                (
                    column
                    for column in rec.columns
                    if column["source_name"] == "entity_value"
                ),
                None,
            )

            if entity_attribute:
                visuals.append(
                    make_slicer(
                        rec,
                        entity_attribute,
                        "Entity Type",
                        {
                            "x": 430,
                            "y": 165,
                            "z": 1200,
                            "height": 92,
                            "width": 250,
                            "tabOrder": 120,
                        },
                    )
                )

            if entity_value:
                visuals.append(
                    make_slicer(
                        rec,
                        entity_value,
                        "Entity",
                        {
                            "x": 696,
                            "y": 165,
                            "z": 1201,
                            "height": 92,
                            "width": 250,
                            "tabOrder": 121,
                        },
                    )
                )

            # Recommendation priority distribution
            visuals.append(
                make_count_chart(
                    rec,
                    contract_column(rec, "priority_level"),
                    "Recommendation Priority Distribution",
                    {
                        "x": 24,
                        "y": 275,
                        "z": 2000,
                        "height": 190,
                        "width": 590,
                        "tabOrder": 2000,
                    },
                    add_count(rec, "Recommendation Priority Count"),
                    "columnChart",
                )
            )

            # Recommendation type distribution
            visuals.append(
                make_count_chart(
                    rec,
                    contract_column(rec, "recommendation_type"),
                    "Finding / Recommendation Type",
                    {
                        "x": 642,
                        "y": 275,
                        "z": 2001,
                        "height": 190,
                        "width": 590,
                        "tabOrder": 2001,
                    },
                    add_count(rec, "Recommendation Type Count"),
                    "barChart",
                )
            )

            # Evidence-backed recommendation table
            visuals.append(
                make_table_visual(
                    rec,
                    contract_columns(
                        rec,
                        (
                            "entity_attribute",
                            "entity_value",
                            "priority_level",
                            "finding",
                            "recommendation",
                            "evidence_summary",
                        ),
                    ),
                    "Evidence-Backed Recommendations",
                    {
                        "x": 24,
                        "y": 485,
                        "z": 2100,
                        "height": 210,
                        "width": 1208,
                        "tabOrder": 2100,
                    },
                )
            )

        else:
            unavailable(
                "Recommendation capability is unavailable for this run. No recommendation evidence was published.",
                24,
                PRIMARY_Y,
                1000,
                100,
                2000,
            )
    # ------------------------------------------------------------------
    # 06 Signal Intelligence
    # ------------------------------------------------------------------
    elif page_name == "Signal Intelligence":
        signal = contract_table(registry, "analysis_results", "Demand_Signal_Evidence",
            ("target_table", "target_measure_column", "signal_table", "signal_measure_column", "lag_periods",
             "direction", "relationship_strength", "aligned_observations"))
        if signal:
            add_card("Signal Relationships", 0, display_label="Signal Links")
            if "__relationship_strength_measure__" in measures:
                visuals.append(make_measure_chart(signal, contract_column(signal, "direction"),
                    measures["__relationship_strength_measure__"][1], "Average Signal Relationship Strength",
                    {"x": 24, "y": PRIMARY_Y, "z": 2000, "height": PRIMARY_H, "width": LEFT_W, "tabOrder": 2000}, "columnChart"))
            visuals.append(make_count_chart(signal, contract_column(signal, "direction"), "Relationship Evidence by Direction",
                {"x": RIGHT_X, "y": PRIMARY_Y, "z": 2001, "height": PRIMARY_H, "width": RIGHT_W, "tabOrder": 2001},
                add_count(signal, "Signal Direction Count"), "barChart"))
        else:
            unavailable("Signal intelligence is unavailable for this run. No worker-produced signal evidence was published.", 24, PRIMARY_Y, 1000, 100, 2000)

    # ------------------------------------------------------------------
    # 07 Data Quality & Readiness
    # ------------------------------------------------------------------
    elif page_name == "Data Quality & Readiness":
        overview = contract_table(registry, "analysis_results", "Table_Overview",
            ("table", "row_count", "column_count", "duplicate_row_count"))
        schema = contract_table(registry, "analysis_results", "Schema_Evidence",
            ("table", "missing_count", "missing_ratio", "non_null_count"))
        readiness = contract_table(
            registry, "analysis_results", "Analytical_Readiness",
            (
                "table",
                "row_count",
                "column_count",
                "numeric_fields_available",
                "temporal_fields_available",
                "dimension_fields_available",
                "identifier_candidates_available",
                "demand_series_candidates",
                "demand_signal_relationships",
                "has_numeric_analysis",
                "has_temporal_analysis",
                "has_dimension_analysis",
                "has_relationship_evidence",
                "has_demand_candidate_evidence",
                "has_demand_signal_evidence",
            )
        )
        if overview:
            records = ensure_table_sum_measure(overview, semantic_definition, "Quality Records", contract_column(overview, "row_count"))
            duplicates = ensure_table_sum_measure(overview, semantic_definition, "Duplicate Records", contract_column(overview, "duplicate_row_count"))
            add_card("Quality Records", 0, display_label="Records", table_binding=(overview.table_name, records))
            add_card("Quality Duplicates", 1, display_label="Duplicates", table_binding=(overview.table_name, duplicates))
            if schema:
                analyzed = ensure_table_count_measure(schema, semantic_definition, "Analyzed Columns")
                missing_avg = ensure_table_average_measure(schema, semantic_definition, "Average Missing Rate", contract_column(schema, "missing_ratio"), format_string="0.00%")
                add_card("Analyzed Columns", 2, display_label="Analyzed Columns", table_binding=(schema.table_name, analyzed))
                add_card("Average Missing Rate", 3, display_label="Avg Missing Rate", table_binding=(schema.table_name, missing_avg))
                visuals.append(make_measure_chart(schema, contract_column(schema, "table"), missing_avg, "Average Missing Rate by Structure",
                    {"x": 24, "y": PRIMARY_Y, "z": 2000, "height": PRIMARY_H, "width": LEFT_W, "tabOrder": 2000}, "barChart"))
            visuals.append(make_chart(overview, contract_column(overview, "table"), contract_column(overview, "duplicate_row_count"), "Duplicate Records by Structure",
                {"x": RIGHT_X, "y": PRIMARY_Y, "z": 2001, "height": PRIMARY_H, "width": RIGHT_W, "tabOrder": 2001}, "barChart"))
            if readiness:
                visuals.append(make_table_visual(
                    readiness,
                    contract_columns(
                        readiness,
                        (
                            "table",
                            "row_count",
                            "column_count",
                            "numeric_fields_available",
                            "temporal_fields_available",
                            "dimension_fields_available",
                            "identifier_candidates_available",
                            "demand_series_candidates",
                            "demand_signal_relationships",
                            "has_numeric_analysis",
                            "has_temporal_analysis",
                            "has_dimension_analysis",
                            "has_relationship_evidence",
                            "has_demand_candidate_evidence",
                            "has_demand_signal_evidence",
                        )
                    ),
                    "Readiness Structure Matrix",
                    {
                        "x": 24,
                        "y": EVIDENCE_Y,
                        "z": 2100,
                        "height": EVIDENCE_H,
                        "width": 1208,
                        "tabOrder": 2100
                    }
                ))
        else:
            unavailable("Data-quality evidence is unavailable for this run. Only metrics actually produced by validation are shown.", 24, PRIMARY_Y, 1000, 100, 2000)

    # ------------------------------------------------------------------
    # 08 Record Explorer â€” the only deliberately row-heavy page.
    # ------------------------------------------------------------------
    elif page_name == "Record Explorer":
        source_tables = registry.artifact_tables("analytical_dataset")
        source_tables.sort(key=lambda t: (t.row_count, len(t.columns), t.table_name), reverse=True)
        if source_tables:
            explorer = source_tables[0]
            profile = profile_for(profiles, explorer)
            filter_columns = [c for c in profile.categorical if 2 <= int(c.get("sample_distinct", 0)) <= 50][:2]
            temporal = profile.temporal[0] if profile.temporal else None
            for idx, column in enumerate(filter_columns):
                visuals.append(make_slicer(explorer, column, column["source_name"],
                    {"x": 24 + idx * 260, "y": 125, "z": 1000 + idx, "height": 92, "width": 240, "tabOrder": 100 + idx}))
            if temporal:
                visuals.append(make_slicer(explorer, temporal, temporal["source_name"],
                    {"x": 24 + len(filter_columns) * 260, "y": 125, "z": 1002, "height": 92, "width": 240, "tabOrder": 102}))
            filtered_measure = ensure_table_measure(explorer, semantic_definition, "Explorer Filtered Records", f"COUNTROWS({tmdl_name(explorer.table_name)})")
            visuals.append(make_card(explorer, filtered_measure, "Filtered Records",
                {"x": 1000, "y": 125, "z": 1100, "height": 92, "width": 208, "tabOrder": 110}))
            display_columns = evidence_columns_for_display(profile, 8)
            visuals.append(make_table_visual(explorer, display_columns, "Current-Run Record Inspection",
                {"x": 24, "y": 225, "z": 2000, "height": 495, "width": 1208, "tabOrder": 2000}))
        else:
            unavailable("No row-level analytical dataset is available for Record Explorer in this run.", 24, PRIMARY_Y, 1000, 100, 2000)

    if len(visuals) > MAX_REPORT_VISUALS_PER_PAGE:
        raise RuntimeError(
            f"Page {page_name!r} generated {len(visuals)} visuals; limit is {MAX_REPORT_VISUALS_PER_PAGE}."
        )
    return visuals


def render_visual(
    page_dir: Path,
    visual_spec: dict[str, Any],
    visual_id: str,
) -> None:
    kind = visual_spec["kind"]
    path = page_dir / "visuals" / visual_id / "visual.json"
    position = visual_spec["position"]

    if kind == "textbox":
        payload = {
            "$schema": VISUAL_SCHEMA,
            "name": visual_id,
            "position": position,
            "visual": {
                "visualType": "textbox",
                "objects": {
                    "general": [
                        {
                            "properties": {
                                "paragraphs": [
                                    {
                                        "textRuns": [
                                            {
                                                "value": visual_spec["text"],
                                                "textStyle": {
                                                    "fontSize": int(visual_spec.get("font_size", 18)),
                                                    "fontFamily": "Segoe UI",
                                                },
                                            }
                                        ]
                                    }
                                ]
                            }
                        }
                    ]
                },
                "visualContainerObjects": {
                    "visualHeader": [
                        {
                            "properties": {
                                "show": literal_value("false")
                            }
                        }
                    ]
                },
            },
        }
        json_write(path, payload)
        return

    if kind == "cardVisual":
        write_visual(
            path,
            name=visual_id,
            visual_type="cardVisual",
            position=position,
            query_state=visual_spec["query_state"],
            title=visual_spec["title"],
            visual_container_objects={
                "padding": [
                    {
                        "properties": {
                            "top": literal_value("0L"),
                            "bottom": literal_value("0L"),
                            "left": literal_value("0L"),
                            "right": literal_value("0L"),
                        }
                    }
                ]
            },
        )
        return

    if kind in {"barChart", "columnChart", "lineChart"}:
        write_visual(
            path,
            name=visual_id,
            visual_type=kind,
            position=position,
            query_state=visual_spec["query_state"],
            title=visual_spec["title"],
        )
        return

    if kind == "tableEx":
        write_visual(
            path,
            name=visual_id,
            visual_type="tableEx",
            position=position,
            query_state=visual_spec["query_state"],
            title=visual_spec["title"],
        )
        return

    if kind == "slicer":
        write_visual(
            path,
            name=visual_id,
            visual_type="slicer",
            position=position,
            query_state=visual_spec["query_state"],
            title=visual_spec["title"],
            visual_container_objects={
                "visualHeader": [
                    {
                        "properties": {
                            "show": literal_value("false")
                        }
                    }
                ]
            },
            objects={
                "data": [
                    {
                        "properties": {
                            "mode": literal_value("'Dropdown'")
                        }
                    }
                ]
            },
        )
        return

    raise RuntimeError(f"Unsupported generated visual kind: {kind}")


def write_pages(
    report_definition: Path,
    page_specs: list[tuple[str, str, list[dict[str, Any]]]],
) -> dict[str, Any]:
    pages_root = report_definition / "pages"
    pages_root.mkdir(parents=True, exist_ok=True)

    page_ids: list[str] = []
    visual_manifest: list[dict[str, Any]] = []
    used_visual_ids: set[str] = set()

    for display_name, question, visuals in page_specs:
        page_id = uuid.uuid4().hex[:20]
        page_ids.append(page_id)

        page_dir = pages_root / page_id
        (page_dir / "visuals").mkdir(
            parents=True,
            exist_ok=True,
        )

        json_write(
            page_dir / "page.json",
            {
                "$schema": PAGE_SCHEMA,
                "name": page_id,
                "displayName": display_name,
                "displayOption": "FitToPage",
                "height": PAGE_HEIGHT,
                "width": PAGE_WIDTH,
                "annotations": [
                    {
                        "name": "PageQuestion",
                        "value": question,
                    }
                ],
            },
        )

        page_visual_entries: list[str] = []

        for visual_index, spec in enumerate(visuals):
            visual_id = uuid.uuid4().hex[:20]
            while visual_id in used_visual_ids:
                visual_id = uuid.uuid4().hex[:20]
            used_visual_ids.add(visual_id)

            render_visual(
                page_dir,
                spec,
                visual_id,
            )
            page_visual_entries.append(visual_id)

            visual_manifest.append(
                {
                    "page_id": page_id,
                    "page_name": display_name,
                    "visual_id": visual_id,
                    "kind": spec["kind"],
                    "position": spec["position"],
                    "title": spec.get("title"),
                    "index": visual_index,
                }
            )

        if not page_visual_entries:
            raise RuntimeError(
                f"Report page {display_name!r} has no visuals."
            )

    json_write(
        pages_root / "pages.json",
        {
            "$schema": PAGES_METADATA_SCHEMA,
            "pageOrder": page_ids,
            "activePageName": page_ids[0],
        },
    )

    return {
        "page_ids": page_ids,
        "visual_manifest": visual_manifest,
    }


# ---------------------------------------------------------------------------
# Report generation
# ---------------------------------------------------------------------------

def generate_report(
    project_root: Path,
    model_meta: dict[str, Any],
    registry: TableDefinitionRegistry,
) -> dict[str, Any]:
    report_root = project_root / model_meta["report_name"]
    definition = report_root / "definition"
    definition.mkdir(parents=True, exist_ok=True)

    json_write(
        report_root / "definition.pbir",
        {
            "$schema": PBIR_SCHEMA,
            "version": "4.0",
            "datasetReference": {
                "byPath": {
                    "path": f"../{model_meta['semantic_name']}"
                }
            },
        },
    )

    json_write(
        definition / "version.json",
        {
            "$schema": PBIR_VERSION_SCHEMA,
            "version": "2.0.0",
        },
    )

    # Keep the report metadata inside the working PBIR scaffold. No report.json
    # is written at the report root, because that would be PBIR-Legacy and would
    # conflict with the PBIR definition/ folder.
    json_write(
        definition / "report.json",
        {
            "$schema": REPORT_SCHEMA,
            "themeCollection": {},
            "filterConfig": {},
            "objects": {},
            "settings": {
                "useEnhancedTooltips": True,
            },
            "resourcePackages": [],
            "annotations": [],
        },
    )

    semantic_definition = Path(model_meta["semantic_definition"])
    measures = create_kpi_measures(
        semantic_definition,
        registry,
    )
    measures["__semantic_definition__"] = str(
        semantic_definition.resolve()
    )

    profiles = build_evidence_profiles(registry)

    signal_supported = worker_contract_available(
        registry,
        "analysis_results",
        "Demand_Signal_Evidence",
        ("target_table", "target_measure_column", "signal_table",
         "signal_measure_column", "lag_periods", "direction",
         "relationship_strength", "aligned_observations"),
    )

    page_contracts = [
        (name, question)
        for name, question in PAGE_CONTRACTS
        if name != "Signal Intelligence" or signal_supported
    ]

    page_specs: list[
        tuple[str, str, list[dict[str, Any]]]
    ] = []

    for page_name, question in page_contracts:
        visuals = create_page_visuals(
            page_name,
            question,
            registry,
            profiles,
            measures,
        )

        # Analytical pages must contain an insight-capable visual. Record
        # Explorer is the only page where a broad table is sufficient.
        if page_name != "Record Explorer":
            meaningful = any(
                v["kind"] in {
                    "cardVisual",
                    "barChart",
                    "columnChart",
                    "lineChart",
                    "textbox",
                }
                for v in visuals
            )
            if not meaningful:
                raise RuntimeError(
                    f"Page {page_name!r} failed the meaningful-content gate."
                )

        page_specs.append(
            (
                page_name,
                question,
                visuals,
            )
        )

    pages_meta = write_pages(
        definition,
        page_specs,
    )

    return {
        "pages": len(page_specs),
        "visuals": len(
            pages_meta["visual_manifest"]
        ),
        "page_contracts": [
            {
                "name": name,
                "question": question,
            }
            for name, question in page_contracts
        ],
        "signal_page_supported": signal_supported,
        "primary_table": measures["__measure_table__"],
        "measures": [
            key
            for key in measures
            if not key.startswith("__")
        ],
        "report_design": (
            "capability-driven-8-page-decision-support"
        ),
        "source_document_alignment": {
            "questions": [
                {
                    "question": "What is Demand Sensing?",
                    "report_mapping": "Report architecture combines historical evidence with current signal evidence and forecasting.",
                    "status": "conceptual_alignment",
                },
                {
                    "question": "What happened last year?",
                    "report_mapping": "Historical evidence is shown only when an explicit worker evidence contract publishes it; no unsupported historical answer is fabricated.",
                    "status": "evidence_gated",
                },
                {
                    "question": "What is happening today that will affect sales tomorrow?",
                    "report_mapping": "Demand Intelligence, Signal Intelligence, and Forecast Outlook connect current-run signal relationships to forecast evidence.",
                    "status": "evidence_gated",
                },
                {
                    "question": "What are Market Signals?",
                    "report_mapping": "Signal Intelligence displays only worker-produced signal/demand evidence; the six example categories in the source document are not assumed to exist in every dataset.",
                    "status": "evidence_gated",
                },
            ],
            "signal_categories_from_source": [
                "Economic Indicators",
                "Commodity Prices",
                "Weather",
                "Web Search Trends",
                "Competitor Activity",
                "Customer Order Patterns",
            ],
        },
        "definition": str(
            definition.resolve()
        ),
        "visual_manifest": pages_meta["visual_manifest"],
    }


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def _path_is_within(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def _load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8")
        )
    except Exception as exc:
        raise RuntimeError(
            f"Invalid JSON: {path}: {exc}"
        ) from exc


def validate_tmdl_file(
    path: Path,
    table: TableDefinition,
    staged_root: Path,
) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8")

    declaration = re.search(
        r"^table\s+(.+?)\s*$",
        text,
        flags=re.MULTILINE,
    )
    if not declaration:
        raise RuntimeError(
            f"TMDL table declaration missing: {path}"
        )

    if declaration.group(1).strip() != tmdl_name(
        table.table_name
    ):
        raise RuntimeError(
            f"TMDL table declaration does not match registry: {path}"
        )

    declared_columns = re.findall(
        r"^\s+column\s+(.+?)\s*$",
        text,
        flags=re.MULTILINE,
    )
    expected_columns = [
        tmdl_name(c["safe_name"])
        for c in table.columns
    ]

    if sorted(declared_columns) != sorted(expected_columns):
        raise RuntimeError(
            f"TMDL column registry mismatch: {path}"
        )

    if len(declared_columns) != len(set(declared_columns)):
        raise RuntimeError(
            f"Duplicate TMDL column declarations: {path}"
        )

    measure_names = re.findall(
        r"^\s+measure\s+(.+?)\s*=",
        text,
        flags=re.MULTILINE,
    )
    if len(measure_names) != len(set(measure_names)):
        raise RuntimeError(
            f"Duplicate TMDL measure declarations: {path}"
        )

    partition = re.search(
        r"^\s+partition\s+.+\s*=\s*m\s*$",
        text,
        flags=re.MULTILINE,
    )
    if not partition:
        raise RuntimeError(
            f"Missing/invalid import partition: {path}"
        )

    if "mode: import" not in text:
        raise RuntimeError(
            f"Import partition has no import mode: {path}"
        )

    references = re.findall(
        r'File\.Contents\("([^"]+)"\)',
        text,
    )
    if not references:
        raise RuntimeError(
            f"TMDL partition contains no File.Contents reference: {path}"
        )

    for reference in references:
        source = Path(reference)
        if not _path_is_within(source, staged_root):
            raise RuntimeError(
                "TMDL source escapes the current-run Power BI staging "
                f"directory: {source}"
            )
        if not source.is_file() or source.stat().st_size == 0:
            raise RuntimeError(
                f"TMDL source is missing/empty: {source}"
            )

    return {
        "path": str(path),
        "table": table.table_name,
        "columns": len(declared_columns),
        "measures": len(measure_names),
        "source_references": len(references),
    }


def collect_model_field_registry(
    registry: TableDefinitionRegistry,
    semantic_definition: Path,
) -> dict[str, set[str]]:
    fields: dict[str, set[str]] = {}
    for table in registry.tables:
        names = {c["safe_name"] for c in table.columns}

        tmdl_path = registry.tmdl_path(
            semantic_definition,
            table.table_name,
        )
        text = tmdl_path.read_text(encoding="utf-8")
        names.update(
            (
                m.group(1).strip()[1:-1].replace("''", "'")
                if m.group(1).strip().startswith("'")
                and m.group(1).strip().endswith("'")
                else m.group(1).strip()
            )
            for m in re.finditer(
                r"^\s+measure\s+(.+?)\s*=",
                text,
                flags=re.MULTILINE,
            )
        )
        fields[table.table_name] = names
    return fields


def validate_expression_bindings(
    payload: Any,
    fields: dict[str, set[str]],
    path: Path,
) -> None:
    """Validate every Column/Measure expression against the canonical registry."""
    if isinstance(payload, dict):
        for key, value in payload.items():
            if key in {"Column", "Measure"} and isinstance(value, dict):
                expression = value.get("Expression", {})
                source_ref = expression.get("SourceRef", {})
                entity = source_ref.get("Entity")
                prop = value.get("Property")

                if not isinstance(entity, str) or not entity:
                    raise RuntimeError(
                        f"Missing SourceRef.Entity in {path}"
                    )
                if entity not in fields:
                    raise RuntimeError(
                        f"Visual references nonexistent table "
                        f"{entity!r}: {path}"
                    )
                if not isinstance(prop, str) or prop not in fields[entity]:
                    raise RuntimeError(
                        f"Visual references nonexistent field "
                        f"{entity}.{prop}: {path}"
                    )

            validate_expression_bindings(
                value,
                fields,
                path,
            )
    elif isinstance(payload, list):
        for item in payload:
            validate_expression_bindings(
                item,
                fields,
                path,
            )


def validate_page_and_visuals(
    project_root: Path,
    report_definition: Path,
    registry: TableDefinitionRegistry,
) -> dict[str, Any]:
    pages_root = report_definition / "pages"
    pages_json = pages_root / "pages.json"
    if not pages_json.is_file():
        raise RuntimeError(
            f"Missing pages.json: {pages_json}"
        )

    pages_meta = _load_json(pages_json)
    page_order = pages_meta.get("pageOrder")
    if not isinstance(page_order, list) or not page_order:
        raise RuntimeError(
            "PBIR pageOrder is missing or empty."
        )

    semantic_definition = Path(
        project_root
        / "Model.SemanticModel"
        / "definition"
    )
    fields = collect_model_field_registry(
        registry,
        semantic_definition,
    )

    total_visuals = 0
    page_summaries: list[dict[str, Any]] = []

    for page_id in page_order:
        if not isinstance(page_id, str):
            raise RuntimeError(
                "Page order contains a non-string page identifier."
            )

        page_dir = pages_root / page_id
        page_json = page_dir / "page.json"
        visuals_dir = page_dir / "visuals"

        if not page_json.is_file():
            raise RuntimeError(
                f"Page definition missing: {page_json}"
            )

        page = _load_json(page_json)
        if page.get("name") != page_id:
            raise RuntimeError(
                f"Page identity mismatch: {page_json}"
            )
        if page.get("displayOption") not in {
            "FitToPage",
            "FitToWidth",
            "ActualSize",
        }:
            raise RuntimeError(
                f"Unsupported page display option: {page_json}"
            )
        if (
            page.get("width") != PAGE_WIDTH
            or page.get("height") != PAGE_HEIGHT
        ):
            raise RuntimeError(
                f"Unexpected page canvas size: {page_json}"
            )

        visual_paths = sorted(
            visuals_dir.glob("*/visual.json")
        )
        if not visual_paths:
            raise RuntimeError(
                f"Page contains no visuals: {page_json}"
            )
        if len(visual_paths) > MAX_REPORT_VISUALS_PER_PAGE:
            raise RuntimeError(
                f"Page exceeds visual limit: {page_json}"
            )

        rectangles: list[tuple[float, float, float, float, str]] = []
        kinds: list[str] = []

        for visual_path in visual_paths:
            visual = _load_json(visual_path)
            total_visuals += 1

            if visual.get("name") != visual_path.parent.name:
                raise RuntimeError(
                    f"Visual identity mismatch: {visual_path}"
                )
            if visual.get("$schema") != VISUAL_SCHEMA:
                raise RuntimeError(
                    f"Unexpected visual schema: {visual_path}"
                )

            # Explicit regression guard for the known Desktop filterConfig
            # failure. It is legal at report/page scope but must not appear
            # inside the visual object.
            if "filterConfig" in visual.get("visual", {}):
                raise RuntimeError(
                    f"Invalid visual.filterConfig detected: {visual_path}"
                )

            position = visual.get("position", {})
            try:
                x = float(position["x"])
                y = float(position["y"])
                width = float(position["width"])
                height = float(position["height"])
            except Exception as exc:
                raise RuntimeError(
                    f"Invalid visual position: {visual_path}"
                ) from exc

            if width <= 0 or height <= 0:
                raise RuntimeError(
                    f"Visual has non-positive dimensions: {visual_path}"
                )
            if (
                x < 0
                or y < 0
                or x + width > PAGE_WIDTH
                or y + height > PAGE_HEIGHT
            ):
                raise RuntimeError(
                    f"Visual is outside page bounds: {visual_path}"
                )

            rectangles.append(
                (x, y, width, height, str(visual_path))
            )

            visual_type = visual.get("visual", {}).get(
                "visualType"
            )
            if not isinstance(visual_type, str):
                raise RuntimeError(
                    f"Visual type missing: {visual_path}"
                )
            kinds.append(visual_type)

            if "query" in visual.get("visual", {}):
                validate_expression_bindings(
                    visual["visual"]["query"],
                    fields,
                    visual_path,
                )

            # All projections must be under queryState.
            query = visual.get("visual", {}).get("query")
            if query is not None:
                if "queryState" not in query:
                    raise RuntimeError(
                        f"Visual query has no queryState: {visual_path}"
                    )
                for role_name, role in query["queryState"].items():
                    if not isinstance(role, dict):
                        raise RuntimeError(
                            f"Invalid query role {role_name}: {visual_path}"
                        )
                    if "projections" not in role:
                        raise RuntimeError(
                            f"Query role has no projections "
                            f"{role_name}: {visual_path}"
                        )

        # Layout contract beyond rectangle overlap: reserve the header band,
        # enforce the KPI/content bands, and keep narrow tables intentionally
        # small. This catches clipping/collision classes that geometry alone
        # cannot see.
        for visual_path in visual_paths:
            visual = _load_json(visual_path)
            position = visual["position"]
            visual_type = visual["visual"]["visualType"]
            x = float(position["x"])
            y = float(position["y"])
            width = float(position["width"])
            height = float(position["height"])

            if visual_type == "textbox":
                text_value = ""
                paragraphs = (
                    visual.get("visual", {})
                    .get("objects", {})
                    .get("general", [])
                )
                if paragraphs:
                    try:
                        text_runs = paragraphs[0]["properties"]["paragraphs"][0]["textRuns"]
                        text_value = "".join(str(r.get("value", "")) for r in text_runs)
                    except Exception:
                        text_value = ""

                if y < 70:
                    if y < 38 and height > 77:
                        raise RuntimeError(
                            f"Header title band is too tall: {visual_path}"
                        )
                    if 38 <= y < 70 and height > 26:
                        raise RuntimeError(
                            f"Header subtitle band is too tall: {visual_path}"
                        )
                elif y < 70:
                    raise RuntimeError(
                        f"Non-header textbox occupies reserved header band: {visual_path}"
                    )

            if y < 70 and visual_type != "textbox":
                raise RuntimeError(
                    f"Non-header visual intrudes into header band: {visual_path}"
                )

            if visual_type == "cardVisual":
                title_objects = visual.get("visual", {}).get(
                    "visualContainerObjects", {}
                ).get("title", [])
                title_text = ""
                if title_objects:
                    try:
                        title_text = str(
                            title_objects[0]["properties"]["text"]["expr"]["Literal"]["Value"]
                        ).strip("'")
                    except Exception:
                        title_text = ""
                if width < 180:
                    raise RuntimeError(
                        f"KPI card is too narrow: {visual_path}"
                    )
                if len(title_text) > 20:
                    raise RuntimeError(
                        f"KPI title is likely to clip at its generated width: "
                        f"{title_text!r} ({visual_path})"
                    )

            if visual_type == "tableEx":
                query = visual.get("visual", {}).get("query", {})
                projections = (
                    query.get("queryState", {})
                    .get("Values", {})
                    .get("projections", [])
                )
                if width <= 400 and len(projections) > 4:
                    raise RuntimeError(
                        f"Narrow evidence table has too many columns: "
                        f"{len(projections)} ({visual_path})"
                    )
                if width <= 600 and len(projections) > 6:
                    raise RuntimeError(
                        f"Medium evidence table has too many columns: "
                        f"{len(projections)} ({visual_path})"
                    )

            # Content-band rules. Analytical visuals belong below the header;
            # KPI cards stay in the KPI band unless they are intentionally in
            # Record Explorer.
            if visual_type != "textbox" and y < 70:
                raise RuntimeError(
                    f"Visual is above the content safe area: {visual_path}"
                )

        # Hard fail actual overlap. Touching edges is not overlap.
        for i in range(len(rectangles)):
            ax, ay, aw, ah, ap = rectangles[i]
            for j in range(i + 1, len(rectangles)):
                bx, by, bw, bh, bp = rectangles[j]
                horizontal = min(ax + aw, bx + bw) - max(ax, bx)
                vertical = min(ay + ah, by + bh) - max(ay, by)
                if horizontal > 0 and vertical > 0:
                    raise RuntimeError(
                        "Visual overlap detected between "
                        f"{ap} and {bp}"
                    )

        page_summaries.append(
            {
                "page_id": page_id,
                "display_name": page.get("displayName"),
                "visual_count": len(visual_paths),
                "visual_types": kinds,
            }
        )

    return {
        "page_count": len(page_order),
        "visual_count": total_visuals,
        "pages": page_summaries,
    }


def validate_powerbi_project(
    project_root: Path,
    model_meta: dict[str, Any],
    registry: TableDefinitionRegistry,
    context: Any,
) -> dict[str, Any]:
    project_name = project_root.name

    if project_name != f"Demand_Intelligence_Report_{context.run_id}":
        raise RuntimeError(
            f"Power BI project name is not current-run canonical name: "
            f"{project_name}"
        )

    report_root = project_root / model_meta["report_name"]
    semantic_root = project_root / model_meta["semantic_name"]
    report_definition = report_root / "definition"
    semantic_definition = semantic_root / "definition"

    required = [
        project_root / f"{project_name}.pbip",
        report_root / "definition.pbir",
        report_definition / "version.json",
        report_definition / "report.json",
        report_definition / "pages" / "pages.json",
        semantic_root / "definition.pbism",
        semantic_definition / "database.tmdl",
        semantic_definition / "model.tmdl",
        report_root / "StaticResources",
    ]
    missing = [
        str(path)
        for path in required
        if not path.exists()
    ]
    if missing:
        raise RuntimeError(
            f"Power BI project is missing required artifacts: {missing}"
        )

    # JSON/PBIP syntax.
    json_files = list(
        project_root.rglob("*.json")
    ) + list(
        project_root.rglob("*.pbip")
    )
    for path in json_files:
        _load_json(path)

    pbip = _load_json(
        project_root / f"{project_name}.pbip"
    )
    if pbip.get("$schema") != PBIP_SCHEMA:
        raise RuntimeError(
            "PBIP schema is not the supported 1.0.0 contract."
        )

    pbir_path = report_root / "definition.pbir"
    pbir = _load_json(pbir_path)
    if pbir.get("$schema") != PBIR_SCHEMA:
        raise RuntimeError(
            "definition.pbir schema mismatch."
        )
    if pbir.get("version") != "4.0":
        raise RuntimeError(
            "definition.pbir must use PBIR format version 4.0."
        )

    by_path = (
        pbir.get("datasetReference", {})
        .get("byPath", {})
    )
    semantic_path = by_path.get("path")
    if not isinstance(semantic_path, str):
        raise RuntimeError(
            "definition.pbir has no byPath semantic-model reference."
        )
    if Path(semantic_path).is_absolute():
        raise RuntimeError(
            "definition.pbir semantic-model reference must be relative."
        )

    resolved_semantic = (
        pbir_path.parent / semantic_path
    ).resolve()
    if resolved_semantic != semantic_root.resolve():
        raise RuntimeError(
            "definition.pbir points to the wrong semantic-model directory."
        )

    if not resolved_semantic.is_dir():
        raise RuntimeError(
            f"Referenced semantic model does not exist: {resolved_semantic}"
        )

    version = _load_json(
        report_definition / "version.json"
    )
    if version.get("$schema") != PBIR_VERSION_SCHEMA:
        raise RuntimeError(
            "PBIR version.json schema mismatch."
        )
    if version.get("version") != "2.0.0":
        raise RuntimeError(
            "PBIR version.json must be 2.0.0."
        )

    pbism = _load_json(
        semantic_root / "definition.pbism"
    )
    if pbism.get("$schema") != PBISM_SCHEMA:
        raise RuntimeError(
            "definition.pbism schema mismatch."
        )

    # Canonical registry/TMDL consistency.
    table_dir = semantic_definition / "tables"
    actual_table_files = sorted(
        table_dir.glob("*.tmdl")
    )
    expected_files = sorted(
        table_dir / t.tmdl_filename
        for t in registry.tables
    )

    if actual_table_files != expected_files:
        raise RuntimeError(
            "Canonical table registry does not match emitted TMDL files.\n"
            f"Expected: {[str(p.name) for p in expected_files]}\n"
            f"Found: {[str(p.name) for p in actual_table_files]}"
        )

    staged_root = Path(
        context.powerbi_data_directory
    ).resolve()

    tmdl_validation = [
        validate_tmdl_file(
            path,
            registry.by_filename[path.name],
            staged_root,
        )
        for path in actual_table_files
    ]

    # There must be no unintended relationships.tmdl.
    relationship_file = semantic_definition / "relationships.tmdl"
    if relationship_file.exists():
        if model_meta.get("relationship_count", 0) <= 0:
            raise RuntimeError(
                "relationships.tmdl exists although no relationship "
                "is materialized."
            )
        text = relationship_file.read_text(
            encoding="utf-8"
        ).strip()
        if not re.search(
            r"^relationship\s+",
            text,
            flags=re.MULTILINE,
        ):
            raise RuntimeError(
                "relationships.tmdl has no relationship declarations."
            )

    # All staged files are current-run files and all TMDL references must point
    # only to them.
    for path in project_root.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() in {".pbip", ".pbir", ".pbism", ".tmdl", ".json"}:
            text = path.read_text(
                encoding="utf-8",
                errors="ignore",
            )
            if str(context.run_id) not in text and "File.Contents" in text:
                # TMDL is allowed to contain the current-run staged absolute path.
                # Any File.Contents source must already be checked by TMDL validation.
                pass

    page_validation = validate_page_and_visuals(
        project_root,
        report_definition,
        registry,
    )

    # Page contract validation.
    display_names = [
        item["display_name"]
        for item in page_validation["pages"]
    ]
    signal_supported = worker_contract_available(
        registry,
        "analysis_results",
        "Demand_Signal_Evidence",
        ("target_table", "target_measure_column", "signal_table",
         "signal_measure_column", "lag_periods", "direction",
         "relationship_strength", "aligned_observations"),
    )
    expected_names = [
        name
        for name, _ in PAGE_CONTRACTS
        if name != "Signal Intelligence" or signal_supported
    ]
    if display_names != expected_names:
        raise RuntimeError(
            "Generated page contract mismatch.\n"
            f"Expected: {expected_names}\n"
            f"Found: {display_names}"
        )

    # Current-run isolation.
    current_run_text_roots = (
        Path(context.powerbi_data_directory).resolve(),
        Path(context.output_directory).resolve(),
    )
    for table in registry.tables:
        if not _path_is_within(
            table.source_file,
            current_run_text_roots[0],
        ):
            raise RuntimeError(
                "Model table source is not staged under the current run "
                f"Power BI data directory: {table.source_file}"
            )

    return {
        "required_files_validated": len(required),
        "json_files_validated": len(json_files),
        "tmdl_files_present": len(actual_table_files),
        "table_definitions": len(registry.tables),
        "tmdl_validation": tmdl_validation,
        "materialized_relationships": int(
            model_meta.get("relationship_count", 0)
        ),
        "relationship_candidates": int(
            model_meta.get("relationship_candidates", 0)
        ),
        "page_validation": page_validation,
        "security": model_meta.get("security"),
        "desktop_validation": "NOT_PERFORMED",
        "status": "STRUCTURAL_AND_SEMANTIC_VALIDATION_PASSED",
    }


# ---------------------------------------------------------------------------
# Workspace/staging and integration
# ---------------------------------------------------------------------------


def _parse_recommendation_mapping(
    value: Any,
    field_name: str,
    row_number: int,
) -> dict[str, Any]:
    """Parse the serialized recommendation entity/scope mapping safely."""
    if value is None:
        return {}

    if isinstance(value, dict):
        parsed = value
    elif isinstance(value, str):
        candidate = value.strip()

        # Current recommendation artifacts contain quoted Python-literal
        # dictionaries. Peel redundant outer quoting before literal parsing.
        for _ in range(2):
            if (
                len(candidate) >= 2
                and candidate[0] == candidate[-1]
                and candidate[0] in {"'", '"'}
            ):
                candidate = candidate[1:-1].strip()

        try:
            parsed = ast.literal_eval(candidate)
        except (SyntaxError, ValueError) as exc:
            raise RuntimeError(
                f"Recommendation presentation normalization failed at "
                f"row {row_number}, field {field_name}: "
                f"unable to parse serialized mapping."
            ) from exc
    else:
        raise RuntimeError(
            f"Recommendation presentation normalization failed at "
            f"row {row_number}, field {field_name}: "
            f"expected mapping or serialized mapping, got "
            f"{type(value).__name__}."
        )

    if not isinstance(parsed, dict):
        raise RuntimeError(
            f"Recommendation presentation normalization failed at "
            f"row {row_number}, field {field_name}: parsed value is not a "
            f"dictionary."
        )

    normalized: dict[str, Any] = {}
    for key, item in parsed.items():
        key_text = str(key).strip()
        if not key_text:
            raise RuntimeError(
                f"Recommendation presentation normalization failed at "
                f"row {row_number}, field {field_name}: empty mapping key."
            )
        normalized[key_text] = item

    return normalized


def prepare_powerbi_presentation_data(
    context: Any,
    staged_files: list[Path],
) -> list[Path]:
    """Prepare Power BI-only presentation fields without modifying worker data.

    The worker-produced recommendation_results.xlsx remains immutable.
    Only the current-run staged copy is transformed for report presentation.

    Entity normalization is dataset-driven:
      * keys appearing in every recommendation are treated as universal
        context attributes;
      * keys appearing only on some rows are treated as entity attributes;
      * rows with no entity-specific attribute become Global;
      * all original recommendation columns and values are preserved.
    """
    if openpyxl is None:
        raise RuntimeError(
            "openpyxl is required for Power BI presentation preparation."
        )

    recommendation_files = [
        path
        for path in staged_files
        if path.name.lower() == "recommendation_results.xlsx"
    ]

    if len(recommendation_files) != 1:
        raise RuntimeError(
            "Expected exactly one staged recommendation_results.xlsx; "
            f"found {len(recommendation_files)}."
        )

    staged_recommendation = recommendation_files[0]
    worker_source = (
        Path(context.output_directory)
        / "recommendation_results.xlsx"
    )

    if not worker_source.is_file():
        raise FileNotFoundError(
            "Worker recommendation artifact is missing: "
            f"{worker_source}"
        )

    # The staged file must be a distinct copy. Never operate directly on the
    # worker artifact.
    if staged_recommendation.resolve() == worker_source.resolve():
        raise RuntimeError(
            "Presentation adapter refused to operate on the worker artifact."
        )

    workbook = openpyxl.load_workbook(
        staged_recommendation,
        read_only=False,
        data_only=False,
    )

    try:
        if "Recommendations" not in workbook.sheetnames:
            raise RuntimeError(
                "Recommendation presentation normalization requires a "
                "'Recommendations' worksheet."
            )

        sheet = workbook["Recommendations"]

        if sheet.max_row < 1 or sheet.max_column < 1:
            raise RuntimeError(
                "Recommendations worksheet is empty."
            )

        headers = [
            sheet.cell(row=1, column=column).value
            for column in range(1, sheet.max_column + 1)
        ]

        header_names = [
            str(value).strip()
            if value is not None
            else ""
            for value in headers
        ]

        if len(set(header_names)) != len(header_names):
            raise RuntimeError(
                "Recommendation presentation normalization found duplicate "
                "column headers."
            )

        required = {
            "recommendation_id",
            "entity",
            "scope",
        }

        missing = sorted(required - set(header_names))
        if missing:
            raise RuntimeError(
                "Recommendation presentation normalization cannot proceed; "
                f"missing required columns: {missing}"
            )

        entity_column = header_names.index("entity") + 1
        scope_column = header_names.index("scope") + 1

        row_records: list[dict[str, Any]] = []

        for row_number in range(2, sheet.max_row + 1):
            record = {
                header_names[column - 1]: sheet.cell(
                    row=row_number,
                    column=column,
                ).value
                for column in range(1, sheet.max_column + 1)
            }

            entity_mapping = _parse_recommendation_mapping(
                record.get("entity"),
                "entity",
                row_number,
            )

            # Validate scope too. It is part of the immutable worker contract
            # and its parseability is an important integrity invariant.
            _parse_recommendation_mapping(
                record.get("scope"),
                "scope",
                row_number,
            )

            row_records.append(
                {
                    "row_number": row_number,
                    "record": record,
                    "entity_mapping": entity_mapping,
                }
            )

        if not row_records:
            raise RuntimeError(
                "Recommendation presentation normalization found no "
                "recommendation rows."
            )

        # Dataset-driven identification of universal/context attributes.
        # A key present in every recommendation is context; keys present only
        # on a subset identify the entity represented by that row.
        key_frequency: dict[str, int] = {}

        for item in row_records:
            for key in item["entity_mapping"]:
                key_frequency[key] = key_frequency.get(key, 0) + 1

        row_count = len(row_records)

        evidence_column = header_names.index("evidence") + 1

        def parse_structured_evidence(raw_value: Any) -> dict[str, Any] | None:
            """Parse serialized worker evidence for Power BI presentation only."""
            value: Any = raw_value

            for _ in range(4):
                if isinstance(value, dict):
                    return value

                if isinstance(value, (bytes, bytearray)):
                    try:
                        value = value.decode("utf-8")
                    except UnicodeDecodeError:
                        return None

                if not isinstance(value, str):
                    return None

                candidate = value.strip()

                if not candidate:
                    return None

                decoded: Any = None

                try:
                    decoded = json.loads(candidate)
                except (TypeError, ValueError, json.JSONDecodeError):
                    try:
                        decoded = ast.literal_eval(candidate)
                    except (ValueError, SyntaxError, TypeError, MemoryError):
                        return None

                if decoded == value:
                    return None

                value = decoded

            return value if isinstance(value, dict) else None

        for item in row_records:
            raw_evidence = sheet.cell(
                row=item["row_number"],
                column=evidence_column,
            ).value

            summary_parts: list[str] = []
            parsed_evidence = parse_structured_evidence(raw_evidence)

            if isinstance(parsed_evidence, dict):

                evidence_mode = parsed_evidence.get("evidence_mode")

                if evidence_mode:
                    summary_parts.append(
                        str(evidence_mode).replace("_", " ").title()
                    )

                supporting_signals = parsed_evidence.get(
                    "supporting_signals"
                )

                if isinstance(supporting_signals, list):
                    count = len(supporting_signals)

                    summary_parts.append(
                        f"{count} supporting signal"
                        f"{'' if count == 1 else 's'}"
                    )

                forecast = parsed_evidence.get("forecast")

                if isinstance(forecast, dict):
                    confidence = forecast.get("confidence")

                    if confidence not in (None, ""):
                        summary_parts.append(
                            f"Forecast confidence: {confidence}"
                        )

                candidate_score = parsed_evidence.get("candidate_score")

                if isinstance(candidate_score, (int, float)):
                    summary_parts.append(
                        f"Candidate score: {candidate_score:.2f}"
                    )

            if not summary_parts:
                summary_parts.append(
                    "No evidence published"
                    if raw_evidence in (None, "")
                    else "Evidence available"
                )

            item["evidence_summary"] = " • ".join(summary_parts)
        universal_keys = {
            key
            for key, frequency in key_frequency.items()
            if frequency == row_count
        }

        entity_keys = set(key_frequency) - universal_keys

        presentation_headers = [
            "entity_attribute",
            "entity_value",
            "entity_context",
            "evidence_summary",
        ]

        for header in presentation_headers:
            if header in header_names:
                raise RuntimeError(
                    "Recommendation presentation normalization found an "
                    f"existing reserved presentation column: {header}"
                )

        start_column = sheet.max_column + 1

        for offset, header in enumerate(
            presentation_headers,
            start=0,
        ):
            sheet.cell(
                row=1,
                column=start_column + offset,
                value=header,
            )

        normalized_count = 0
        global_count = 0

        for item in row_records:
            mapping = item["entity_mapping"]

            specific_items = [
                (key, value)
                for key, value in mapping.items()
                if key in entity_keys
            ]

            if len(specific_items) == 0:
                entity_attribute = "Global"
                entity_value = "Global"
                global_count += 1
            elif len(specific_items) == 1:
                entity_attribute, raw_value = specific_items[0]
                entity_value = (
                    str(raw_value)
                    if raw_value is not None
                    else ""
                )
                normalized_count += 1
            else:
                # Preserve every non-universal entity attribute without
                # choosing an arbitrary "first" key.
                specific_items.sort(key=lambda pair: pair[0])
                entity_attribute = "Multi-Entity"
                entity_value = " | ".join(
                    f"{key}={value}"
                    for key, value in specific_items
                )
                normalized_count += 1

            context_items = [
                (key, value)
                for key, value in mapping.items()
                if key in universal_keys
            ]
            context_items.sort(key=lambda pair: pair[0])

            entity_context = " | ".join(
                f"{key}={value}"
                for key, value in context_items
            )

            row_number = item["row_number"]

            sheet.cell(
                row=row_number,
                column=start_column,
                value=entity_attribute,
            )
            sheet.cell(
                row=row_number,
                column=start_column + 1,
                value=entity_value,
            )
            sheet.cell(
                row=row_number,
                column=start_column + 2,
                value=entity_context,
            )
            sheet.cell(
                row=row_number,
                column=start_column + 3,
                value=item["evidence_summary"],
            )

        # Final structural checks before saving.
        expected_max_column = len(header_names) + len(
            presentation_headers
        )

        if sheet.max_column != expected_max_column:
            raise RuntimeError(
                "Recommendation presentation normalization produced an "
                "unexpected column count: "
                f"{sheet.max_column} != {expected_max_column}"
            )

        if sheet.max_row != row_count + 1:
            raise RuntimeError(
                "Recommendation presentation normalization changed row "
                f"count: {sheet.max_row} != {row_count + 1}"
            )

        if normalized_count + global_count != row_count:
            raise RuntimeError(
                "Recommendation presentation normalization row accounting "
                "failed."
            )

        # Explicitly verify that the original worker columns remain intact.
        for column_index, expected_header in enumerate(
            header_names,
            start=1,
        ):
            actual_header = sheet.cell(
                row=1,
                column=column_index,
            ).value

            if actual_header != expected_header:
                raise RuntimeError(
                    "Recommendation presentation normalization altered an "
                    f"original header at column {column_index}: "
                    f"{actual_header!r} != {expected_header!r}"
                )

        workbook.save(staged_recommendation)

    finally:
        workbook.close()

    # Confirm the worker artifact itself was not modified during the adapter.
    # This compares the worker file before/after only within the current call,
    # so no source mutation can silently pass.
    return staged_files


def get_powerbi_source_files(context: Any) -> list[Path]:
    required = [
        context.output_directory / "analytical_dataset.xlsx",
        context.output_directory / "analysis_results.xlsx",
        context.output_directory / "forecast_results.xlsx",
        context.output_directory / "recommendation_results.xlsx",
    ]
    missing = [
        str(path)
        for path in required
        if not path.is_file() or path.stat().st_size == 0
    ]
    if missing:
        raise FileNotFoundError(
            f"Power BI current-run source artifact(s) missing: {missing}"
        )

    # Explicitly reject artifacts outside the current run.
    run_root = Path(context.output_directory).resolve()
    for source in required:
        if not _path_is_within(source, run_root):
            raise RuntimeError(
                f"Power BI source escapes current-run output directory: {source}"
            )
    return required


def stage_data_files(
    context: Any,
    source_files: list[Path],
) -> list[Path]:
    destination = Path(
        context.powerbi_data_directory
    )
    destination.mkdir(
        parents=True,
        exist_ok=True,
    )

    copied: list[Path] = []
    for source in source_files:
        target = destination / source.name
        shutil.copy2(source, target)
        if not target.is_file() or target.stat().st_size == 0:
            raise RuntimeError(
                f"Power BI staging failed for {source}"
            )
        copied.append(target)

    return copied


def write_documentation(
    context: Any,
    project: Path,
    metadata: dict[str, Any],
) -> Path:
    documentation = Path(
        context.powerbi_documentation_directory
    )
    documentation.mkdir(
        parents=True,
        exist_ok=True,
    )
    path = documentation / "README.txt"

    lines = [
        "Enterprise Demand Intelligence Platform â€” Power BI Run",
        "",
        f"Run ID: {context.run_id}",
        f"Generated UTC: {utc_now_iso()}",
        f"Project: {project}",
        "",
        "Report contract:",
        "  Power BI is the final BI/decision-support presentation layer.",
        "  Analytical pages are capability-driven and use worker-produced evidence.",
        "  Record Explorer is the only broad row-level inspection surface.",
        "",
        "Relationship policy:",
        "  Candidates are documented only; materialization requires a complete",
        "  uniqueness/cardinality/coverage/direction/ambiguity safety test.",
        "",
        "Security policy:",
        f"  {metadata['semantic_model']['security']['status']}",
        "  Operational RLS is not claimed without operational identity mapping.",
        "",
        "Source-document alignment:",
        "  Demand-sensing questions are represented only where the current-run",
        "  worker evidence contracts support them. The source document's signal",
        "  categories are contextual; unsupported categories are never fabricated.",
        "",
        "Desktop gate:",
        "  NOT_PERFORMED by generator. Open, refresh and save in the installed",
        "  Power BI Desktop version before claiming final Desktop compatibility.",
    ]
    path.write_text(
        "\n".join(lines) + "\n",
        encoding="utf-8",
    )
    return path


def generate_powerbi(
    context: Any,
    log_fn: Callable[[str], None] | None = None,
) -> tuple[bool, dict[str, Any], str | None]:
    emit = log_fn or (lambda _: None)
    started = time.perf_counter()
    project: Path | None = None

    try:
        step = time.perf_counter()
        source_files = get_powerbi_source_files(context)
        emit(
            "POWERBI STEP source discovery: "
            f"{time.perf_counter() - step:.3f}s"
        )

        step = time.perf_counter()
        staged_files = stage_data_files(
            context,
            source_files,
        )
        staged_files = prepare_powerbi_presentation_data(
            context,
            staged_files,
        )
        registry = TableDefinitionRegistry.discover(
            staged_files
        )
        emit(
            "POWERBI STEP data staging/schema: "
            f"{time.perf_counter() - step:.3f}s"
        )

        # Canonical locked workspace/project location.
        project = (
            Path(context.powerbi_root)
            / f"Demand_Intelligence_Report_{context.run_id}"
        )
        if project.exists():
            raise FileExistsError(
                f"Power BI project already exists; refusing to overwrite: {project}"
            )
        project.mkdir(
            parents=True,
            exist_ok=False,
        )

        model = generate_model(
            project,
            registry,
        )
        emit(
            "POWERBI STEP semantic model: "
            f"{time.perf_counter() - step:.3f}s"
        )

        step = time.perf_counter()
        report = generate_report(
            project,
            model,
            registry,
        )
        emit(
            "POWERBI STEP report generation: "
            f"{time.perf_counter() - step:.3f}s"
        )

        step = time.perf_counter()
        validation = validate_powerbi_project(
            project,
            model,
            registry,
            context,
        )
        emit(
            "POWERBI STEP validation: "
            f"{time.perf_counter() - step:.3f}s"
        )

        total = time.perf_counter() - started

        metadata: dict[str, Any] = {
            "status": "GENERATED_STRUCTURAL_AND_SEMANTIC_VALIDATION",
            "run_id": context.run_id,
            "workspace": str(
                Path(context.powerbi_root).resolve()
            ),
            "project": str(project.resolve()),
            "pbip": str(
                (
                    project
                    / f"{project.name}.pbip"
                ).resolve()
            ),
            "data_directory": str(
                Path(
                    context.powerbi_data_directory
                ).resolve()
            ),
            "documentation_directory": str(
                Path(
                    context.powerbi_documentation_directory
                ).resolve()
            ),
            "source_files": [
                str(path.resolve())
                for path in source_files
            ],
            "staged_data_files": [
                str(path.resolve())
                for path in staged_files
            ],
            "semantic_model": model,
            "report": report,
            "validation": validation,
            "timings_seconds": {
                "total": total,
            },
        }

        documentation = write_documentation(
            context,
            project,
            metadata,
        )
        metadata["documentation"] = str(
            documentation.resolve()
        )

        json_write(
            project / "integration_manifest.json",
            metadata,
        )

        emit(
            f"Power BI project: CREATED â€” {project}"
        )
        emit(
            f"Power BI pages: {report['pages']}; "
            f"visuals: {report['visuals']}"
        )
        emit(
            "Power BI relationships: "
            f"{validation['materialized_relationships']} materialized; "
            f"{validation['relationship_candidates']} candidates documented"
        )
        emit(
            "Security status: "
            f"{model['security']['status']}"
        )
        emit(
            f"POWERBI TOTAL: {total:.3f}s"
        )

        return True, metadata, None

    except Exception as exc:
        elapsed = time.perf_counter() - started
        emit(
            f"POWERBI FAILED after {elapsed:.3f}s"
        )

        # Failed/partial projects must never be treated as valid outputs.
        # Keep the directory only when explicitly requested for debugging.
        if project is not None and project.exists():
            preserve = (
                os.getenv(
                    "EDI_PRESERVE_FAILED_POWERBI_PROJECT",
                    "",
                ).strip().lower()
                in {"1", "true", "yes"}
            )
            if not preserve:
                try:
                    shutil.rmtree(project)
                    emit(
                        "POWERBI partial project removed after validation/generation failure."
                    )
                except Exception as cleanup_exc:
                    emit(
                        "POWERBI partial-project cleanup failed: "
                        f"{type(cleanup_exc).__name__}: {cleanup_exc}"
                    )

        return (
            False,
            {},
            f"{type(exc).__name__}: {exc}",
        )


# ---------------------------------------------------------------------------
# Optional Desktop launch
# ---------------------------------------------------------------------------

def discover_powerbi_desktop() -> Path | None:
    candidates: list[Path] = []

    program_files = os.getenv("ProgramFiles")
    local_app_data = os.getenv("LocalAppData")

    if program_files:
        candidates.append(
            Path(program_files)
            / "Microsoft Power BI Desktop"
            / "bin"
            / "PBIDesktop.exe"
        )

    if local_app_data:
        candidates.append(
            Path(local_app_data)
            / "Programs"
            / "Microsoft Power BI Desktop"
            / "bin"
            / "PBIDesktop.exe"
        )

    return next(
        (
            path
            for path in candidates
            if path.is_file()
        ),
        None,
    )


def maybe_open_powerbi_project(
    context: Any,
    project_path: Path,
    log_fn: Callable[[str], None] | None = None,
) -> None:
    emit = log_fn or (lambda _: None)
    desktop = discover_powerbi_desktop()

    if desktop is None:
        emit(
            "Power BI Desktop was not discovered; "
            "project remains on disk."
        )
        return

    try:
        subprocess.Popen(
            [
                str(desktop),
                str(project_path),
            ]
        )
        emit(
            f"Power BI Desktop launched: {desktop}"
        )
    except Exception as exc:
        emit(
            "Power BI Desktop launch skipped: "
            f"{type(exc).__name__}: {exc}"
        )






