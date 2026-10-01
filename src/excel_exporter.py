"""
Enterprise Demand Intelligence Platform
GENERIC SOURCE TO EXCEL EXPORTER

Purpose
-------
Export the currently discovered dataset structure into an Excel workbook
without requiring a particular file format, table name, column name, value,
entity count, or business domain.

The exporter is intentionally read-only with respect to the source dataset.
It performs only the structural conversion required to represent discovered
tables as Excel worksheets.

Architecture guarantees
-----------------------
- Dataset path is resolved by the generic loader.
- Supported source formats are controlled only by configuration.
- Tables are discovered from the actual source contents.
- Columns are copied from the discovered source tables.
- No business aliases or semantic column lists are used.
- No cleaning, imputation, aggregation, filtering, or business transformation
  is performed.
- The source dataset is never modified.
- Excel worksheet names are generated only to satisfy Excel's technical
  worksheet-name constraints.
- Worksheet-name collisions are handled deterministically.
- Tables larger than Excel's technical row limit are split across worksheets
  while preserving all rows and columns.
"""

from __future__ import annotations

import json
import re
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
# CONFIGURATION
# ============================================================

import config

from src.data_loader import discover_dataset


# ============================================================
# TECHNICAL EXCEL LIMITS
# ============================================================

EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
EXCEL_MAX_SHEET_NAME_LENGTH = 31


# ============================================================
# GENERIC HELPERS
# ============================================================

def json_safe(value: Any) -> Any:
    """Convert values into JSON-safe representations."""

    if value is None:
        return None

    if isinstance(value, dict):
        return {
            str(key): json_safe(item)
            for key, item in value.items()
        }

    if isinstance(value, (list, tuple, set)):
        return [
            json_safe(item)
            for item in value
        ]

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except Exception:
            pass

    if hasattr(value, "item"):
        try:
            return json_safe(value.item())
        except Exception:
            pass

    if isinstance(value, float):
        if not pd.notna(value):
            return None

    if isinstance(value, (str, int, float, bool)):
        return value

    return str(value)


def make_safe_sheet_name(name: str) -> str:
    """
    Convert a discovered table name into a technically valid Excel
    worksheet name.

    This function does not translate or semantically reinterpret the name.
    """

    value = str(name)

    value = re.sub(
        r"[:\\/?*\[\]]",
        "_",
        value,
    )

    value = value.strip() or "table"

    return value[:EXCEL_MAX_SHEET_NAME_LENGTH]


def make_unique_sheet_name(
    original_name: str,
    used_names: set[str],
) -> str:
    """
    Create a deterministic, collision-free Excel worksheet name.
    """

    base = make_safe_sheet_name(original_name)

    if base not in used_names:
        used_names.add(base)
        return base

    counter = 2

    while True:
        suffix = f"_{counter}"
        available_length = (
            EXCEL_MAX_SHEET_NAME_LENGTH
            - len(suffix)
        )

        candidate = (
            base[:available_length]
            + suffix
        )

        if candidate not in used_names:
            used_names.add(candidate)
            return candidate

        counter += 1


def split_for_excel(
    dataframe: pd.DataFrame,
) -> list[pd.DataFrame]:
    """
    Split a DataFrame only when required by Excel's technical row limit.

    No rows or columns are filtered.
    """

    if len(dataframe) <= EXCEL_MAX_ROWS - 1:
        return [dataframe]

    chunk_size = EXCEL_MAX_ROWS - 1

    return [
        dataframe.iloc[start:start + chunk_size]
        for start in range(
            0,
            len(dataframe),
            chunk_size,
        )
    ]


# ============================================================
# TABLE EXPORT
# ============================================================

def export_tables_to_excel(
    tables: dict[str, pd.DataFrame],
    output_path: Path,
) -> dict[str, Any]:
    """
    Export every discovered table to Excel.

    Returns a structural export report.
    """

    if not tables:
        raise ValueError(
            "No tables were discovered in the supplied dataset."
        )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    used_sheet_names: set[str] = set()

    table_reports: list[dict[str, Any]] = []

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:

        for table_name, dataframe in tables.items():

            if len(dataframe.columns) > EXCEL_MAX_COLUMNS:
                raise ValueError(
                    f"Table '{table_name}' contains "
                    f"{len(dataframe.columns):,} columns, exceeding "
                    f"Excel's technical column limit of "
                    f"{EXCEL_MAX_COLUMNS:,}."
                )

            chunks = split_for_excel(dataframe)

            written_sheet_names: list[str] = []

            for chunk_index, chunk in enumerate(chunks):

                if len(chunks) == 1:
                    requested_name = str(table_name)
                else:
                    requested_name = (
                        f"{table_name}_{chunk_index + 1}"
                    )

                sheet_name = make_unique_sheet_name(
                    requested_name,
                    used_sheet_names,
                )

                chunk.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False,
                )

                written_sheet_names.append(sheet_name)

                print(
                    f"  Exported: "
                    f"{str(table_name):<32} "
                    f"{len(chunk):>10,} rows "
                    f"{len(chunk.columns):>6,} columns "
                    f"-> {sheet_name}"
                )

            table_reports.append(
                {
                    "source_table": str(table_name),
                    "source_rows": int(len(dataframe)),
                    "source_columns": int(
                        len(dataframe.columns)
                    ),
                    "worksheet_count": len(
                        written_sheet_names
                    ),
                    "worksheets": written_sheet_names,
                }
            )

    return {
        "table_count": len(tables),
        "total_source_rows": int(
            sum(len(frame) for frame in tables.values())
        ),
        "total_source_columns": int(
            sum(len(frame.columns) for frame in tables.values())
        ),
        "tables": table_reports,
    }


# ============================================================
# EXPORT VALIDATION
# ============================================================

def validate_export(
    output_path: Path,
    tables: dict[str, pd.DataFrame],
    export_report: dict[str, Any],
) -> bool:
    """
    Validate workbook existence, worksheet presence, column preservation,
    and row-count preservation.

    Validation reads the generated workbook only. It never changes the
    source dataset.
    """

    if not output_path.exists():
        print(
            "Validation failed: workbook was not created."
        )
        return False

    try:
        workbook = pd.ExcelFile(
            output_path,
            engine="openpyxl",
        )
    except Exception as exc:
        print(
            "Validation failed while opening workbook: "
            f"{exc}"
        )
        return False

    actual_sheets = workbook.sheet_names

    expected_sheets = [
        sheet
        for table in export_report["tables"]
        for sheet in table["worksheets"]
    ]

    if set(actual_sheets) != set(expected_sheets):
        print(
            "Validation failed: worksheet set does not match "
            "the generated export report."
        )
        return False

    # --------------------------------------------------------
    # Validate every source table against its generated chunks.
    # --------------------------------------------------------

    for table in export_report["tables"]:

        source_name = table["source_table"]
        source_frame = tables[source_name]

        reconstructed_rows = 0
        reconstructed_columns: list[str] | None = None

        for sheet_name in table["worksheets"]:

            exported = pd.read_excel(
                workbook,
                sheet_name=sheet_name,
            )

            reconstructed_rows += len(exported)

            current_columns = [
                str(column)
                for column in exported.columns
            ]

            if reconstructed_columns is None:
                reconstructed_columns = current_columns
            elif reconstructed_columns != current_columns:
                print(
                    f"Validation failed: column structure differs "
                    f"across worksheets for '{source_name}'."
                )
                return False

        source_columns = [
            str(column)
            for column in source_frame.columns
        ]

        if reconstructed_columns != source_columns:
            print(
                f"Validation failed: columns changed for "
                f"'{source_name}'."
            )
            return False

        if reconstructed_rows != len(source_frame):
            print(
                f"Validation failed: row count changed for "
                f"'{source_name}' "
                f"({len(source_frame):,} -> "
                f"{reconstructed_rows:,})."
            )
            return False

    return True


# ============================================================
# REPORT WRITING
# ============================================================

def write_export_report(
    output_path: Path,
    discovery: dict[str, Any],
    export_report: dict[str, Any],
) -> Path:
    """Write a generic structural export report."""

    report_path = (
        output_path.parent
        / "excel_export_report.json"
    )

    payload = {
        "export": {
            "output_path": str(output_path),
            "format": "xlsx",
            "source_modified": False,
            "business_semantics_applied": False,
            "business_aliases_used": False,
            "transformations_applied": [],
        },
        "source": {
            "path": discovery["dataset"]["path"],
            "format": discovery["dataset"]["format"],
            "size_bytes": discovery["dataset"]["size_bytes"],
            "size_mb": discovery["dataset"]["size_mb"],
        },
        "summary": {
            "discovered_tables": discovery["summary"]["table_count"],
            "source_rows": discovery["summary"]["total_rows"],
            "exported_tables": export_report["table_count"],
            "exported_rows": export_report["total_source_rows"],
        },
        "tables": export_report["tables"],
    }

    with report_path.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            json_safe(payload),
            handle,
            indent=2,
            ensure_ascii=False,
        )

    return report_path


# ============================================================
# MAIN PIPELINE
# ============================================================

def run_exporter(
    dataset_path: str | Path | None = None,
) -> None:

    print("=" * 72)
    print(
        "ENTERPRISE DEMAND INTELLIGENCE PLATFORM"
    )
    print(
        "GENERIC SOURCE TO EXCEL EXPORTER"
    )
    print("=" * 72)

    print()
    print(
        "Architecture mode: DATASET-DRIVEN"
    )
    print(
        "Semantic mode: STRUCTURAL — NO BUSINESS NAME ASSUMPTIONS"
    )
    print(
        "Source modification: DISABLED"
    )
    print(
        "Business schema required: NO"
    )

    # --------------------------------------------------------
    # STEP 1 — DISCOVER SOURCE DATASET
    # --------------------------------------------------------

    print()
    print(
        "STEP 1 — DISCOVERING SOURCE DATASET"
    )

    discovery = discover_dataset(
        dataset_path=dataset_path,
    )

    tables = discovery["tables"]

    print(
        f"Dataset: {discovery['dataset']['path']}"
    )
    print(
        f"Format : {discovery['dataset']['format']}"
    )
    print(
        f"Tables : {len(tables)}"
    )
    print(
        f"Rows   : {discovery['summary']['total_rows']:,}"
    )

    # --------------------------------------------------------
    # STEP 2 — EXPORT
    # --------------------------------------------------------

    configured_output = getattr(
        config,
        "EXCEL_EXPORT_OUTPUT_PATH",
        None,
    )

    if configured_output:
        output_path = Path(configured_output).expanduser().resolve()
    else:
        output_path = (
            Path(config.OUTPUT_DATA_DIRECTORY)
            / "raw_dataset.xlsx"
        )

    print()
    print(
        "STEP 2 — EXPORTING DISCOVERED TABLES"
    )

    export_report = export_tables_to_excel(
        tables=tables,
        output_path=output_path,
    )

    # --------------------------------------------------------
    # STEP 3 — VALIDATE
    # --------------------------------------------------------

    print()
    print(
        "STEP 3 — VALIDATING EXCEL EXPORT"
    )

    valid = validate_export(
        output_path=output_path,
        tables=tables,
        export_report=export_report,
    )

    if not valid:
        raise RuntimeError(
            "Excel export validation failed."
        )

    print(
        "Excel structural validation: PASSED."
    )

    # --------------------------------------------------------
    # STEP 4 — REPORT
    # --------------------------------------------------------

    report_path = write_export_report(
        output_path=output_path,
        discovery=discovery,
        export_report=export_report,
    )

    # --------------------------------------------------------
    # FINAL
    # --------------------------------------------------------

    print()
    print("=" * 72)
    print(
        "GENERIC EXCEL EXPORTER COMPLETED SUCCESSFULLY"
    )
    print("=" * 72)

    print()
    print(
        f"Workbook: {output_path}"
    )
    print(
        f"Report   : {report_path}"
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
        "  Business transformations: NO"
    )
    print(
        "  Source data modified: NO"
    )


if __name__ == "__main__":

    try:
        run_exporter()
    except Exception as error:
        print()
        print("=" * 72)
        print(
            "GENERIC EXCEL EXPORTER FAILED"
        )
        print("=" * 72)
        print()
        print(f"ERROR: {error}")
        raise
