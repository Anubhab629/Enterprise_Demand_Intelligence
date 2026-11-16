"""Enterprise Demand Intelligence — authenticated, read-only Streamlit dashboard."""

from __future__ import annotations

import json
from collections.abc import Mapping
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from independent_intelligence_app import render_independent_intelligence

from streamlit_app.authentication import verify_credentials
from streamlit_app.analytics import (
    boolean_columns,
    categorical_columns,
    dataset_metrics,
    datetime_columns,
    expand_nested_columns,
    numeric_columns,
    numeric_measure_columns,
)
from streamlit_app.data_service import (
    PIPELINE_STAGES,
    get_artifact_catalog,
    get_artifact_path,
    get_latest_successful_run,
    get_run_history,
    get_successful_runs,
    load_original_dataset,
    load_run_json,
    load_run_tables,
    load_run_workbook,
)
from streamlit_app.visualizations import (
    render_boolean_comparisons,
    render_category_comparisons,
    render_geographic_charts,
    render_missing_value_charts,
    render_numeric_distributions,
    render_relationship_charts,
    render_time_series_charts,
)

st.set_page_config(
    page_title="Enterprise Demand Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


@st.cache_resource(show_spinner=False)
def register_application_launch() -> int:
    """Increment the persistent counter once per Streamlit server process."""
    counter_path = (
        Path(__file__).resolve().parent.parent
        / "data"
        / "application_views.json"
    )
    counter_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        stored_data = json.loads(counter_path.read_text(encoding="utf-8"))
        previous_count = int(stored_data.get("views", 0))
    except (OSError, json.JSONDecodeError, TypeError, ValueError, AttributeError):
        previous_count = 0

    current_count = max(0, previous_count) + 1

    temporary_path = counter_path.with_name(
        f"{counter_path.name}.tmp"
    )
    temporary_path.write_text(
        json.dumps({"views": current_count}, indent=2),
        encoding="utf-8",
    )
    temporary_path.replace(counter_path)

    return current_count


application_views = register_application_launch()


st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.4rem;
        padding-bottom: 2.5rem;
        max-width: 1800px;
    }
    [data-testid="stMetric"] {
        border: 1px solid rgba(128, 128, 128, 0.25);
        border-radius: 12px;
        padding: 15px 16px;
        min-height: 105px;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.86rem;
    }
    .edi-subtitle {
        color: #8993a4;
        margin-top: -0.55rem;
        margin-bottom: 1.4rem;
    }
    .edi-section {
        margin-top: 1.1rem;
        margin-bottom: 0.45rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------



def show_login() -> None:
    """Display the EDI dashboard login page."""

    st.markdown(
        """
        <div style="text-align: center; margin-bottom: 35px;">
            <h1 style="color: #17365D;">
                Enterprise Demand Intelligence
            </h1>
            <p style="color: #64748b; font-size: 17px;">
                Enter the username and password you set in main.py
                to access the dashboard.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _, center, _ = st.columns([1, 1.8, 1])

    with center:
        with st.container(border=True):
            st.subheader("Dashboard Login")

            with st.form("edi_login_form", clear_on_submit=False):
                username = st.text_input("Username")
                password = st.text_input(
                    "Password",
                    type="password",
                )

                submitted = st.form_submit_button(
                    "Login",
                    type="primary",
                    use_container_width=True,
                )

            if submitted:
                if verify_credentials(username, password):
                    st.session_state["edi_authenticated"] = True
                    st.rerun()
                else:
                    st.error(
                        "Invalid username or password. Please try again."
                    )


if not st.session_state.get("edi_authenticated", False):
    show_login()
    st.stop()


# ---------------------------------------------------------------------------
# General-purpose data helpers
# ---------------------------------------------------------------------------

def nonempty_frames(value: Any) -> dict[str, pd.DataFrame]:
    """Normalize a table mapping and remove empty/non-DataFrame entries."""
    if isinstance(value, pd.DataFrame):
        return {"dataset": value} if not value.empty else {}

    if not isinstance(value, Mapping):
        return {}

    result: dict[str, pd.DataFrame] = {}
    for name, frame in value.items():
        if isinstance(frame, pd.DataFrame) and not frame.empty:
            result[str(name)] = frame.copy()

    return result


def json_to_frames(value: Any, name: str = "results") -> dict[str, pd.DataFrame]:
    """Convert JSON structures into readable tables without assuming a schema."""
    if isinstance(value, pd.DataFrame):
        return {name: value} if not value.empty else {}

    if isinstance(value, Mapping):
        frames: dict[str, pd.DataFrame] = {}

        for key, child in value.items():
            if isinstance(child, pd.DataFrame):
                if not child.empty:
                    frames[str(key)] = child
            elif isinstance(child, list):
                if child and all(isinstance(item, Mapping) for item in child):
                    frames[str(key)] = pd.json_normalize(child, sep=".")
                elif child and not any(
                    isinstance(item, (list, Mapping)) for item in child
                ):
                    frames[str(key)] = pd.DataFrame({str(key): child})
            elif isinstance(child, Mapping):
                nested = json_to_frames(child, str(key))
                frames.update(nested)

        if frames:
            return frames

        try:
            frame = pd.json_normalize(value, sep=".")
            return {name: frame} if not frame.empty else {}
        except (TypeError, ValueError):
            return {}

    if isinstance(value, list):
        if not value:
            return {}
        if all(isinstance(item, Mapping) for item in value):
            frame = pd.json_normalize(value, sep=".")
            return {name: frame} if not frame.empty else {}
        if not any(isinstance(item, (list, Mapping)) for item in value):
            return {name: pd.DataFrame({name: value})}

    return {}


def prepare_frame(frame: pd.DataFrame) -> pd.DataFrame:
    """Prepare a readable display copy without uncontrolled column expansion.

    Source rows and source columns are retained. Nested values are expanded only
    when the resulting width stays within a resource guard; otherwise nested
    values are serialized into their original columns.
    """
    if not isinstance(frame, pd.DataFrame):
        return pd.DataFrame()
    if frame.empty:
        return frame.copy(deep=True)

    original = frame.copy(deep=True)
    # Bound the display expansion relative to the actual source schema, not to
    # any particular workbook, table, or business field.
    max_display_columns = max(30, min(120, original.shape[1] * 3))
    try:
        expanded = expand_nested_columns(
            original,
            max_depth=3,
            separator=".",
            parse_json_strings=True,
        )
        if isinstance(expanded, pd.DataFrame) and expanded.shape[1] <= max_display_columns:
            if expanded.columns.duplicated().any():
                expanded = expanded.loc[:, ~expanded.columns.duplicated()].copy()
            if len(expanded) == len(original):
                return expanded
    except Exception:
        # Fall back to a readable representation rather than failing the page.
        pass

    for column in original.columns:
        original[column] = original[column].map(
            lambda value: json.dumps(value, ensure_ascii=False, default=str, sort_keys=True)
            if isinstance(value, (dict, list, tuple, set))
            else value
        )
    return original


def safe_label(value: Any) -> str:
    return str(value).replace("_", " ").strip().title()


def format_integer(value: int | float) -> str:
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError, OverflowError):
        return "—"


def _table_utility_score(frame: pd.DataFrame) -> float:
    """Rank tables by structural usefulness without relying on table names."""
    if not isinstance(frame, pd.DataFrame) or frame.empty or frame.shape[1] == 0:
        return float("-inf")
    rows, columns = frame.shape
    nonempty_columns = int(frame.notna().any(axis=0).sum())
    numeric_count = len(numeric_measure_columns(frame))
    date_count = len(datetime_columns(frame))
    category_count = len(categorical_columns(frame))
    boolean_count = len(boolean_columns(frame))
    missing_ratio = float(frame.isna().to_numpy().mean()) if rows and columns else 1.0
    # Log scaling avoids always selecting the largest table while rewarding
    # enough rows, a useful schema, mixed analytical types, and completeness.
    import math
    return (
        math.log1p(rows) * 1.2
        + math.log1p(nonempty_columns) * 1.5
        + min(numeric_count, 12) * 0.8
        + min(date_count, 4) * 0.7
        + min(category_count, 12) * 0.35
        + min(boolean_count, 4) * 0.2
        - missing_ratio * 2.0
    )


def selected_frame(
    frames: Mapping[str, pd.DataFrame],
    widget_key: str,
    label: str = "Dataset",
    preferred_names: tuple[str, ...] = (),
) -> tuple[str, pd.DataFrame] | tuple[None, None]:
    """Select a table; choose the structurally most useful table by default.

    ``preferred_names`` remains accepted for compatibility but is deliberately
    ignored so a dataset's naming convention cannot determine selection.
    """
    if not frames:
        return None, None

    names = list(frames.keys())
    if len(names) == 1:
        name = names[0]
        return name, prepare_frame(frames[name])

    default_name = max(names, key=lambda name: _table_utility_score(frames[name]))
    default_index = names.index(default_name)
    chosen = st.selectbox(
        label,
        names,
        index=default_index,
        format_func=safe_label,
        key=widget_key,
    )
    return chosen, prepare_frame(frames[chosen])


def show_kpis(frame: pd.DataFrame, prefix: str = "") -> None:
    """Display metrics calculated from the actual selected frame."""
    metrics = dataset_metrics({"selected": frame})
    numeric_count = len(numeric_columns(frame))
    date_count = len(datetime_columns(frame))
    category_count = len(categorical_columns(frame))
    bool_count = len(boolean_columns(frame))

    cells = max(int(frame.shape[0]) * int(frame.shape[1]), 0)
    missing = int(frame.isna().sum().sum())
    completeness = ((cells - missing) / cells * 100) if cells else 0.0

    cards = [
        ("Records", format_integer(len(frame))),
        ("Columns", format_integer(len(frame.columns))),
        ("Completeness", f"{completeness:.1f}%"),
        ("Numeric fields", format_integer(numeric_count)),
        ("Date fields", format_integer(date_count)),
    ]

    st.markdown(
        f"#### {prefix}Dataset snapshot" if prefix else "#### Dataset snapshot"
    )
    columns = st.columns(len(cards))
    for column, (label, value) in zip(columns, cards):
        with column:
            st.metric(label, value)

    with st.expander("Additional schema details", expanded=False):
        details = {
            "Categorical fields": category_count,
            "Boolean fields": bool_count,
            "Missing cells": missing,
            "Duplicate rows": int(frame.duplicated().sum()),
            "Data-service table count": metrics.get("tables", 1),
        }
        detail_columns = st.columns(len(details))
        for column, (label, value) in zip(detail_columns, details.items()):
            with column:
                st.metric(label, format_integer(value))


def style_record_rows(frame: pd.DataFrame) -> "pd.io.formats.style.Styler":
    """Apply presentation-only record-group colours."""
    palette = [
        "#E8F1FF", "#E8F7EF", "#FFF2DF", "#F3EAFE",
        "#FCE8EC", "#E3F6F7", "#F4F0E6", "#EAF0F5",
        "#F0F4D8", "#FBE9D7", "#E9E7FF", "#DDF3E4",
    ]
    semantic_names = (
        "record_type", "recommendation_type", "priority_level", "status",
        "order_status", "category", "table", "segment", "region", "group",
        "classification", "trend", "confidence", "stage",
    )

    group_column = None
    for column in frame.columns:
        name = str(column).strip().lower().replace(" ", "_").replace("-", "_")
        values = frame[column].dropna()
        if name in semantic_names and 2 <= values.nunique() <= len(palette):
            group_column = column
            break

    # Avoid failures with string, duplicate, or non-contiguous source indexes.
    frame = frame.reset_index(drop=True)

    if group_column is not None:
        values = frame[group_column].astype("string").fillna("Missing")
        labels = sorted(
            values.unique().tolist(),
            key=lambda value: str(value).casefold(),
        )
        colours = {
            label: palette[index % len(palette)]
            for index, label in enumerate(labels)
        }

        def colour_row(row: pd.Series) -> list[str]:
            label = str(row[group_column]) if pd.notna(row[group_column]) else "Missing"
            background = colours.get(label, palette[0])
            return [
                f"background-color: {background}; color: #172033"
                for _ in row.index
            ]

        styler = frame.style.apply(colour_row, axis=1)
        styler.set_caption(f"Record colours indicate groups in: {group_column}")
        return styler

    def band_rows(row: pd.Series) -> list[str]:
        background = "#FFFFFF" if row.name % 2 == 0 else "#F1F5F9"
        return [
            f"background-color: {background}; color: #172033"
            for _ in row.index
        ]

    styler = frame.style.apply(band_rows, axis=1)
    styler.set_caption("Alternating row colours distinguish records")
    return styler


def recommendation_display_frame(frame: pd.DataFrame) -> pd.DataFrame:
    """Build a presentation-only view using readable *_display columns.

    The source DataFrame and workbook are never modified. Raw structured values
    remain available through the original-data CSV download.
    """
    view = frame.copy(deep=True)
    display_columns = [
        column for column in view.columns
        if str(column).endswith("_display")
        and str(column)[:-len("_display")] in view.columns
    ]

    for display_column in display_columns:
        source_column = str(display_column)[:-len("_display")]
        readable = view[display_column]
        has_readable_value = readable.notna() & readable.astype(str).str.strip().ne("")
        view[source_column] = readable.where(has_readable_value, view[source_column])

    if display_columns:
        view = view.drop(columns=display_columns)
    return view


def show_full_table(
    frame: pd.DataFrame,
    key: str,
    download_frame: pd.DataFrame | None = None,
) -> None:
    """Render every record; optionally download a separate source DataFrame."""
    source_for_download = frame if download_frame is None else download_frame
    frame = prepare_frame(frame)
    if frame.empty:
        st.info("This dataset contains no records to display.")
        return

    cell_count = int(frame.shape[0]) * int(frame.shape[1])
    # Pandas Styler has a default 262,144-cell render limit. Keep styling only
    # for smaller frames; large tables remain fully visible as ordinary data.
    style_cell_limit = 200_000
    if cell_count <= style_cell_limit:
        st.caption(
            f"Showing all {len(frame):,} records and {len(frame.columns):,} columns. "
            "Record colours distinguish groups when a suitable category is present."
        )
        display_frame = style_record_rows(frame)
    else:
        st.caption(
            f"Showing all {len(frame):,} records and {len(frame.columns):,} columns "
            f"({cell_count:,} cells). Row styling is disabled for this large table "
            "to keep rendering reliable; no rows or columns are omitted."
        )
        display_frame = frame
    st.dataframe(
        display_frame,
        use_container_width=True,
        height=560,
        hide_index=True,
        key=key,
    )
    st.download_button(
        "Download displayed dataset (CSV)",
        data=source_for_download.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{key}.csv",
        mime="text/csv",
        key=f"{key}_csv_download",
    )


def show_chart_section(title: str, description: str | None = None) -> None:
    st.markdown(f"### {title}")
    if description:
        st.caption(description)


def show_demand_summary(frame: pd.DataFrame) -> None:
    """Show transparent numeric summaries without assigning guessed semantics."""
    measures = numeric_measure_columns(frame)
    if not measures:
        st.info(
            "No numeric measure candidates were identified from the selected table's "
            "structure. The full schema remains available for inspection."
        )
        return

    st.markdown("#### Numeric measure snapshots")
    st.caption(
        "These are structural numeric summaries, not confirmed demand KPIs. "
        "Field meaning must be established from dataset documentation or pipeline evidence."
    )
    for start in range(0, min(len(measures), 12), 4):
        subset = measures[start:start + 4]
        columns = st.columns(len(subset))
        for slot, source_column in zip(columns, subset):
            values = pd.to_numeric(frame[source_column], errors="coerce").dropna()
            if values.empty:
                continue
            slot.metric(f"Mean · {source_column}", f"{values.mean():,.3g}")
            slot.caption(f"{len(values):,} valid values")


def load_artifact_frames(
    run: Mapping[str, Any],
    relative_path: str,
) -> dict[str, pd.DataFrame]:
    """Read a known allowlisted artifact using the data-service interface."""
    try:
        return nonempty_frames(load_run_tables(run, relative_path))
    except Exception:
        pass

    if relative_path.lower().endswith((".xlsx", ".xls")):
        try:
            return nonempty_frames(load_run_workbook(run, relative_path))
        except Exception:
            pass

    if relative_path.lower().endswith(".json"):
        try:
            return json_to_frames(
                load_run_json(run, relative_path),
                Path(relative_path).stem,
            )
        except Exception:
            pass

    return {}


def load_source_frames(
    run: Mapping[str, Any],
    source: str,
) -> dict[str, pd.DataFrame]:
    """Load one source through the read-only data service."""
    if source == "Original input":
        try:
            return nonempty_frames(load_original_dataset(run))
        except Exception as exc:
            st.warning(f"Could not load the verified original input dataset: {exc}")
            return {}

    paths = {
        "Cleaned dataset": "processed/processed_dataset.json",
        "Analytical dataset": "output/analytical_dataset.json",
        "Analysis results": "output/analysis_results.json",
        "Forecast results": "output/forecast_results.json",
        "Recommendations": "output/recommendation_results.xlsx",
    }
    path = paths.get(source)
    if not path:
        return {}

    try:
        return load_artifact_frames(run, path)
    except Exception as exc:
        st.warning(f"Could not load {source.lower()}: {exc}")
        return {}


def choose_run(
    successful_runs: list[dict[str, Any]],
) -> dict[str, Any] | None:
    """Let the user inspect an existing successful run; never execute a run."""
    if not successful_runs:
        return None

    options = {
        str(item.get("run_id", "Unknown run")): item
        for item in successful_runs
    }
    labels = list(options.keys())

    default_run = get_latest_successful_run()
    default_id = (
        str(default_run.get("run_id"))
        if isinstance(default_run, Mapping)
        else labels[0]
    )
    index = labels.index(default_id) if default_id in labels else 0

    chosen_id = st.sidebar.selectbox(
        "Successful pipeline run",
        labels,
        index=index,
        key="edi_selected_run",
    )
    return options[chosen_id]


def render_empty_state(title: str, message: str) -> None:
    st.subheader(title)
    st.info(message)


def render_exploration_charts(
    frame: pd.DataFrame,
    *,
    allow_time_series: bool = True,
) -> bool:
    """Render supported chart families only; omit unsupported chart sections."""
    if not isinstance(frame, pd.DataFrame) or frame.empty or frame.shape[1] == 0:
        return False

    numeric = numeric_measure_columns(frame)
    dates = datetime_columns(frame)
    categories = categorical_columns(frame)

    useful_categories = [
        column for column in categories
        if 2 <= frame[column].nunique(dropna=True)
        <= min(20, max(2, int(len(frame) * 0.8)))
    ]
    booleans = boolean_columns(frame)
    has_missing = bool(frame.isna().any().any())

    names = {
        str(column).strip().lower().replace(" ", "_")
        for column in frame.columns
    }
    has_geo = (
        any(name in names for name in ("latitude", "lat"))
        and any(name in names for name in ("longitude", "lon", "lng"))
    )

    supported = []
    if has_missing:
        supported.append("Data quality")
    if numeric:
        supported.append("Numeric distributions")
    if useful_categories:
        supported.append("Category comparisons")
    if allow_time_series and dates and numeric:
        supported.append("Time patterns")
    if len(numeric) >= 2:
        supported.append("Measure relationships")
    if booleans:
        supported.append("Boolean comparisons")
    if has_geo:
        supported.append("Geographic analysis")

    # Empty states should not become empty headings, charts, or warning boxes.
    if not supported:
        return False

    if has_missing:
        show_chart_section("Data quality", "Missing-value patterns in the selected data.")
        render_missing_value_charts(frame)

    if numeric:
        show_chart_section(
            "Numeric distributions",
            "Distributions of actual numeric measures; identifiers are excluded.",
        )
        render_numeric_distributions(frame[numeric])

    if useful_categories:
        show_chart_section(
            "Category comparisons",
            "Only categories with repeated, useful values are shown.",
        )
        chart_columns = useful_categories + [
            column for column in numeric if column not in useful_categories
        ]
        render_category_comparisons(frame[chart_columns])

    if allow_time_series and dates and numeric:
        show_chart_section(
            "Time patterns",
            "Trends from actual date fields and numeric measures.",
        )
        chart_columns = dates + [
            column for column in numeric if column not in dates
        ]
        render_time_series_charts(frame[chart_columns])

    if len(numeric) >= 2:
        show_chart_section(
            "Measure relationships",
            "Relationships between numeric measures.",
        )
        render_relationship_charts(frame[numeric])

    if booleans:
        show_chart_section(
            "Boolean comparisons",
            "Comparisons for actual Boolean fields.",
        )
        render_boolean_comparisons(frame[booleans])

    if has_geo:
        show_chart_section(
            "Geographic analysis",
            "Locations using actual latitude and longitude values.",
        )
        render_geographic_charts(frame)

    return True


# ---------------------------------------------------------------------------
# Selected run and navigation
# ---------------------------------------------------------------------------

st.sidebar.title("EDI Workspace")
st.sidebar.caption("Enterprise Demand Intelligence")

st.sidebar.metric(
    "Total Application Views",
    f"{application_views:,}",
)

if st.sidebar.button("Sign out", key="edi_sign_out", use_container_width=True):
    st.session_state["edi_authenticated"] = False
    st.rerun()


# ---------------------------------------------------------------------------
# Independent Company Intelligence — independent page navigation
# ---------------------------------------------------------------------------

pages = [
    "Executive Overview",
    "Dataset Explorer",
    "Demand Analytics",
    "Forecasting",
    "Recommendations",
    "Reports & Run History",
    "Independent Company Intelligence",
]

page = st.sidebar.radio(
    "Navigation",
    pages,
    key="edi_navigation",
)

# This page does not require a successful EDI pipeline run.
if page == "Independent Company Intelligence":
    render_independent_intelligence()
    st.stop()


# ---------------------------------------------------------------------------
# Existing EDI successful-run validation — unchanged
# ---------------------------------------------------------------------------

successful_runs = get_successful_runs()
if not successful_runs:
    st.title("Enterprise Demand Intelligence")
    st.warning(
        "No verified successful pipeline runs were found. Run the EDI pipeline "
        "first, then refresh this dashboard. No sample data is substituted."
    )
    history = get_run_history()
    if history:
        st.subheader("Recorded run status")
        st.dataframe(
            pd.DataFrame([
                {
                    "Run ID": item.get("run_id", ""),
                    "Successful": bool(item.get("successful", False)),
                    "Manifest updated": (
                        datetime.fromtimestamp(item["modified_at"])
                        .astimezone()
                        .strftime("%Y-%m-%d %H:%M:%S")
                        if item.get("modified_at")
                        else ""
                    ),
                }
                for item in history
            ]),
            use_container_width=True,
            hide_index=True,
        )
    st.stop()

run = choose_run(successful_runs)
if not isinstance(run, Mapping):
    st.error("The selected successful run could not be loaded.")
    st.stop()

run_id = str(run.get("run_id", "Unknown run"))
manifest = run.get("manifest", {})
if not isinstance(manifest, Mapping):
    manifest = {}

st.sidebar.caption(f"Selected run: {run_id}")
if run.get("modified_at"):
    st.sidebar.caption(
        "Manifest updated: "
        + datetime.fromtimestamp(run["modified_at"])
        .astimezone()
        .strftime("%Y-%m-%d %H:%M:%S")
    )


# ---------------------------------------------------------------------------
# Page 1 — Executive Overview
# ---------------------------------------------------------------------------

if page == "Executive Overview":
    st.title("Executive Overview")
    st.markdown(
        '<p class="edi-subtitle">A live view of the data available in your selected run.</p>',
        unsafe_allow_html=True,
    )

    source = st.selectbox(
        "Dataset stage",
        ["Original input", "Cleaned dataset"],
        key="edi_overview_source",
        help="Choose the verified original input or the pipeline's cleaned output.",
    )
    frames = load_source_frames(run, source)

    if frames:
        total_source_rows = sum(len(table) for table in frames.values())
        st.caption(
            f"Source inventory: {len(frames):,} tables · "
            f"{total_source_rows:,} total rows across those tables. "
            "The table selector below shows one table at a time; "
            "the selected table's row and column counts are shown separately."
        )

    if not frames:
        render_empty_state(
            "Dataset unavailable",
            "The selected dataset could not be loaded. Check the selected run's "
            "manifest and artifact availability.",
        )
    else:
        table_name, frame = selected_frame(
            frames,
            "edi_overview_table_v2",
            "Dataset table",
        )

        if frame is not None:
            st.caption(f"Source: {source} · Table: {table_name}")
            show_kpis(frame)
            st.divider()

            heading = "Actual Dataset" if source == "Original input" else "Cleaned Dataset"
            show_chart_section(
                heading,
                "All rows and columns are available in one scrollable table. "
                "Nested objects are expanded into readable columns where possible.",
            )
            show_full_table(frame, f"overview_{source}_{table_name}")


# ---------------------------------------------------------------------------
# Page 2 — Dataset Explorer
# ---------------------------------------------------------------------------

elif page == "Dataset Explorer":
    st.title("Dataset Explorer")
    st.markdown(
        '<p class="edi-subtitle">Explore data quality, distributions, relationships, and time patterns.</p>',
        unsafe_allow_html=True,
    )

    source = st.selectbox(
        "Explore data from",
        [
            "Original input",
            "Cleaned dataset",
            "Analytical dataset",
            "Analysis results",
        ],
        key="edi_explorer_source",
    )
    frames = load_source_frames(run, source)

    if not frames:
        render_empty_state(
            "No dataset available",
            "The selected source did not produce any readable tables.",
        )
    else:
        table_name, frame = selected_frame(
            frames,
            "edi_explorer_table_v2",
            "Dataset",
        )
        if frame is not None:
            st.caption(f"Source: {source} · Dataset: {table_name}")
            show_kpis(frame)
            st.divider()
            render_exploration_charts(frame)


# ---------------------------------------------------------------------------
# Page 3 — Demand Analytics
# ---------------------------------------------------------------------------

elif page == "Demand Analytics":
    st.title("Demand Analytics")
    st.markdown(
        '<p class="edi-subtitle">Demand-focused exploration using actual available fields.</p>',
        unsafe_allow_html=True,
    )

    source = st.selectbox(
        "Analyze data from",
        [
            "Analytical dataset",
            "Analysis results",
            "Original input",
            "Cleaned dataset",
        ],
        key="edi_demand_source",
    )
    frames = load_source_frames(run, source)

    if frames:
        st.caption(
            f"Source inventory: {len(frames):,} tables · "
            f"{sum(len(table) for table in frames.values()):,} total rows. "
            "Choose a demand-related table from the dataset selector."
        )

    if not frames:
        render_empty_state(
            "Demand data unavailable",
            "No readable dataset was found for the selected source.",
        )
    else:
        table_name, frame = selected_frame(
            frames,
            "edi_demand_table_v2",
            "Dataset",
        )

        if frame is not None:
            st.caption(f"Source: {source} · Dataset: {table_name}")
            show_kpis(frame)

            # Use structural evidence rather than a fixed vocabulary of column names.
            # Numeric measures, temporal fields, and useful dimensions are candidates
            # for exploration; their business meaning is not asserted by this page.
            measure_fields = numeric_measure_columns(frame)
            time_fields = datetime_columns(frame)
            dimension_fields = [
                column for column in categorical_columns(frame)
                if 2 <= frame[column].nunique(dropna=True)
                <= min(30, max(2, int(len(frame) * 0.8)))
            ]
            candidate_fields = list(dict.fromkeys(
                time_fields + measure_fields + dimension_fields
            ))

            if measure_fields:
                show_demand_summary(frame)
            if candidate_fields:
                candidate_frame = frame.loc[:, candidate_fields].copy()
                render_exploration_charts(
                    candidate_frame,
                    allow_time_series=True,
                )
            else:
                st.info(
                    "No temporal fields, numeric measures, or useful categorical "
                    "dimensions were detected in this table. Select another table "
                    "or inspect the full dataset schema."
                )


# ---------------------------------------------------------------------------
# Page 4 — Forecasting
# ---------------------------------------------------------------------------

elif page == "Forecasting":
    st.title("Forecasting")
    st.markdown(
        '<p class="edi-subtitle">Inspect forecast artifacts produced by the pipeline.</p>',
        unsafe_allow_html=True,
    )

    frames = load_source_frames(run, "Forecast results")
    if not frames:
        render_empty_state(
            "Forecast results unavailable",
            "No readable forecast result tables were found for this run. "
            "This page does not generate forecasts itself.",
        )
    else:
        table_name, frame = selected_frame(
            frames,
            "edi_forecast_table_v2",
            "Forecast result dataset",
        )

        if frame is not None:
            st.caption(f"Forecast artifact · {table_name}")
            show_kpis(frame)
            st.caption(
                "Each row represents a forecast-series summary. The selected "
                "artifact does not provide one row per future day, so the dashboard "
                "does not draw a fabricated daily forecast line."
            )

            # Forecast artifacts may use arbitrary field names. Explore numeric
            # measures structurally and let the artifact/schema provide meaning.
            forecast_numeric = numeric_measure_columns(frame)

            if forecast_numeric:
                chart_columns = list(dict.fromkeys(
                    [
                        column
                        for column in ("table", "ForecastConfidence", "status")
                        if column in frame.columns
                    ] + forecast_numeric
                ))
                render_exploration_charts(
                    frame.loc[:, chart_columns],
                    allow_time_series=False,
                )


# ---------------------------------------------------------------------------
# Page 5 — Recommendations
# ---------------------------------------------------------------------------

elif page == "Recommendations":
    st.title("Recommendations")
    st.markdown(
        '<p class="edi-subtitle">Switch between recommendations, forecast records, evidence, and underlying datasets.</p>',
        unsafe_allow_html=True,
    )

    source_options = [
        "Recommendations",
        "Forecast results",
        "Analysis results",
        "Analytical dataset",
        "Cleaned dataset",
        "Original input",
    ]
    selected_source = st.selectbox(
        "Dataset source",
        source_options,
        index=0,
        key="edi_recommendation_source",
    )
    frames = load_source_frames(run, selected_source)

    if not frames:
        render_empty_state(
            f"{selected_source} unavailable",
            "No readable tables were found for this source in the selected successful run. "
            "Check the artifact inventory on the Reports & Run History page.",
        )
    else:
        worksheet_names = list(frames)
        preferred_name = (
            "Recommendations"
            if selected_source == "Recommendations"
            and "Recommendations" in frames
            else worksheet_names[0]
        )
        default_index = worksheet_names.index(preferred_name)

        selected_worksheet = st.selectbox(
            "Table, worksheet, or evidence dataset",
            worksheet_names,
            index=default_index,
            format_func=safe_label,
            key=f"edi_recommendation_table_{selected_source.replace(' ', '_').lower()}",
        )
        table_name = selected_worksheet
        frame = prepare_frame(frames[selected_worksheet])

        st.caption(f"Source: {selected_source} · Table/worksheet: {table_name}")
        show_kpis(frame)
        st.divider()

        show_chart_section(
            "Data insights",
            "Charts use actual columns in the selected dataset; identifiers are not treated as business measures.",
        )

        if selected_source == "Recommendations":
            render_exploration_charts(frame, allow_time_series=False)

        elif selected_source == "Forecast results":
            # Forecast artifacts may use arbitrary field names. Explore numeric
            # measures structurally and let the artifact/schema provide meaning.
            forecast_numeric = numeric_measure_columns(frame)
            selected_columns = [
                column for column in frame.columns
                if column in forecast_numeric
                or str(column).lower() in (
                    "table",
                    "forecastconfidence",
                    "status",
                )
            ]
            if selected_columns:
                render_exploration_charts(
                    frame[selected_columns],
                    allow_time_series=False,
                )

        elif selected_source in (
            "Analysis results",
            "Analytical dataset",
            "Cleaned dataset",
            "Original input",
        ):
            render_exploration_charts(
                frame,
                allow_time_series=bool(datetime_columns(frame)),
            )

        st.divider()
        show_chart_section(
            "Selected records",
            "Full-width, scrollable table of selected records. Record colours distinguish groups when a suitable category exists.",
        )
        if selected_source == "Recommendations":
            # Prefer the workbook's existing human-readable display fields in
            # the UI. Keep the original workbook data intact for CSV export.
            show_full_table(
                recommendation_display_frame(frame),
                f"recommendations_{selected_source.replace(' ', '_').lower()}_{table_name}",
                download_frame=frame,
            )
        else:
            show_full_table(
                frame,
                f"recommendations_{selected_source.replace(' ', '_').lower()}_{table_name}",
            )


# ---------------------------------------------------------------------------
# Page 6 — Reports & Run History
# ---------------------------------------------------------------------------

elif page == "Reports & Run History":
    st.title("Reports & Run History")
    st.markdown(
        '<p class="edi-subtitle">Operational visibility into pipeline runs, stages, logs, and artifacts.</p>',
        unsafe_allow_html=True,
    )

    all_runs = get_run_history()
    total_runs = len(all_runs)
    successful_count = sum(
        1 for item in all_runs if item.get("successful", False)
    )
    unsuccessful_count = total_runs - successful_count

    metric_columns = st.columns(3)
    metric_columns[0].metric("Total runs", format_integer(total_runs))
    metric_columns[1].metric("Successful runs", format_integer(successful_count))
    metric_columns[2].metric("Other runs", format_integer(unsuccessful_count))

    st.divider()
    show_chart_section("Selected run")

    run_details = {
        "Run ID": run_id,
        "Successful": bool(run.get("successful", False)),
        "Run directory": str(run.get("run_directory", "")),
        "Manifest": str(run.get("manifest_path", "")),
    }
    detail_columns = st.columns(2)
    for index, (label, value) in enumerate(run_details.items()):
        with detail_columns[index % 2]:
            st.text_input(
                label,
                value=str(value),
                disabled=True,
                key=f"edi_run_detail_{index}",
            )

    st.divider()
    show_chart_section("Pipeline stage status")
    stages = manifest.get("stages", {})
    stage_rows = []

    if isinstance(stages, Mapping):
        for stage in PIPELINE_STAGES:
            stage_info = stages.get(stage, {})
            if not isinstance(stage_info, Mapping):
                stage_info = {}

            stage_rows.append(
                {
                    "Stage": safe_label(stage),
                    "Status": str(stage_info.get("status", "UNKNOWN")),
                    "Started": str(
                        stage_info.get("started_at", stage_info.get("start_time", ""))
                    ),
                    "Finished": str(
                        stage_info.get("finished_at", stage_info.get("end_time", ""))
                    ),
                    "Duration": str(
                        stage_info.get("duration_seconds", stage_info.get("duration", ""))
                    ),
                }
            )

    if stage_rows:
        st.dataframe(
            pd.DataFrame(stage_rows),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Stage metadata is not available in this run's manifest.")

    st.divider()
    show_chart_section("Artifact inventory")

    try:
        catalog = get_artifact_catalog(run)
    except Exception as exc:
        catalog = []
        st.warning(f"Could not read the artifact catalog: {exc}")

    if catalog:
        catalog_frame = pd.DataFrame([
            {
                "Stage": safe_label(item.get("stage", "")),
                "Artifact": str(item.get("relative_path", "")),
                "Available": bool(item.get("exists", False)),
                "Size (bytes)": item.get("size_bytes", 0),
            }
            for item in catalog
        ])
        st.dataframe(
            catalog_frame,
            use_container_width=True,
            hide_index=True,
        )

        existing_artifacts = [
            str(item.get("relative_path"))
            for item in catalog
            if item.get("exists") and item.get("relative_path")
        ]
        if existing_artifacts:
            chosen_artifact = st.selectbox(
                "Download an available artifact",
                existing_artifacts,
                key="edi_download_artifact",
            )
            try:
                artifact_path = get_artifact_path(run, chosen_artifact)
                if artifact_path is not None and artifact_path.is_file():
                    st.download_button(
                        "Download selected artifact",
                        data=artifact_path.read_bytes(),
                        file_name=artifact_path.name,
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                            if artifact_path.suffix.lower() == ".xlsx"
                            else "application/json"
                            if artifact_path.suffix.lower() == ".json"
                            else "application/octet-stream"
                        ),
                        key="edi_artifact_download_button",
                    )
                else:
                    st.info("The selected artifact is no longer available.")
            except Exception as exc:
                st.warning(f"Artifact access was denied or failed: {exc}")
    else:
        st.info("No artifact inventory is available for this run.")

    st.divider()
    show_chart_section("Run log")

    run_directory = run.get("run_directory")
    if run_directory:
        trusted_run_directory = Path(run_directory).resolve()
        log_path = (trusted_run_directory / "execution.log").resolve()
        try:
            log_path.relative_to(trusted_run_directory)
        except ValueError:
            log_path = None

        if log_path is not None and log_path.is_file():
            try:
                log_text = log_path.read_text(
                    encoding="utf-8",
                    errors="replace",
                )
                with st.expander("View execution.log", expanded=False):
                    st.code(log_text[-30000:] or "The log is empty.")
            except OSError as exc:
                st.warning(f"Could not read the run log: {exc}")
        else:
            st.info("No execution.log file is available for this run.")
    else:
        st.info("The selected run directory is not available.")

    st.divider()
    st.markdown("### Power BI integration")
    powerbi_metadata = manifest.get("powerbi_metadata")

    if isinstance(powerbi_metadata, Mapping) and powerbi_metadata:
        summary_keys = (
            "status",
            "workspace",
            "project",
            "report",
            "dataset_id",
            "published",
        )
        summary = {
            key: value
            for key, value in powerbi_metadata.items()
            if key in summary_keys
            and not isinstance(value, (Mapping, list, tuple))
        }

        if summary:
            summary_columns = st.columns(min(3, max(1, len(summary))))
            for index, (key, value) in enumerate(summary.items()):
                summary_columns[index % len(summary_columns)].metric(
                    safe_label(key),
                    str(value)[:120],
                )
        else:
            st.caption("Power BI metadata is recorded for this run.")

        with st.expander("View complete Power BI metadata", expanded=False):
            st.json(dict(powerbi_metadata))

    elif powerbi_metadata:
        st.caption("Power BI metadata is available for this run.")
        with st.expander("View complete Power BI metadata", expanded=False):
            st.json(powerbi_metadata)
    else:
        st.caption("No Power BI integration metadata is recorded for this run.")

    st.divider()
    show_chart_section("Run history")

    if all_runs:
        history_rows = []
        for item in all_runs:
            item_manifest = item.get("manifest", {})
            if not isinstance(item_manifest, Mapping):
                item_manifest = {}

            history_rows.append(
                {
                    "Run ID": str(item.get("run_id", "")),
                    "Successful": bool(item.get("successful", False)),
                    "Modified": (
                        datetime.fromtimestamp(item["modified_at"]).strftime(
                            "%Y-%m-%d %H:%M:%S"
                        )
                        if item.get("modified_at")
                        else ""
                    ),
                    "Dataset": str(item_manifest.get("dataset_path", "")),
                }
            )

        st.dataframe(
            pd.DataFrame(history_rows),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("No run history records are available.")