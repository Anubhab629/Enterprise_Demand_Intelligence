"""Dataset-adaptive chart components for the EDI Streamlit dashboard.

The module discovers chart options from actual records. It does not fabricate
business measures, forecasts, recommendation scores, or geographic locations.

Each chart family renders at most one chart at a time. A separate selector
allows the user to switch between available analyses in that family.
"""
from __future__ import annotations

import hashlib
import warnings
from typing import Any

import pandas as pd
import plotly.express as px
import streamlit as st

from streamlit_app.analytics import (
    boolean_columns,
    categorical_columns,
    datetime_columns,
    numeric_columns,
)

PALETTE = [
    "#60A5FA",
    "#34D399",
    "#FBBF24",
    "#F472B6",
    "#A78BFA",
    "#FB7185",
    "#22D3EE",
    "#A3E635",
    "#F97316",
    "#C084FC",
]

MAX_CATEGORIES = 20
MAX_SCATTER_POINTS = 5000
MAX_TIME_SERIES_POINTS = 3000


def _chart_theme(fig: Any) -> Any:
    """Apply the shared dark theme to a Plotly figure."""
    fig.update_layout(
        template="plotly_dark",
        colorway=PALETTE,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin={"l": 28, "r": 24, "t": 62, "b": 35},
        font={"size": 12, "color": "#E5E7EB"},
        title={"font": {"size": 17}},
        legend={
            "bgcolor": "rgba(0,0,0,0)",
            "borderwidth": 0,
        },
        hoverlabel={"font_size": 12},
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor="#28364B",
        zeroline=False,
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor="#28364B",
        zeroline=False,
    )
    return fig


def _safe_key(frame: pd.DataFrame, suffix: str) -> str:
    """Create stable Streamlit widget keys for a frame and widget purpose."""
    signature = repr(
        (
            list(map(str, frame.columns)),
            list(map(str, frame.dtypes)),
            frame.shape,
        )
    )

    try:
        sample = pd.util.hash_pandas_object(
            frame.head(10).astype(str),
            index=True,
        ).values.tobytes()
        signature += hashlib.sha256(sample).hexdigest()
    except (TypeError, ValueError):
        pass

    digest = hashlib.sha256(
        signature.encode("utf-8")
    ).hexdigest()[:12]

    return f"edi_chart_{suffix}_{digest}"


def _valid_frame(frame: Any) -> bool:
    """Check that a non-empty tabular frame was provided."""
    return (
        isinstance(frame, pd.DataFrame)
        and not frame.empty
        and len(frame.columns) > 0
    )


def _parse_dates(series: pd.Series) -> pd.Series:
    """Parse values from a date-like column discovered by analytics."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)

        try:
            return pd.to_datetime(
                series,
                errors="coerce",
                format="mixed",
            )
        except (TypeError, ValueError):
            return pd.to_datetime(
                series,
                errors="coerce",
            )


def _safe_numeric(
    frame: pd.DataFrame,
    column: Any,
) -> pd.Series:
    """Convert a numeric column to usable observations."""
    return pd.to_numeric(
        frame[column],
        errors="coerce",
    )


def _numeric_fields(frame: pd.DataFrame) -> list[Any]:
    """Return available numeric chart fields."""
    return numeric_columns(frame)


def _category_fields(frame: pd.DataFrame) -> list[Any]:
    """Return usable low-cardinality category fields."""
    return categorical_columns(
        frame,
        max_unique=MAX_CATEGORIES,
    )


def _family_selector(
    family: str,
    analyses: dict[str, Any],
) -> Any | None:
    """Render a family-specific selector and return its chosen analysis."""
    if not analyses:
        return None

    options = list(analyses.keys())

    if len(options) == 1:
        selected = options[0]
        st.caption(f"{family}: {selected}")
    else:
        selected = st.selectbox(
            f"{family} analysis",
            options,
            key=f"edi_family_{family.lower().replace(' ', '_')}",
        )

    return analyses.get(selected)


def _render_chart(fig: Any, key: str) -> None:
    """Render one chart with the shared theme."""
    st.plotly_chart(
        _chart_theme(fig),
        use_container_width=True,
        key=key,
    )


def _show_empty(message: str) -> None:
    """Explain why a chart family cannot be displayed."""
    st.info(message)


def render_numeric_distributions(
    frame: pd.DataFrame,
    *,
    title: str = "Numeric distribution",
) -> None:
    """Show one selected histogram or box plot for a numeric field."""
    if not _valid_frame(frame):
        return

    numeric = _numeric_fields(frame)

    if not numeric:
        return

    st.subheader(title)

    analyses: dict[str, tuple[Any, str]] = {}

    for column in numeric:
        analyses[f"Histogram · {column}"] = (column, "histogram")
        analyses[f"Box plot · {column}"] = (column, "box")

    selected = _family_selector(
        "Numeric distribution",
        analyses,
    )

    if selected is None:
        return

    column, chart_type = selected
    values = _safe_numeric(frame, column).dropna()

    if values.empty:
        _show_empty(
            f"No valid numeric observations are available for {column}."
        )
        return

    if chart_type == "histogram":
        bins = min(
            50,
            max(5, int(values.nunique())),
        )

        fig = px.histogram(
            x=values,
            nbins=bins,
            title=f"Distribution of {column}",
            labels={
                "x": str(column),
                "count": "Records",
            },
            color_discrete_sequence=[PALETTE[0]],
        )
        fig.update_traces(
            marker_line_color="#0B1120",
            marker_line_width=0.5,
        )
    else:
        fig = px.box(
            y=values,
            points="outliers",
            title=f"Spread of {column}",
            labels={"y": str(column)},
            color_discrete_sequence=[PALETTE[1]],
        )

    _render_chart(
        fig,
        _safe_key(frame, f"distribution_{column}_{chart_type}"),
    )


def render_category_comparisons(
    frame: pd.DataFrame,
    *,
    title: str = "Category comparison",
) -> None:
    """Show one selected category comparison at a time."""
    if not _valid_frame(frame):
        return

    categories = _category_fields(frame)
    numeric = _numeric_fields(frame)

    if not categories:
        return

    st.subheader(title)

    analyses: dict[str, tuple[Any, Any | None, str]] = {}

    for category in categories:
        analyses[f"Record count · {category}"] = (
            category,
            None,
            "count",
        )

        for measure in numeric:
            analyses[f"{measure} by {category}"] = (
                category,
                measure,
                "measure",
            )

    selected = _family_selector(
        "Category comparison",
        analyses,
    )

    if selected is None:
        return

    category, measure, mode = selected

    if mode == "count":
        data = (
            frame[category]
            .dropna()
            .astype(str)
            .value_counts()
            .head(MAX_CATEGORIES)
            .rename_axis("Category")
            .reset_index(name="Records")
        )

        if data.empty:
            _show_empty("No category values are available.")
            return

        fig = px.bar(
            data.sort_values("Records"),
            x="Records",
            y="Category",
            orientation="h",
            title=f"Records by {category}",
            text="Records",
            color="Category",
            color_discrete_sequence=PALETTE,
        )
        fig.update_layout(showlegend=False)
        fig.update_traces(textposition="outside")
    else:
        data = pd.DataFrame(
            {
                "Category": frame[category].astype("string"),
                "Measure": _safe_numeric(frame, measure),
            }
        ).dropna()

        data = (
            data.groupby("Category", as_index=False)["Measure"]
            .mean()
            .sort_values("Measure", ascending=False)
            .head(MAX_CATEGORIES)
        )

        if data.empty:
            _show_empty(
                f"No valid {measure} values are available for comparison."
            )
            return

        fig = px.bar(
            data.sort_values("Measure"),
            x="Measure",
            y="Category",
            orientation="h",
            title=f"Average {measure} by {category}",
            labels={"Measure": str(measure)},
            color="Category",
            color_discrete_sequence=PALETTE,
        )
        fig.update_layout(showlegend=False)

    _render_chart(
        fig,
        _safe_key(frame, f"category_{category}_{measure}_{mode}"),
    )


def render_boolean_comparisons(
    frame: pd.DataFrame,
    *,
    title: str = "Boolean comparisons",
) -> None:
    """Show one Boolean-value frequency chart at a time."""
    if not _valid_frame(frame):
        return

    fields = boolean_columns(frame)

    if not fields:
        return

    st.subheader(title)

    selected = _family_selector(
        "Boolean comparison",
        {str(column): column for column in fields},
    )

    if selected is None:
        return

    counts = (
        frame[selected]
        .astype("string")
        .fillna("Missing")
        .value_counts()
        .rename_axis("Value")
        .reset_index(name="Records")
    )

    fig = px.bar(
        counts,
        x="Value",
        y="Records",
        title=f"Value counts · {selected}",
        color="Value",
        color_discrete_sequence=PALETTE,
        text="Records",
    )

    fig.update_layout(showlegend=False)

    _render_chart(
        fig,
        _safe_key(frame, f"boolean_{selected}"),
    )


def render_missing_value_charts(
    frame: pd.DataFrame,
    *,
    title: str = "Data quality",
) -> None:
    """Show one missing-value or completeness chart at a time."""
    if not _valid_frame(frame):
        return

    st.subheader(title)

    missing = frame.isna().sum()
    missing = missing[missing > 0].sort_values(ascending=False)

    analyses: dict[str, str] = {
        "Missing values by column": "missing",
        "Completeness by column": "completeness",
    }

    selected = _family_selector(
        "Data quality",
        analyses,
    )

    if selected == "missing":
        if missing.empty:
            _show_empty("No missing values were found in this dataset.")
            return

        data = (
            missing.head(MAX_CATEGORIES)
            .rename_axis("Column")
            .reset_index(name="Missing values")
        )

        fig = px.bar(
            data.sort_values("Missing values"),
            x="Missing values",
            y="Column",
            orientation="h",
            title="Columns with the most missing values",
            color="Missing values",
            color_continuous_scale="Blues",
        )
    else:
        completeness = (
            (1 - frame.isna().mean()) * 100
        ).sort_values()

        data = (
            completeness.head(MAX_CATEGORIES)
            .rename_axis("Column")
            .reset_index(name="Completeness (%)")
        )

        fig = px.bar(
            data,
            x="Completeness (%)",
            y="Column",
            orientation="h",
            title="Column completeness",
            range_x=[0, 100],
            color="Completeness (%)",
            color_continuous_scale="Teal",
        )

    _render_chart(
        fig,
        _safe_key(frame, f"data_quality_{selected}"),
    )


def render_time_series_charts(
    frame: pd.DataFrame,
    *,
    title: str = "Time patterns",
) -> None:
    """Show one date-based analysis at a time when dates and measures exist."""
    if not _valid_frame(frame):
        return

    temporal = datetime_columns(frame)
    numeric = _numeric_fields(frame)

    if not temporal or not numeric:
        return

    st.subheader(title)

    left, right = st.columns(2)

    with left:
        date_field = st.selectbox(
            "Date/time field",
            temporal,
            key=_safe_key(frame, "time_date_field"),
        )

    with right:
        measure = st.selectbox(
            "Measure",
            numeric,
            key=_safe_key(frame, "time_measure"),
        )

    analyses = {
        "Trend over time": "trend",
        "Distribution over time": "distribution",
    }

    selected = _family_selector(
        "Time-series",
        analyses,
    )

    if selected is None:
        return

    parsed = _parse_dates(frame[date_field])

    data = pd.DataFrame(
        {
            "__date": parsed,
            "__value": _safe_numeric(frame, measure),
        }
    ).dropna()

    if data.empty:
        _show_empty(
            "The selected date and measure do not contain enough valid observations."
        )
        return

    # Aggregate repeated timestamps using the mean. This avoids implicitly
    # assuming that every numeric field is additive.
    grouped = (
        data.groupby("__date", as_index=False)["__value"]
        .mean()
        .sort_values("__date")
    )

    if len(grouped) > MAX_TIME_SERIES_POINTS:
        grouped = (
            grouped.set_index("__date")
            .resample("D")["__value"]
            .mean()
            .dropna()
            .reset_index()
        )

    if grouped.empty:
        _show_empty("No usable time-based values remain after aggregation.")
        return

    if selected == "trend":
        fig = px.line(
            grouped,
            x="__date",
            y="__value",
            markers=True,
            title=f"{measure} over time",
            labels={
                "__date": str(date_field),
                "__value": str(measure),
            },
            color_discrete_sequence=[PALETTE[0]],
        )
        fig.update_traces(line={"width": 2.5})
    else:
        grouped["Period"] = grouped["__date"].dt.to_period("M").astype(str)

        fig = px.box(
            grouped,
            x="Period",
            y="__value",
            title=f"Variation in {measure} by month",
            labels={
                "Period": "Month",
                "__value": str(measure),
            },
            color_discrete_sequence=[PALETTE[1]],
        )
        fig.update_layout(xaxis_tickangle=-35)

    fig.update_layout(hovermode="x unified")

    _render_chart(
        fig,
        _safe_key(frame, f"time_{date_field}_{measure}_{selected}"),
    )


def render_relationship_charts(
    frame: pd.DataFrame,
    *,
    title: str = "Measure relationships",
) -> None:
    """Show one selected numeric relationship at a time."""
    if not _valid_frame(frame):
        return

    numeric = _numeric_fields(frame)

    if len(numeric) < 2:
        return

    categories = _category_fields(frame)

    st.subheader(title)

    analyses: dict[str, tuple[Any, Any, str]] = {}

    for x_field in numeric:
        for y_field in numeric:
            if x_field == y_field:
                continue

            analyses[
                f"{y_field} vs {x_field} · Scatter"
            ] = (x_field, y_field, "scatter")

            analyses[
                f"{y_field} vs {x_field} · Trend"
            ] = (x_field, y_field, "trend")

    selected = _family_selector(
        "Relationship",
        analyses,
    )

    if selected is None:
        return

    x_field, y_field, chart_type = selected

    color_field = None

    if categories:
        color_field = st.selectbox(
            "Optional category coloring",
            ["No category coloring", *categories],
            key=_safe_key(frame, "relationship_color"),
        )

        if color_field == "No category coloring":
            color_field = None

    columns = [x_field, y_field]
    if color_field is not None:
        columns.append(color_field)

    data = frame[columns].copy()
    data[x_field] = _safe_numeric(data, x_field)
    data[y_field] = _safe_numeric(data, y_field)
    data = data.dropna(subset=[x_field, y_field])

    if len(data) > MAX_SCATTER_POINTS:
        data = data.sample(
            MAX_SCATTER_POINTS,
            random_state=42,
        )

    if len(data) < 2:
        _show_empty(
            "Not enough valid numeric pairs are available for this relationship."
        )
        return

    if chart_type == "scatter":
        fig = px.scatter(
            data,
            x=x_field,
            y=y_field,
            color=color_field,
            opacity=0.7,
            title=f"{y_field} vs {x_field}",
            color_discrete_sequence=PALETTE,
        )
    else:
        fig = px.scatter(
            data,
            x=x_field,
            y=y_field,
            color=color_field,
            trendline="ols",
            opacity=0.65,
            title=f"Trend relationship · {y_field} vs {x_field}",
            color_discrete_sequence=PALETTE,
        )

    _render_chart(
        fig,
        _safe_key(frame, f"relationship_{x_field}_{y_field}_{chart_type}"),
    )


def render_geographic_charts(
    frame: pd.DataFrame,
    *,
    title: str = "Geographic analysis",
) -> None:
    """Render a genuine coordinate map only when latitude and longitude exist."""
    if not _valid_frame(frame):
        return

    numeric = _numeric_fields(frame)

    def find_coordinate(tokens: tuple[str, ...]) -> Any | None:
        for column in frame.columns:
            name = str(column).strip().casefold().replace(" ", "_")

            if name in tokens:
                return column

        return None

    latitude = find_coordinate(
        ("lat", "latitude", "y_coordinate", "geo_latitude")
    )
    longitude = find_coordinate(
        ("lon", "lng", "longitude", "x_coordinate", "geo_longitude")
    )

    # Do not treat a state, city, or country label as a map without a
    # dependable geocoding or boundary-matching source.
    if latitude is None or longitude is None:
        return

    st.subheader(title)

    measures = [
        column
        for column in numeric
        if column not in {latitude, longitude}
    ]

    measure_options = ["Point density", *measures]

    measure = st.selectbox(
        "Map measure",
        measure_options,
        key=_safe_key(frame, "geographic_measure"),
    )

    data = frame.copy()
    data["__lat"] = _safe_numeric(data, latitude)
    data["__lon"] = _safe_numeric(data, longitude)

    data = data[
        data["__lat"].between(-90, 90)
        & data["__lon"].between(-180, 180)
    ].copy()

    if data.empty:
        _show_empty(
            "No valid latitude/longitude pairs are available for mapping."
        )
        return

    if measure == "Point density":
        fig = px.scatter_mapbox(
            data,
            lat="__lat",
            lon="__lon",
            zoom=1,
            title="Observed geographic coordinates",
            mapbox_style="open-street-map",
            color_discrete_sequence=[PALETTE[0]],
        )
    else:
        data["__measure"] = _safe_numeric(data, measure)
        data = data.dropna(subset=["__measure"])

        if data.empty:
            _show_empty(
                f"No valid {measure} values are available for the map."
            )
            return

        fig = px.scatter_mapbox(
            data,
            lat="__lat",
            lon="__lon",
            size="__measure",
            color="__measure",
            zoom=1,
            title=f"{measure} by coordinate",
            mapbox_style="open-street-map",
            color_continuous_scale="Viridis",
        )

    fig.update_layout(
        mapbox={
            "style": "open-street-map",
        },
        margin={"l": 0, "r": 0, "t": 55, "b": 0},
    )

    _render_chart(
        fig,
        _safe_key(frame, f"geographic_{measure}"),
    )


def render_dataset_overview_charts(
    frame: pd.DataFrame,
) -> None:
    """Render the standard chart families for a selected dataset."""
    render_missing_value_charts(frame)
    render_category_comparisons(frame)
    render_numeric_distributions(frame)
    render_time_series_charts(frame)
    render_relationship_charts(frame)
    render_geographic_charts(frame)


def render_data_quality_charts(
    frame: pd.DataFrame,
) -> None:
    """Render data-quality chart families."""
    render_missing_value_charts(frame)
    render_boolean_comparisons(frame)


def render_demand_analytics_charts(
    frame: pd.DataFrame,
) -> None:
    """Render data-supported charts for an analytical result."""
    render_time_series_charts(frame, title="Demand and time patterns")
    render_category_comparisons(frame, title="Category performance")
    render_numeric_distributions(frame, title="Measure distributions")
    render_relationship_charts(frame, title="Measure relationships")
    render_geographic_charts(frame, title="Geographic analysis")


def render_forecast_charts(
    frame: pd.DataFrame,
) -> None:
    """Render charts for actual forecast result records."""
    render_time_series_charts(frame, title="Forecast time patterns")
    render_category_comparisons(frame, title="Forecast comparisons")
    render_numeric_distributions(frame, title="Forecast distributions")
    render_relationship_charts(frame, title="Forecast relationships")


def render_recommendation_charts(
    frame: pd.DataFrame,
) -> None:
    """Render charts supported by actual recommendation records."""
    render_category_comparisons(frame, title="Recommendation comparisons")
    render_numeric_distributions(frame, title="Recommendation distributions")
    render_relationship_charts(frame, title="Recommendation relationships")


def show_dataset_charts(
    frame: pd.DataFrame,
    *,
    show_profile: bool = True,
    show_numeric: bool = True,
    show_categories: bool = True,
    show_relationships: bool = True,
    show_time_series: bool = True,
) -> None:
    """Backward-compatible chart entry point used by the existing app."""
    if not isinstance(frame, pd.DataFrame):
        st.warning("The selected object is not a tabular DataFrame.")
        return

    if frame.empty or len(frame.columns) == 0:
        st.info("This table has no records or columns to visualize.")
        return

    if show_profile:
        st.subheader("Dataset profile")
        metrics = [
            ("Records", f"{len(frame):,}"),
            ("Columns", f"{len(frame.columns):,}"),
            ("Missing cells", f"{int(frame.isna().sum().sum()):,}"),
        ]

        containers = st.columns(len(metrics))

        for container, (label, value) in zip(containers, metrics):
            container.metric(label, value)

    if show_categories:
        render_category_comparisons(frame)

    if show_numeric:
        render_numeric_distributions(frame)

    if show_time_series:
        render_time_series_charts(frame)

    if show_relationships:
        render_relationship_charts(frame)