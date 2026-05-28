"""Plotly figures for the dashboard (port of the Chart.js / SVG-map visuals)."""
from __future__ import annotations

from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from data import CHANNEL_LABELS, MOS

# India states GeoJSON (Datameet, ISO names). Loaded lazily; map degrades to a
# bar chart if the network is unavailable.
INDIA_GEOJSON_URL = (
    "https://raw.githubusercontent.com/geohacker/india/master/state/india_telengana.geojson"
)

# Channel palette — JSW navy family anchors the "project" channels; retail and
# self-stocking stay as warmer accents. Hues + luminance both differ so the four
# remain distinguishable under deuteranopia / protanopia.
_PALETTE = {
    "rt": "#6366F1",    # retail — indigo
    "ss": "#F59E0B",    # self-stocking — amber
    "pdir": "#1B3A6B",  # project-direct — JSW navy
    "pd": "#2563EB",    # project-thru-dist — JSW blue
}
# Semi-transparent fills for line-area mode (must align with _PALETTE order).
_PALETTE_FILL = {
    "rt": "rgba(99,102,241,.10)",
    "ss": "rgba(245,158,11,.10)",
    "pdir": "rgba(27,58,107,.12)",
    "pd": "rgba(37,99,235,.10)",
}
_PRIMARY = "#1B3A6B"
_PRIMARY2 = "#2563EB"

_FONT_FAMILY = "Inter, system-ui, -apple-system, 'Segoe UI', sans-serif"


def _apply_jsw_layout(
    fig: go.Figure,
    *,
    height: int = 300,
    y_title: str | None = None,
    x_title: str | None = None,
    show_legend: bool = True,
) -> go.Figure:
    """Apply consistent JSW typography, gridlines, legend, and hover styling."""
    fig.update_layout(
        font=dict(family=_FONT_FAMILY, size=12, color="#0F172A"),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        margin=dict(l=10, r=10, t=10, b=10),
        height=height,
        showlegend=show_legend,
        legend=dict(
            orientation="h", y=-0.18, x=0,
            font=dict(size=11, color="#475569"),
            bgcolor="rgba(0,0,0,0)",
        ),
        hoverlabel=dict(
            bgcolor="#FFFFFF", bordercolor="#CBD5E1",
            font=dict(family=_FONT_FAMILY, size=12, color="#0F172A"),
        ),
    )
    fig.update_xaxes(
        gridcolor="#F1F5F9", zerolinecolor="#E2E8F0", linecolor="#E2E8F0",
        title=dict(text=x_title, font=dict(size=11, color="#475569")) if x_title else None,
        tickfont=dict(size=11, color="#475569"),
    )
    fig.update_yaxes(
        gridcolor="#F1F5F9", zerolinecolor="#E2E8F0", linecolor="#E2E8F0",
        title=dict(text=y_title, font=dict(size=11, color="#475569")) if y_title else None,
        tickfont=dict(size=11, color="#475569"),
    )
    return fig


def _gran_key(d: datetime | None, gran: str) -> str:
    if d is None:
        return ""
    if gran == "day":
        return d.strftime("%Y-%m-%d")
    if gran == "week":
        iso = d.isocalendar()
        return f"{iso[0]}-W{iso[1]:02d}"
    if gran == "month":
        return d.strftime("%Y-%m")
    if gran == "year":
        return str(d.year)
    return d.strftime("%Y-%m-%d")


def channel_trend(df: pd.DataFrame, gran: str = "month", chart_type: str = "line") -> go.Figure:
    """Ordered MT over time, split by channel."""
    fig = go.Figure()
    if not len(df):
        return _empty(fig, "No data")
    work = df[df["_d"].notna()].copy()
    work["_k"] = work["_d"].apply(lambda d: _gran_key(d, gran))
    keys = sorted(work["_k"].unique())
    for ch, label in CHANNEL_LABELS.items():
        sub = work[work["_ch"] == ch]
        if not len(sub):
            continue
        series = sub.groupby("_k")["_q"].sum().reindex(keys, fill_value=0)
        if chart_type == "bar":
            fig.add_bar(x=keys, y=series.values, name=label, marker_color=_PALETTE[ch])
        else:
            fig.add_scatter(
                x=keys, y=series.values, mode="lines+markers", name=label,
                line=dict(color=_PALETTE[ch], width=2.5),
                marker=dict(size=6),
                fill="tozeroy", fillcolor=_PALETTE_FILL[ch],
            )
    fig.update_layout(
        barmode="stack" if chart_type == "bar" else None,
        hovermode="x unified",
    )
    return _apply_jsw_layout(fig, y_title="Ordered MT")


def _hbar(df: pd.DataFrame, label_col: str, value_col: str, title: str,
          color: str = _PRIMARY2) -> go.Figure:
    fig = go.Figure()
    if not len(df):
        return _empty(fig, "No data")
    d = df.sort_values(value_col, ascending=True)
    fig.add_bar(
        x=d[value_col], y=d[label_col], orientation="h", marker_color=color,
        texttemplate="%{x:,.0f}", textposition="outside",
        textfont=dict(size=11, color="#334155"), cliponaxis=False,
    )
    return _apply_jsw_layout(fig, x_title=title, show_legend=False)


def top_states(df: pd.DataFrame, n: int = 10) -> go.Figure:
    agg = (df.groupby("_st")["_q"].sum().reset_index()
           .sort_values("_q", ascending=False).head(n)) if len(df) else pd.DataFrame()
    return _hbar(agg, "_st", "_q", "Ordered MT", _PRIMARY)


def top_distributors(df: pd.DataFrame, n: int = 10) -> go.Figure:
    agg = (df.groupby("_dn")["_q"].sum().reset_index()
           .sort_values("_q", ascending=False).head(n)) if len(df) else pd.DataFrame()
    return _hbar(agg, "_dn", "_q", "Ordered MT", _PRIMARY2)


def _pie(df: pd.DataFrame, label_col: str, value_col: str, n: int = 10) -> go.Figure:
    fig = go.Figure()
    if not len(df):
        return _empty(fig, "No data")
    agg = (df.groupby(label_col)[value_col].sum().reset_index()
           .sort_values(value_col, ascending=False))
    if len(agg) > n:
        head = agg.head(n)
        other = agg[value_col][n:].sum()
        head = pd.concat([head, pd.DataFrame([{label_col: "Other", value_col: other}])])
        agg = head
    fig.add_pie(
        labels=agg[label_col], values=agg[value_col], hole=0.55,
        marker=dict(line=dict(color="#FFFFFF", width=2)),
        textfont=dict(size=11, color="#0F172A"),
    )
    fig.update_layout(legend=dict(orientation="v", font=dict(size=10)))
    return _apply_jsw_layout(fig)


def grade_mix(df: pd.DataFrame) -> go.Figure:
    return _pie(df, "_gr", "_q")


def plant_mix(df: pd.DataFrame) -> go.Figure:
    return _pie(df, "_cm", "_q")


def india_map(state_values: pd.DataFrame, metric_label: str) -> go.Figure:
    """Choropleth of a metric by ship-to state. Falls back to a bar chart offline."""
    fig = go.Figure()
    if not len(state_values):
        return _empty(fig, "No data")
    try:
        import json
        import urllib.request

        with urllib.request.urlopen(INDIA_GEOJSON_URL, timeout=8) as resp:
            geojson = json.load(resp)
        feature_key = "properties.NAME_1"
        fig = px.choropleth(
            state_values, geojson=geojson, locations="state",
            featureidkey=feature_key, color="value",
            color_continuous_scale=["#EEF4FA", _PRIMARY],
        )
        fig.update_geos(fitbounds="locations", visible=False)
        fig.update_layout(
            font=dict(family=_FONT_FAMILY, size=12, color="#0F172A"),
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            margin=dict(l=0, r=0, t=0, b=0), height=520,
            coloraxis_colorbar=dict(
                title=dict(text=metric_label, font=dict(size=11, color="#475569")),
                tickfont=dict(size=11, color="#475569"),
            ),
        )
        return fig
    except Exception:
        d = state_values.sort_values("value", ascending=True)
        fig.add_bar(
            x=d["value"], y=d["state"], orientation="h", marker_color=_PRIMARY,
            texttemplate="%{x:,.0f}", textposition="outside",
            textfont=dict(size=11, color="#334155"), cliponaxis=False,
        )
        fig.update_layout(title=dict(
            text="State map unavailable offline — showing ranking",
            font=dict(size=12, color="#64748B"),
        ))
        return _apply_jsw_layout(fig, height=520, x_title=metric_label, show_legend=False)


def be_trajectory(daily_map: dict[str, float], be_total: float,
                  month_label: str, days_in_month: int) -> go.Figure:
    """Cumulative invoiced MT vs straight-line BE pace."""
    fig = go.Figure()
    days = list(range(1, days_in_month + 1))
    by_day = {int(k.split("-")[2]): v for k, v in daily_map.items()}
    cum, running = [], 0.0
    for d in days:
        running += by_day.get(d, 0.0)
        cum.append(running)
    pace = [be_total * d / days_in_month for d in days]
    fig.add_scatter(
        x=days, y=cum, mode="lines+markers", name="Cumulative invoiced",
        line=dict(color="#10B981", width=2.5), marker=dict(size=6),
        fill="tozeroy", fillcolor="rgba(16,185,129,.10)",
    )
    fig.add_scatter(
        x=days, y=pace, mode="lines", name="BE pace",
        line=dict(color="#94A3B8", width=2, dash="dash"),
    )
    fig.update_layout(hovermode="x unified")
    return _apply_jsw_layout(fig, x_title=f"Day of {month_label}", y_title="MT")


def _empty(fig: go.Figure, msg: str) -> go.Figure:
    fig.add_annotation(
        text=msg, showarrow=False,
        font=dict(family=_FONT_FAMILY, size=13, color="#94A3B8"),
    )
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return _apply_jsw_layout(fig, show_legend=False)
