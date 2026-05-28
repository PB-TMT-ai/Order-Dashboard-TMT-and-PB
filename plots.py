"""Plotly figures for the dashboard (port of the Chart.js / SVG-map visuals).

All visual choices come from `theme.py` (JSW palette, fonts, gridlines, etc.).
Line charts default to legend-driven hover isolation: click a legend entry
to focus on a single trace; double-click to restore all.
"""
from __future__ import annotations

from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from data import CHANNEL_LABELS, MOS
from theme import (
    CHANNEL_COLORS, GAP_GRADIENT, JSW_NAVY, JSW_RED, MUTED_TEXT,
    RANKING_DIST, RANKING_STATES, SEQ_GRADIENT, isolate_on_hover,
)

# India states GeoJSON (Datameet, ISO names). Loaded lazily; map degrades to a
# bar chart if the network is unavailable.
INDIA_GEOJSON_URL = (
    "https://raw.githubusercontent.com/geohacker/india/master/state/india_telengana.geojson"
)

# Semi-transparent fills under line traces (per channel — must align with
# CHANNEL_COLORS hue).
_CHANNEL_FILL = {
    "rt": "rgba(99,102,241,.08)",     # indigo
    "ss": "rgba(16,185,129,.08)",     # emerald
    "pdir": "rgba(245,158,11,.08)",   # orange
    "pd": "rgba(139,92,246,.08)",     # purple
}


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


def channel_trend(df: pd.DataFrame, gran: str = "month") -> go.Figure:
    """Ordered MT over time, split by channel (line chart only)."""
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
        fig.add_scatter(
            x=keys, y=series.values, mode="lines+markers", name=label,
            line=dict(color=CHANNEL_COLORS[ch], width=2.5, shape="spline",
                      smoothing=0.8),
            marker=dict(size=6),
            fill="tozeroy", fillcolor=_CHANNEL_FILL[ch],
            hovertemplate=f"<b>{label}</b><br>%{{x}}<br>%{{y:,.0f}} MT<extra></extra>",
        )
    fig.update_layout(height=340, yaxis_title="Ordered MT",
                      legend=dict(orientation="h", y=-0.18))
    return isolate_on_hover(fig)


def _hbar(df: pd.DataFrame, label_col: str, value_col: str, title: str,
          color: str = JSW_NAVY) -> go.Figure:
    fig = go.Figure()
    if not len(df):
        return _empty(fig, "No data")
    d = df.sort_values(value_col, ascending=True)
    fig.add_bar(
        x=d[value_col], y=d[label_col], orientation="h", marker_color=color,
        texttemplate="%{x:,.0f}", textposition="outside",
        textfont=dict(size=11, color=MUTED_TEXT), cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>%{x:,.0f} MT<extra></extra>",
    )
    fig.update_layout(height=320, xaxis_title=title, showlegend=False)
    return fig


def top_states(df: pd.DataFrame, n: int = 10) -> go.Figure:
    agg = (df.groupby("_st")["_q"].sum().reset_index()
           .sort_values("_q", ascending=False).head(n)) if len(df) else pd.DataFrame()
    return _hbar(agg, "_st", "_q", "Ordered MT", RANKING_STATES)


def top_distributors(df: pd.DataFrame, n: int = 10) -> go.Figure:
    agg = (df.groupby("_dn")["_q"].sum().reset_index()
           .sort_values("_q", ascending=False).head(n)) if len(df) else pd.DataFrame()
    return _hbar(agg, "_dn", "_q", "Ordered MT", RANKING_DIST)


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
        textfont=dict(size=11),
        hovertemplate="<b>%{label}</b><br>%{value:,.0f} MT (%{percent})<extra></extra>",
    )
    fig.update_layout(height=320, legend=dict(orientation="v", font=dict(size=10)))
    return fig


def grade_mix(df: pd.DataFrame) -> go.Figure:
    return _pie(df, "_gr", "_q")


def plant_mix(df: pd.DataFrame) -> go.Figure:
    return _pie(df, "_cm", "_q")


def india_map(state_values: pd.DataFrame, metric_label: str,
              gradient: list | None = None) -> go.Figure:
    """Choropleth of a metric by state. Falls back to a bar chart offline."""
    fig = go.Figure()
    if not len(state_values):
        return _empty(fig, "No data")
    if gradient is None:
        gradient = SEQ_GRADIENT
    try:
        import json
        import urllib.request

        with urllib.request.urlopen(INDIA_GEOJSON_URL, timeout=8) as resp:
            geojson = json.load(resp)
        feature_key = "properties.NAME_1"
        fig = px.choropleth(
            state_values, geojson=geojson, locations="state",
            featureidkey=feature_key, color="value",
            color_continuous_scale=gradient,
        )
        fig.update_traces(
            hovertemplate="<b>%{location}</b><br>%{z:,.0f}<extra></extra>",
        )
        fig.update_geos(fitbounds="locations", visible=False, bgcolor="rgba(0,0,0,0)")
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=520,
                          coloraxis_colorbar=dict(title=metric_label,
                                                  thickness=14, len=0.7))
        return fig
    except Exception:
        d = state_values.sort_values("value", ascending=True)
        fig.add_bar(x=d["value"], y=d["state"], orientation="h",
                    marker_color=JSW_NAVY)
        fig.update_layout(height=520, xaxis_title=metric_label,
                          title="State map unavailable offline — showing ranking")
        return fig


def india_gap_map(state_values: pd.DataFrame, mode: str = "abs") -> go.Figure:
    """Diverging choropleth: BE vs Actuals gap by state.

    `state_values` columns: state, value (gap), be, actuals.
    `mode`: "abs" (MT gap, auto-scaled) or "pct" (% gap, -100..+100).
    """
    fig = go.Figure()
    if not len(state_values):
        return _empty(fig, "No data")
    # Symmetric range so 0 = white in the diverging gradient
    vmax = max(abs(state_values["value"].min()),
               abs(state_values["value"].max()), 1.0)
    label = "Gap %" if mode == "pct" else "Gap MT"
    try:
        import json
        import urllib.request

        with urllib.request.urlopen(INDIA_GEOJSON_URL, timeout=8) as resp:
            geojson = json.load(resp)
        fig = px.choropleth(
            state_values, geojson=geojson, locations="state",
            featureidkey="properties.NAME_1", color="value",
            color_continuous_scale=GAP_GRADIENT,
            range_color=(-vmax, vmax),
            custom_data=["be", "actuals"],
        )
        fmt = "%{z:+.1f}%" if mode == "pct" else "%{z:+,.0f} MT"
        fig.update_traces(
            hovertemplate=("<b>%{location}</b><br>"
                           f"Gap: {fmt}<br>"
                           "BE: %{customdata[0]:,.0f}<br>"
                           "Actuals: %{customdata[1]:,.0f}<extra></extra>"),
        )
        fig.update_geos(fitbounds="locations", visible=False, bgcolor="rgba(0,0,0,0)")
        fig.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=520,
                          coloraxis_colorbar=dict(title=label,
                                                  thickness=14, len=0.7))
        return fig
    except Exception:
        d = state_values.sort_values("value", ascending=True)
        fig.add_bar(x=d["value"], y=d["state"], orientation="h",
                    marker_color=[JSW_RED if v < 0 else "#10B981"
                                  for v in d["value"]])
        fig.update_layout(height=520, xaxis_title=label,
                          title="State map unavailable offline — showing ranking")
        return fig


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
        line=dict(color="#10B981", width=2.5, shape="spline", smoothing=0.8),
        marker=dict(size=6),
        fill="tozeroy", fillcolor="rgba(16,185,129,.10)",
        hovertemplate="Day %{x}<br>%{y:,.0f} MT<extra>Cumulative</extra>",
    )
    fig.add_scatter(
        x=days, y=pace, mode="lines", name="BE pace",
        line=dict(color=JSW_NAVY, width=2, dash="dash"),
        hovertemplate="Day %{x}<br>%{y:,.0f} MT<extra>BE pace</extra>",
    )
    fig.update_layout(height=340, xaxis_title=f"Day of {month_label}",
                      yaxis_title="MT", legend=dict(orientation="h", y=-0.18))
    return isolate_on_hover(fig)


def _empty(fig: go.Figure, msg: str) -> go.Figure:
    fig.add_annotation(text=msg, showarrow=False,
                       font=dict(size=13, color=MUTED_TEXT))
    fig.update_layout(height=300, xaxis=dict(visible=False),
                      yaxis=dict(visible=False))
    return fig
