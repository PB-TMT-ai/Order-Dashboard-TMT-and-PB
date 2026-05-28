"""JSW One brand theme — colors, Plotly defaults, gradient helpers.

Centralizes every visual choice so charts/maps/UI render consistently.
Imported by plots.py, app.py, drawer.py.
"""
from __future__ import annotations

import plotly.graph_objects as go
import plotly.io as pio

# ─── JSW One brand palette (from logo) ───────────────────────────────────────
JSW_NAVY = "#002E5D"     # primary — deep navy from "JSW ONE" wordmark
JSW_RED = "#ED1C24"      # accent — red from swoosh and "MSME"
JSW_NAVY_DARK = "#001A35"
JSW_NAVY_LIGHT = "#1F4E8C"
JSW_RED_LIGHT = "#F25C61"

# ─── Channel colors (distinct, accessible) ───────────────────────────────────
CHANNEL_COLORS = {
    "rt": JSW_NAVY,           # Retail — primary navy
    "ss": "#F59E0B",          # Self-stocking — amber
    "pdir": "#10B981",        # Project (direct) — emerald
    "pd": JSW_RED,            # Project thru Dist — JSW red
}

# ─── KPI card border colors (matches channel colors plus extras) ─────────────
KPI_BORDER = {
    "or": JSW_NAVY,          # Ordered
    "re": "#F59E0B",         # Released
    "in": "#10B981",         # Invoiced
    "inp": "#0EA5E9",        # Invoiced in period
    "gap": JSW_RED,          # BE gap
}

# ─── Diverging gradient for BE-vs-Actuals (red = under, green = over) ────────
# Use proper red→white→green progression centered at 0
GAP_GRADIENT = [
    [0.0, JSW_RED],          # worst underperformer
    [0.25, "#F87171"],
    [0.5, "#F9FAFB"],         # at pace
    [0.75, "#86EFAC"],
    [1.0, "#059669"],        # best overperformer
]

# ─── Sequential gradient for choropleth (light → JSW navy) ───────────────────
SEQ_GRADIENT = ["#EEF2F8", "#B8C9E0", "#7991BD", "#3D5C9A", JSW_NAVY]

# ─── Neutral / structural colors ─────────────────────────────────────────────
GRID = "#E5E7EB"
MUTED_TEXT = "#64748B"
DIM_TEXT = "#94A3B8"
INK = "#1E293B"
CARD_BG = "#FFFFFF"
PAGE_BG = "#FAFBFC"


# ─── Plotly default template (applied via pio.templates.default) ─────────────
def _build_template() -> go.layout.Template:
    """Custom Plotly template — JSW palette + tight, clean defaults."""
    t = go.layout.Template()
    t.layout = go.Layout(
        font=dict(family="-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
                  size=12, color=INK),
        colorway=[JSW_NAVY, JSW_RED, "#10B981", "#F59E0B", "#0EA5E9",
                  "#8B5CF6", "#EC4899", "#14B8A6"],
        paper_bgcolor=CARD_BG,
        plot_bgcolor=CARD_BG,
        margin=dict(l=12, r=12, t=12, b=12),
        xaxis=dict(gridcolor=GRID, zerolinecolor=GRID, linecolor=GRID,
                   tickfont=dict(size=11, color=MUTED_TEXT)),
        yaxis=dict(gridcolor=GRID, zerolinecolor=GRID, linecolor=GRID,
                   tickfont=dict(size=11, color=MUTED_TEXT), tickformat=".2s"),
        legend=dict(orientation="h", y=-0.18, x=0, font=dict(size=11)),
        hoverlabel=dict(bgcolor=CARD_BG, bordercolor=JSW_NAVY,
                        font=dict(size=12, color=INK)),
    )
    return t


pio.templates["jsw"] = _build_template()
pio.templates.default = "jsw"


def isolate_on_hover(fig: go.Figure) -> go.Figure:
    """Add hover-isolation behaviour to a line figure.

    On hover over a trace, that trace stays at full opacity while every
    other trace fades to grey. Uses Plotly's built-in `hoveron` interaction
    by reducing non-hovered opacity via a JS-side callback isn't possible
    in pure Streamlit; instead we set up traces so the legend doubles as
    a single-line isolator (click a legend entry to isolate).
    """
    fig.update_layout(
        legend=dict(itemclick="toggleothers", itemdoubleclick="toggle"),
        hovermode="x unified",
    )
    return fig
