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

# ─── Channel colors (distinct, accessible — rainbow palette for 4 series) ───
# JSW navy/red are reserved as accents on borders/KPIs/buttons; charts use
# this 4-color rainbow so the channel lines stay distinguishable.
CHANNEL_COLORS = {
    "rt": "#6366F1",          # Retail — indigo
    "ss": "#10B981",          # Self-stocking — emerald
    "pdir": "#F59E0B",        # Project (direct) — orange
    "pd": "#8B5CF6",          # Project thru Dist — purple
}

# Top-N bar colors (single-series accents)
RANKING_STATES = "#6366F1"     # Top-10 ship-to states — indigo
RANKING_DIST = "#10B981"       # Top-10 distributors — emerald

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


# ─── Card / pill control CSS injected once per page ─────────────────────────
CARD_CSS = """
<style>
/* Style st.radio (horizontal) as pill segmented control */
div[data-testid="stRadio"][role="radiogroup"] > div,
div[data-testid="stRadio"] > div[role="radiogroup"] {
    background:#F1F5F9; border:1px solid #E2E8F0; border-radius:10px;
    padding:3px; display:inline-flex; gap:2px;
}
div[data-testid="stRadio"] label {
    background:transparent; border-radius:7px; padding:6px 14px;
    margin:0 !important; font-size:12px; font-weight:600; color:#64748B;
    cursor:pointer; transition:all .15s ease; border:none;
}
div[data-testid="stRadio"] label:hover { color:#0F172A; background:#FFFFFF; }
div[data-testid="stRadio"] label > div:first-child { display:none !important; }
div[data-testid="stRadio"] label[data-checked="true"],
div[data-testid="stRadio"] label:has(input:checked) {
    background:#002E5D !important; color:#FFFFFF !important;
    box-shadow:0 1px 3px rgba(0,46,93,.25);
}

/* CSV download button — small green pill */
div[data-testid="stDownloadButton"] > button {
    background:#ECFDF5; color:#059669; border:1px solid #A7F3D0;
    border-radius:8px; padding:5px 14px; font-size:12px; font-weight:600;
    height:32px; min-height:32px; line-height:1;
}
div[data-testid="stDownloadButton"] > button:hover {
    background:#D1FAE5; border-color:#6EE7B7; color:#047857;
}

/* Section card — st.container(border=True) overrides */
div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius:12px !important; border-color:#E2E8F0 !important;
    box-shadow:0 1px 3px rgba(15,23,42,.04);
    background:#FFFFFF;
}

/* Chart card title + subtitle spacing */
.chart-title { font-size:15px; font-weight:700; color:#0F172A;
               margin:0 0 2px 0; letter-spacing:-.01em; }
.chart-sub   { font-size:12px; color:#64748B; margin:0 0 8px 0;
               line-height:1.4; }
</style>
"""


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
