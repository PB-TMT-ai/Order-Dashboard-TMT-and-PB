"""Right-side slide-over drawer for the dashboard.

A single global drawer surface, driven by `st.session_state.drawer`. Any
KPI / chart / table can call `open_drawer(...)` to populate it; the
universal render at the bottom of every page draws it.

The drawer renders as a fixed-position panel injected via raw HTML/CSS
(Streamlit doesn't ship a native slide-over). Open/close state lives in
session state and a re-run cycle handles the transition.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

import pandas as pd
import streamlit as st

from theme import CARD_BG, GRID, INK, JSW_NAVY, JSW_RED, MUTED_TEXT


# ─── State container ─────────────────────────────────────────────────────────
@dataclass
class DrawerState:
    open: bool = False
    title: str = ""
    subtitle: str = ""
    summary: list[tuple[str, str]] = field(default_factory=list)
    df: pd.DataFrame | None = None
    filename: str = "drawer_rows.csv"
    context: dict[str, Any] = field(default_factory=dict)


def _state() -> DrawerState:
    if "drawer" not in st.session_state:
        st.session_state.drawer = DrawerState()
    return st.session_state.drawer


def open_drawer(title: str, df: pd.DataFrame | None = None, *,
                subtitle: str = "", summary: list[tuple[str, str]] | None = None,
                filename: str = "drawer_rows.csv",
                context: dict | None = None) -> None:
    """Populate state, then directly invoke the dialog in the same script run.

    Calling `_drawer_dialog()` here (instead of just flipping a flag and
    running st.rerun()) is what makes Streamlit actually open the modal — a
    dialog must be invoked during the same script run as the user
    interaction that triggers it.
    """
    s = _state()
    s.open = True
    s.title = title
    s.subtitle = subtitle
    s.summary = summary or []
    s.df = df
    s.filename = filename
    s.context = context or {}
    _drawer_dialog()


def close_drawer() -> None:
    s = _state()
    s.open = False
    s.title = ""
    s.subtitle = ""
    s.summary = []
    s.df = None
    s.context = {}


# ─── Trigger button helper ───────────────────────────────────────────────────
def trigger(label: str, key: str, *, title: str, df: pd.DataFrame | None = None,
            subtitle: str = "", summary: list[tuple[str, str]] | None = None,
            filename: str = "drawer_rows.csv") -> bool:
    """Render a small 🔍 button. Returns True the run it's clicked."""
    if st.button(label, key=key, help=f"View underlying data: {title}"):
        open_drawer(title, df, subtitle=subtitle, summary=summary,
                    filename=filename)
        st.rerun()
        return True
    return False


# ─── Universal CSS (injected once) ───────────────────────────────────────────
_DRAWER_CSS = f"""
<style>
/* Re-style Streamlit's native dialog as a right-side slide-over panel.
   Streamlit wraps the modal in: [data-testid="stDialog"] > div > [role=dialog] */

/* Right-anchor the inner role=dialog element */
[data-testid="stDialog"] [role="dialog"] {{
    position: fixed !important;
    top: 0 !important; right: 0 !important; left: auto !important;
    bottom: 0 !important;
    transform: none !important;
    height: 100vh !important; max-height: 100vh !important;
    width: min(640px, 95vw) !important;
    max-width: 95vw !important;
    margin: 0 !important;
    border-radius: 0 !important;
    border-left: 4px solid {JSW_NAVY} !important;
    box-shadow: -10px 0 30px rgba(15,23,42,.18) !important;
    animation: dr-slide 280ms cubic-bezier(.22,1,.36,1);
    overflow-y: auto;
}}
/* Also right-anchor the inner wrapper that BaseWeb modal uses */
[data-testid="stDialog"] > div {{
    justify-content: flex-end !important;
}}
@keyframes dr-slide {{
    from {{ transform: translateX(100%); }}
    to   {{ transform: translateX(0); }}
}}
/* Hide BaseWeb's default header (it shows 'Detail' from @st.dialog) */
[data-testid="stDialog"] [role="dialog"] > div:first-child {{
    display: none !important;
}}
.dr-summary {{
  display: flex; flex-wrap: wrap; gap: 8px; margin: 10px 0 14px;
}}
.dr-chip {{
  background: #F1F5F9; border-radius: 6px; padding: 6px 10px;
  font-size: 11px; color: {MUTED_TEXT}; flex: 0 0 auto;
}}
.dr-chip b {{ display: block; font-size: 14px; color: {JSW_NAVY};
              font-weight: 700; margin-top: 1px; }}
</style>
"""


def _summary_html(s: DrawerState) -> str:
    if not s.summary:
        return ""
    chips = "".join(
        f'<div class="dr-chip">{label}<b>{value}</b></div>'
        for label, value in s.summary
    )
    return f'<div class="dr-summary">{chips}</div>'


# Module-level dialog function — Streamlit's @st.dialog requires that.
# The actual title is rendered as an H3 inside the body so it can be dynamic.
@st.dialog("Detail", width="large")
def _drawer_dialog() -> None:
    s = _state()
    st.markdown(_DRAWER_CSS, unsafe_allow_html=True)
    if s.title:
        st.markdown(
            f'<div style="font-size:18px;font-weight:700;color:#0F172A;'
            f'margin:-12px 0 4px;">{s.title}</div>',
            unsafe_allow_html=True)
    if s.subtitle:
        st.caption(s.subtitle)
    if s.summary:
        st.markdown(_summary_html(s), unsafe_allow_html=True)
    if s.df is not None and len(s.df):
        c1, c2 = st.columns([1, 4])
        with c1:
            st.download_button(
                "⇩ CSV", s.df.to_csv(index=False).encode("utf-8"),
                file_name=s.filename, mime="text/csv",
                key="dr_dl", use_container_width=True)
        st.dataframe(s.df, use_container_width=True, hide_index=True,
                     height=460)
    elif s.df is not None:
        st.info("No rows match.")
    if st.button("Close", key="dr_close", type="primary"):
        close_drawer()
        st.rerun()


def render() -> None:
    """Render the global drawer (call once at the bottom of every page)."""
    s = _state()
    if not s.open:
        return
    _drawer_dialog()
