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
    """Populate and open the drawer.

    Args:
      title: header text
      df: rows to display + export
      subtitle: small grey caption under title
      summary: list of (label, value) chips at the top
      filename: CSV download name
      context: arbitrary state the consumer wants to thread through (e.g. for
        the contents to render a chart inside the drawer)
    """
    s = _state()
    s.open = True
    s.title = title
    s.subtitle = subtitle
    s.summary = summary or []
    s.df = df
    s.filename = filename
    s.context = context or {}


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
.dr-backdrop {{
  position: fixed; inset: 0; background: rgba(15, 23, 42, 0.32);
  z-index: 9990; backdrop-filter: blur(2px);
  animation: dr-fade 180ms ease-out;
}}
.dr-panel {{
  position: fixed; top: 0; right: 0; height: 100vh; width: min(560px, 92vw);
  background: {CARD_BG}; z-index: 9999;
  box-shadow: -8px 0 24px rgba(15, 23, 42, 0.18);
  border-left: 4px solid {JSW_NAVY};
  display: flex; flex-direction: column;
  animation: dr-slide 260ms cubic-bezier(.22,1,.36,1);
}}
@keyframes dr-slide {{ from {{ transform: translateX(100%); }} to {{ transform: translateX(0); }} }}
@keyframes dr-fade {{ from {{ opacity: 0; }} to {{ opacity: 1; }} }}
.dr-head {{
  padding: 18px 22px 12px; border-bottom: 1px solid {GRID};
  background: linear-gradient(180deg, #FAFBFC 0%, {CARD_BG} 100%);
}}
.dr-title {{ font-size: 16px; font-weight: 700; color: {INK}; line-height: 1.3; }}
.dr-sub {{ font-size: 12px; color: {MUTED_TEXT}; margin-top: 4px; }}
.dr-summary {{
  display: flex; flex-wrap: wrap; gap: 8px; margin-top: 12px;
}}
.dr-chip {{
  background: #F1F5F9; border-radius: 6px; padding: 6px 10px;
  font-size: 11px; color: {MUTED_TEXT}; flex: 0 0 auto;
}}
.dr-chip b {{ display: block; font-size: 14px; color: {JSW_NAVY};
              font-weight: 700; margin-top: 1px; }}
</style>
"""


def render() -> None:
    """Render the global drawer (call once at the bottom of every page)."""
    s = _state()
    if not s.open:
        return

    st.markdown(_DRAWER_CSS, unsafe_allow_html=True)
    st.markdown('<div class="dr-backdrop"></div>', unsafe_allow_html=True)

    # Native Streamlit dialog (right-side panel via CSS overrides). We use
    # st.dialog when available (Streamlit ≥ 1.32); fall back to expander.
    has_dialog = hasattr(st, "dialog")
    if has_dialog:
        _render_dialog()
    else:
        _render_fallback()


def _summary_html(s: DrawerState) -> str:
    if not s.summary:
        return ""
    chips = "".join(
        f'<div class="dr-chip">{label}<b>{value}</b></div>'
        for label, value in s.summary
    )
    return f'<div class="dr-summary">{chips}</div>'


def _render_dialog() -> None:
    s = _state()

    @st.dialog(s.title, width="large")
    def _dlg() -> None:
        if s.subtitle:
            st.caption(s.subtitle)
        if s.summary:
            st.markdown(_DRAWER_CSS + _summary_html(s), unsafe_allow_html=True)
        if s.df is not None and len(s.df):
            st.download_button(
                "⇩ Download CSV", s.df.to_csv(index=False).encode("utf-8"),
                file_name=s.filename, mime="text/csv",
                key=f"dr_dl_{s.title}",
            )
            st.dataframe(s.df, use_container_width=True, hide_index=True,
                         height=420)
        elif s.df is not None:
            st.info("No rows.")
        c1, _ = st.columns([1, 4])
        if c1.button("Close", key="dr_close", type="primary"):
            close_drawer()
            st.rerun()

    _dlg()


def _render_fallback() -> None:
    s = _state()
    with st.expander(s.title, expanded=True):
        if s.subtitle:
            st.caption(s.subtitle)
        if s.summary:
            st.markdown(_summary_html(s), unsafe_allow_html=True)
        if s.df is not None and len(s.df):
            st.download_button(
                "⇩ Download CSV", s.df.to_csv(index=False).encode("utf-8"),
                file_name=s.filename, mime="text/csv",
                key=f"dr_dl_fb_{s.title}",
            )
            st.dataframe(s.df, use_container_width=True, hide_index=True)
        if st.button("Close", key="dr_close_fb"):
            close_drawer()
            st.rerun()
