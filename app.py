"""JSW Order Intelligence Dashboard — Streamlit entry point.

Sidebar filters, KPI strip, and tabs. The data layer (data.py / be_logic.py) is
complete; some tabs are marked TODO for the iteration session (see README).
"""
from __future__ import annotations

import hashlib
import html
import io
from datetime import datetime

import pandas as pd
import streamlit as st

import be_logic
import data
import drawer
import plots
import storage_io as storage
import theme  # noqa: F401 — registers the JSW Plotly template

st.set_page_config(page_title="JSW One — Order Intelligence", layout="wide")


# --------------------------------------------------------------------------- #
# Data loading
# --------------------------------------------------------------------------- #
# 'calamine' (Rust) reads .xlsx far faster and with much less memory than
# openpyxl — essential for large (tens of MB) exports.
def _excel_file(file_bytes: bytes):
    try:
        return pd.ExcelFile(io.BytesIO(file_bytes), engine="calamine")
    except Exception:
        return pd.ExcelFile(io.BytesIO(file_bytes))


_ORDER_COLS = set(data.K.values())
_INVOICE_COLS = {"Order ID", "Invoice date", "Invoiced qty", "Invoice number"}


def _norm_col(s: str) -> str:
    """Normalise an Excel column header for tolerant matching."""
    return " ".join(str(s).strip().lower().split())


def _read_canonical(xls, sheet: str, wanted: set[str]) -> pd.DataFrame:
    """Read a sheet, keep only the wanted columns (case/space-insensitive match),
    and rename them back to the canonical names the data layer expects.
    """
    wanted_by_norm = {_norm_col(c): c for c in wanted}
    df = pd.read_excel(xls, sheet_name=sheet,
                       usecols=lambda c: _norm_col(c) in wanted_by_norm)
    return df.rename(columns={c: wanted_by_norm[_norm_col(c)] for c in df.columns})


def parse_order_workbook(file_bytes: bytes):
    """Parse the Excel and return (enriched_df, inv_index, raw_invoice_df).

    The raw invoice DataFrame is returned so it can be persisted alongside the
    enriched orders as Parquet — keeping enough state to rebuild inv_index on
    later sessions without re-parsing the Excel.

    Column names are matched case- and space-insensitively (so e.g.
    "Opportunity Date" vs "Opportunity date" resolves), and only the columns
    the dashboard uses are loaded to keep memory in check.
    """
    xls = _excel_file(file_bytes)
    inv_index = {}
    inv_df = None
    inv_sheet = next((s for s in xls.sheet_names if s.lower() == "invoice"), None) \
        or next((s for s in xls.sheet_names if "invoice" in s.lower()), None)
    if inv_sheet:
        inv_df = _read_canonical(xls, inv_sheet, _INVOICE_COLS)
        inv_index = data.build_invoice_index(inv_df)

    order_sheet = next((s for s in xls.sheet_names if "order" in s.lower()), xls.sheet_names[0])
    order_df = _read_canonical(xls, order_sheet, _ORDER_COLS)
    enriched = data.enrich(order_df, inv_index)
    del order_df
    return enriched, inv_index, inv_df


@st.cache_data(show_spinner="Reading BE sheet names...")
def list_be_sheets(file_bytes: bytes) -> list[str]:
    return _excel_file(file_bytes).sheet_names


@st.cache_data(show_spinner=False)
def parse_be(file_bytes: bytes, sheet: str) -> list[dict]:
    try:
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None,
                            engine="calamine")
    except Exception:
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None)
    aoa = raw.where(pd.notna(raw), "").values.tolist()
    return be_logic.parse_be_sheet(aoa)


@st.cache_data(ttl=30, show_spinner=False)
def _storage_status() -> tuple[bool, str]:
    return storage.check_connection()


# --------------------------------------------------------------------------- #
# Session state
# --------------------------------------------------------------------------- #
if "be_version" not in st.session_state:
    st.session_state.be_version = None
if "be_restore_tried" not in st.session_state:
    st.session_state.be_restore_tried = False


def restore_be(month: str | None = None, week: str | None = None):
    """Reconstruct a saved BE version from cloud storage.

    If month+week given, load that exact slot. Otherwise pick the latest
    versioned slot (latest week of latest month), then fall back to the
    legacy single-slot file.
    """
    if month and week:
        be_bytes, meta = storage.load_be_version(month, week)
        meta = meta or {}
    else:
        latest = storage.latest_be_version()
        if latest:
            be_bytes, meta = storage.load_be_version(latest["month"], latest["week"])
            meta = meta or {"sheet": latest.get("sheet"), "month": latest["month"],
                            "week": latest["week"], "uploaded": latest.get("uploaded", "")}
        else:
            # Legacy single-BE fallback
            be_bytes = storage.load_be_file()
            meta = storage.load_be_meta() or {}
    if not be_bytes or not meta.get("sheet") or not meta.get("month"):
        return None
    rows = parse_be(be_bytes, meta["sheet"])
    if not rows:
        return None
    label, y, m0 = be_logic.month_label_from_value(meta["month"])
    return be_logic.BeVersion(
        upload_date=meta.get("uploaded", ""), month_label=label, month_y=y, month_m=m0,
        week=meta.get("week", ""), sheet=meta["sheet"], rows=rows,
        total_be=sum(r["total"] for r in rows))


def get_data():
    """Return (df, inv_index), parsed once per file and cached in session_state.

    Parsing and the (potentially large) cloud upload run only when the file
    actually changes — not on every rerun.
    """
    up = st.session_state.get("order_upload")
    if up is not None:
        file_bytes = up.getvalue()
        h = hashlib.md5(file_bytes).hexdigest()
        if st.session_state.get("order_hash") != h:
            with st.spinner("Parsing order file…"):
                enriched, inv_index, inv_df = parse_order_workbook(file_bytes)
            st.session_state.order_data = (enriched, inv_index)
            st.session_state.order_hash = h
            st.session_state.order_size = len(file_bytes)
            if storage.is_configured():
                with st.spinner("Saving processed data to cloud storage…"):
                    ok = storage.save_processed(enriched, inv_df)
                st.session_state.order_saved = ok
                st.session_state.order_save_err = None if ok else storage.last_error()
        return st.session_state.order_data

    # No file in the uploader: load the saved processed copy from storage.
    # Parquet is tiny + fast — no Excel re-parsing, no OOM risk.
    if "order_data" not in st.session_state:
        if storage.is_configured():
            with st.spinner("Loading saved data…"):
                orders, invoices = storage.load_processed()
        else:
            orders, invoices = None, None
        if orders is not None and len(orders):
            inv_index = data.build_invoice_index(invoices) if invoices is not None \
                else {}
            st.session_state.order_data = (orders, inv_index)
        else:
            st.session_state.order_data = (None, None)
    return st.session_state.order_data


# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #
hdr_l, hdr_r = st.columns([3, 1])
with hdr_l:
    st.title("JSW One — Order Intelligence")
with hdr_r:
    if storage.is_configured():
        ok, msg = _storage_status()
        (st.success if ok else st.error)(msg)
    else:
        st.caption("Cloud storage: not configured (no persistence)")

st.file_uploader("Load order Excel (.xlsx)", type=["xlsx", "xls"], key="order_upload")

df, inv_index = get_data()

# Persistence status — tells you whether colleagues will see this file
if storage.is_configured() and st.session_state.get("order_saved") is not None:
    if st.session_state.order_saved:
        st.caption("✅ Saved to cloud storage — colleagues opening the app will load "
                   "this file automatically.")
    else:
        size_mb = st.session_state.get("order_size", 0) / 1_000_000
        st.warning(
            f"⚠️ Loaded for you, but **could not save to cloud storage** (your "
            f"colleagues will still see the upload prompt). File is {size_mb:.0f} MB.\n\n"
            f"Error: `{st.session_state.get('order_save_err')}`")

if df is None or not len(df):
    st.info("Upload your Excel file. The dashboard reads the **Order** sheet "
            "(and an **Invoice** sheet if present) and builds all metrics.")
    st.stop()

# Restore the last-saved BE once per session (after order data is available)
if (st.session_state.be_version is None and not st.session_state.be_restore_tried
        and storage.is_configured()):
    st.session_state.be_restore_tried = True
    st.session_state.be_version = restore_be()


# --------------------------------------------------------------------------- #
# Sidebar filters
# --------------------------------------------------------------------------- #
def build_filters(df: pd.DataFrame) -> dict:
    f = data.default_filters()
    sb = st.sidebar
    sb.header("Filters")

    f["pt"] = sb.radio("Product type", ["All", "TMT", "P&T"], horizontal=True)

    dm = sb.radio("Opportunity date", ["All / Last X", "Month", "Year", "Range"],
                  horizontal=True)
    if dm == "All / Last X":
        choice = sb.select_slider(
            "Window", options=["7", "15", "30", "60", "90", "180", "365", "All"],
            value="All")
        f["dm"] = "lx"
        f["lx"] = 9999 if choice == "All" else int(choice)
    elif dm == "Month":
        f["dm"] = "mo"
        f["mo"] = sb.text_input("Month (YYYY-MM)", value="")
    elif dm == "Year":
        years = sorted(y for y in df["_y"].unique() if y)
        sel = sb.selectbox("Year", ["All years"] + years)
        f["dm"] = "yr"
        f["yr"] = "" if sel == "All years" else sel
    else:
        f["dm"] = "rg"
        c1, c2 = sb.columns(2)
        f["df"] = c1.text_input("From (YYYY-MM-DD)", value="")
        f["dt"] = c2.text_input("To (YYYY-MM-DD)", value="")

    f["dis"] = sb.radio("Distributor", ["All", "Yes", "No"], horizontal=True)

    labels = {
        "ot": "Order type", "dn": "Specific distributor", "sts": "Ship to state",
        "stc": "Ship to city", "bts": "Bill to state", "gr": "Grade",
        "dia": "Diameter (mm)", "frm": "Form", "ptm": "Payment terms",
        "dlm": "Delivery method", "cm": "CM / Dispatch plant",
        "ws": "Warehouse / shop / site", "stat": "Order status",
    }
    for key, col in data.MULTI_FILTERS.items():
        options = sorted(v for v in df[col].dropna().unique() if v)
        if key == "dia":
            options = sorted((v for v in df[col].dropna().unique() if v), key=data.num)
        f[key] = sb.multiselect(labels[key], options, default=[])
    return f


f = build_filters(df)
filtered = data.apply_filters(df, f)


# --------------------------------------------------------------------------- #
# KPI strip
# --------------------------------------------------------------------------- #
_KPI_CSS = """
<style>
.kpi-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
    gap:12px;margin-bottom:12px}
.kc{background:#fff;border-radius:10px;padding:14px 16px;
    box-shadow:0 2px 8px rgba(15,23,42,.06);
    border-top:3px solid #002E5D;border-left:4px solid transparent;
    transition:transform .15s ease, box-shadow .15s ease}
.kc:hover{transform:translateY(-1px);box-shadow:0 4px 12px rgba(15,23,42,.08)}
.kc.k-or{border-left-color:#002E5D}
.kc.k-re{border-left-color:#F59E0B}
.kc.k-in{border-left-color:#10B981}
.kc.k-inp{border-left-color:#0EA5E9}
.kc.k-gap{border-left-color:#ED1C24}
.kl{font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;
    letter-spacing:.5px;margin-bottom:4px}
.kgaplbl{font-weight:400;color:#94A3B8;font-size:10px;text-transform:none;
    letter-spacing:0}
.kv{font-size:28px;font-weight:700;line-height:1.1;color:#0F172A;
    font-variant-numeric:tabular-nums}
.kv.dn{color:#ED1C24}.kv.up{color:#059669}
.ku{font-size:10px;font-weight:500;color:#94A3B8;margin-left:3px}
.ks{font-size:11px;color:#64748B;margin-top:4px}
.ch-subs{margin-top:6px;border-top:1px dashed #E2E8F0;padding-top:6px}
.ch-sub{display:flex;justify-content:space-between;font-size:10px;padding:2px 0;
    color:#64748B;line-height:1.4}
.ch-sub b{color:#334155;font-weight:600;font-variant-numeric:tabular-nums}
.ibar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px}
.ic{background:#fff;border:1px solid #CBD5E1;border-radius:6px;padding:7px 12px;
    font-size:11px;color:#64748B;box-shadow:0 1px 3px rgba(15,23,42,.05);
    transition:background .15s ease,border-color .15s ease}
.ic:hover{background:#F1F5F9;border-color:#94A3B8}
.iv{font-weight:700;font-size:14px;color:#002E5D;display:block;margin-bottom:1px;
    font-variant-numeric:tabular-nums}
</style>
"""

_CH_ROWS = [("Retail", "rt"), ("Self-stocking", "ss"),
            ("Project (direct)", "pdir"), ("Project thru Dist", "pd")]


def _ch_subs(ch: dict) -> str:
    return "".join(
        f'<div class="ch-sub"><span>{lbl}</span><b>{data.fmt(ch.get(k, 0))}</b></div>'
        for lbl, k in _CH_ROWS)


def _kpi_card(cls: str, label: str, value: str, sub: str,
              ch_html: str = "", value_cls: str = "") -> str:
    return (
        f'<div class="kc {cls}">'
        f'<div class="kl">{label}</div>'
        f'<div class="kv {value_cls}">{value}<span class="ku">MT</span></div>'
        f'<div class="ks">{html.escape(sub)}</div>'
        f'<div class="ch-subs">{ch_html}</div>'
        f'</div>'
    )


kpis = data.compute_kpis(filtered)
period = data.get_active_period(f)

# Invoiced-in-period (invoice-date scoped, ignores order-date filter), per channel
nd = data.apply_non_date_filters(df, f)
ch_inp = {c: 0.0 for c in ("rt", "ss", "pdir", "pd")}
if len(nd):
    if period:
        iip = data.invoiced_in_period(nd, period["from"], period["to"], inv_index)
    else:
        iip = nd["_iq"]
    inv_in_period = float(iip.sum())
    for ch, val in iip.groupby(nd["_ch"]).sum().items():
        ch_inp[ch] = float(val)
else:
    inv_in_period = 0.0

be = st.session_state.be_version
ag = be_logic.be_actuals_agg(filtered, be, inv_index) if be else None

_pct = lambda part: f"{round(part / kpis.ordered * 100) if kpis.ordered else 0}% of ordered"
cards = [
    _kpi_card("k-or", "Ordered", data.fmt(kpis.ordered),
              f"{kpis.line_items:,} line items", _ch_subs(kpis.ch_or)),
    _kpi_card("k-re", "Released", data.fmt(kpis.released),
              _pct(kpis.released), _ch_subs(kpis.ch_re)),
    _kpi_card("k-in", "Invoiced ⓘ", data.fmt(kpis.invoiced),
              _pct(kpis.invoiced), _ch_subs(kpis.ch_in)),
    _kpi_card("k-inp", "Invoiced in period ⓘ", data.fmt(inv_in_period),
              period["label"] if period else "All time", _ch_subs(ch_inp)),
]
if ag:
    gap = ag.matched_act - ag.tot_be  # negative ⇒ behind plan
    wk = f" {be.week}" if be.week else ""
    cards.append(_kpi_card(
        "k-gap", f'BE gap <span class="kgaplbl">({be.month_label}{wk})</span>',
        data.fmt(gap), f"BE {data.fmt(ag.tot_be)} · Inv {data.fmt(ag.matched_act)}",
        value_cls="dn" if gap < 0 else "up"))
else:
    cards.append(_kpi_card("k-gap", "BE gap", "—", "load a BE file"))

st.markdown(_KPI_CSS + '<div class="kpi-row">' + "".join(cards) + "</div>",
            unsafe_allow_html=True)

# Info bar chips
st_sum = filtered.groupby("_st")["_q"].sum() if len(filtered) else None
top_st = st_sum.idxmax() if st_sum is not None and len(st_sum) and st_sum.max() > 0 else None
cm_sum = filtered.groupby("_cm")["_q"].sum() if len(filtered) else None
top_cm = cm_sum.idxmax() if cm_sum is not None and len(cm_sum) and cm_sum.max() > 0 else None
pending = max(kpis.ordered - kpis.released, 0.0)
pend_inv = max(kpis.released - kpis.invoiced, 0.0)
pts = float(filtered.loc[filtered["_pt"] == "P&T", "_q"].sum()) if len(filtered) else 0.0
dist_cnt = int(filtered.loc[filtered["_dis"] == "Yes", "_dn"].nunique()) if len(filtered) else 0

chips = []
if top_st:
    chips.append((f"{data.fmt(st_sum.max())} MT", f"🏆 Top state: {top_st}"))
chips.append((f"{data.fmt(pending)} MT", "⏳ Pending release"))
chips.append((f"{data.fmt(pend_inv)} MT", "📋 Released, not invoiced"))
if top_cm:
    chips.append((" ".join(str(top_cm).split()[:2]), "🏭 Top plant"))
if pts > 0:
    chips.append((f"{data.fmt(pts)} MT", "🔧 P&T volume"))
chips.append((f"{dist_cnt:,}", "🤝 Active distributors"))

st.markdown(
    '<div class="ibar">' + "".join(
        f'<div class="ic"><span class="iv">{html.escape(str(v))}</span>'
        f'{html.escape(lbl)}</div>' for v, lbl in chips) + "</div>",
    unsafe_allow_html=True)


# ─── Inject card / pill control CSS once for the rest of the page ───────────
st.markdown(theme.CARD_CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Chart-card helper
# --------------------------------------------------------------------------- #
def _chart_header(title: str, subtitle: str, *, csv_df: pd.DataFrame | None = None,
                  csv_name: str = "data.csv", key: str = "",
                  controls=None) -> object | None:
    """Render a card header row: title+subtitle (left) | CSV button | controls.

    `controls` (optional) is a callable rendered into the right-most column;
    its return value is passed back so the caller can act on it.
    """
    has_ctrl = controls is not None
    has_csv = csv_df is not None
    if has_ctrl and has_csv:
        ratios = [5, 1, 2]
    elif has_csv or has_ctrl:
        ratios = [6, 1] if has_csv else [4, 2]
    else:
        ratios = [1]
    cols = st.columns(ratios)
    with cols[0]:
        st.markdown(f'<div class="chart-title">{html.escape(title)}</div>',
                    unsafe_allow_html=True)
        if subtitle:
            st.markdown(f'<div class="chart-sub">{html.escape(subtitle)}</div>',
                        unsafe_allow_html=True)
    idx = 1
    if has_csv:
        with cols[idx]:
            st.download_button(
                "⇩ CSV", csv_df.to_csv(index=False).encode("utf-8"),
                file_name=csv_name, key=f"dl_{key}", mime="text/csv",
                use_container_width=True)
        idx += 1
    ctrl_value = None
    if has_ctrl:
        with cols[idx]:
            ctrl_value = controls()
    return ctrl_value


# --------------------------------------------------------------------------- #
# Tabs
# --------------------------------------------------------------------------- #
tab_ov, tab_be, tab_dr, tab_oh, tab_cp, tab_sc, tab_ln = st.tabs(
    ["Overview", "Vs BE", "Drill-down", "Orders in Hand",
     "Period compare", "Scheme analysis", "Line items"])

with tab_ov:
    # ── Order trend by channel ──────────────────────────────────────────────
    with st.container(border=True):
        def _gran_ctrl():
            return st.radio(" ", ["Day", "Week", "Month", "Year"], index=2,
                            horizontal=True, key="gran",
                            label_visibility="collapsed")
        # Build the chart's underlying data for CSV export
        _trend = filtered[filtered["_d"].notna()].copy()
        _trend_csv = (_trend.assign(_dt=_trend["_d"].dt.date)
                      .groupby(["_dt", "_ch"])["_q"].sum().reset_index()
                      .rename(columns={"_dt": "Date", "_ch": "Channel",
                                       "_q": "Ordered MT"})) if len(_trend) \
            else pd.DataFrame(columns=["Date", "Channel", "Ordered MT"])
        gran_label = _chart_header(
            "Order trend by channel",
            "Ordered MT split by channel. Click a legend entry to isolate; double-click to restore.",
            csv_df=_trend_csv, csv_name="order_trend.csv", key="trend",
            controls=_gran_ctrl,
        )
        gran = (gran_label or "Month").lower()
        st.plotly_chart(plots.channel_trend(filtered, gran),
                        use_container_width=True)

    # ── Top 10 ship-to states / Grade mix ───────────────────────────────────
    g1, g2 = st.columns(2)
    with g1:
        with st.container(border=True):
            _states_csv = (filtered.groupby("_st")["_q"].sum().reset_index()
                           .sort_values("_q", ascending=False).head(10)
                           .rename(columns={"_st": "Ship-to state",
                                            "_q": "Ordered MT"}))
            _chart_header("Top 10 ship-to states", "By Ordered MT",
                          csv_df=_states_csv, csv_name="top_states.csv",
                          key="states")
            st.plotly_chart(plots.top_states(filtered), use_container_width=True)
    with g2:
        with st.container(border=True):
            _grade_csv = (filtered.groupby("_gr")["_q"].sum().reset_index()
                          .rename(columns={"_gr": "Grade", "_q": "Ordered MT"})
                          .sort_values("Ordered MT", ascending=False))
            _chart_header("Grade mix", "Share of Ordered MT by grade",
                          csv_df=_grade_csv, csv_name="grade_mix.csv",
                          key="grade")
            st.plotly_chart(plots.grade_mix(filtered), use_container_width=True)

    # ── Top 10 distributors / Dispatch plant mix ────────────────────────────
    g3, g4 = st.columns(2)
    with g3:
        with st.container(border=True):
            _dist_csv = (filtered.groupby("_dn")["_q"].sum().reset_index()
                         .sort_values("_q", ascending=False).head(10)
                         .rename(columns={"_dn": "Distributor",
                                          "_q": "Ordered MT"}))
            _chart_header("Top 10 distributors", "By Ordered MT",
                          csv_df=_dist_csv, csv_name="top_distributors.csv",
                          key="dist")
            st.plotly_chart(plots.top_distributors(filtered),
                            use_container_width=True)
    with g4:
        with st.container(border=True):
            _plant_csv = (filtered.groupby("_cm")["_q"].sum().reset_index()
                          .rename(columns={"_cm": "CM / Plant",
                                           "_q": "Ordered MT"})
                          .sort_values("Ordered MT", ascending=False))
            _chart_header("Dispatch plant mix", "Share of Ordered MT by plant",
                          csv_df=_plant_csv, csv_name="plant_mix.csv",
                          key="plant")
            st.plotly_chart(plots.plant_mix(filtered), use_container_width=True)

    # ── India heat map ──────────────────────────────────────────────────────
    with st.container(border=True):
        def _metric_ctrl():
            return st.radio(" ", ["Ordered MT", "Released MT",
                                  "Invoiced MT", "Line count"],
                            index=0, horizontal=True, key="map_metric",
                            label_visibility="collapsed")
        metric_opt = _chart_header(
            "India heat map — Ship-to state",
            "Color intensity scales with the selected metric.",
            csv_df=None, key="map", controls=_metric_ctrl,
        ) or "Ordered MT"
        metric_col = {"Ordered MT": "_q", "Released MT": "_rq",
                      "Invoiced MT": "_iq"}.get(metric_opt)
        if metric_col:
            sv = filtered.groupby("_st")[metric_col].sum().reset_index()
        else:
            sv = filtered.groupby("_st").size().reset_index(name="value")
            sv = sv.rename(columns={0: "value"})
        sv.columns = ["state", "value"]
        st.plotly_chart(plots.india_map(sv, metric_opt),
                        use_container_width=True)

with tab_be:
    st.subheader("Vs Best Estimate")

    # ── BE version picker (latest of latest by default) ─────────────────────
    def _be_slot(b: be_logic.BeVersion | None) -> tuple[str, str] | None:
        if b is None:
            return None
        return (f"{b.month_y:04d}-{b.month_m + 1:02d}", b.week or "")

    versions = storage.list_be_versions() if storage.is_configured() else []
    if versions:
        v_labels = [f"{v['month']} · {v['week']}" for v in versions]
        sel = st.selectbox(
            "Active BE version", v_labels, index=0, key="be_version_picker",
            help="Switch between saved BE versions for any month/week.")
        sel_v = versions[v_labels.index(sel)]
        if _be_slot(be) != (sel_v["month"], sel_v["week"]):
            with st.spinner("Loading BE…"):
                st.session_state.be_version = restore_be(sel_v["month"], sel_v["week"])
            st.rerun()

    with st.expander("Upload new BE", expanded=be is None):
        be_up = st.file_uploader("BE Excel", type=["xlsx", "xls"], key="be_upload")
        if be_up is not None:
            be_bytes = be_up.getvalue()
            sheets = list_be_sheets(be_bytes)
            prio = [s for s in sheets if any(t in s.lower() for t in ("distributor be", "be week"))]
            ordered = prio + [s for s in sheets if s not in prio]
            sheet = st.selectbox("BE sheet", ordered)
            c_m, c_w = st.columns(2)
            month_val = c_m.text_input("BE for month (YYYY-MM)",
                                       value=datetime.now().strftime("%Y-%m"))
            week_val = c_w.selectbox("Week", ["W1", "W2", "W3", "W4"], index=0,
                                     help="Up to 4 BE versions per month.")
            exists = (storage.is_configured()
                      and storage.be_version_exists(month_val, week_val))
            if exists:
                st.warning(f"⚠️ A BE for **{month_val} · {week_val}** already exists. "
                           "Loading will **overwrite** it.")
            if st.button("Load BE", type="primary"):
                rows = parse_be(be_bytes, sheet)
                if not rows:
                    st.error("Could not extract any BE rows from this sheet.")
                else:
                    label, y, m0 = be_logic.month_label_from_value(month_val)
                    total = sum(r["total"] for r in rows)
                    st.session_state.be_version = be_logic.BeVersion(
                        upload_date=datetime.now().isoformat(), month_label=label,
                        month_y=y, month_m=m0, week=week_val, sheet=sheet,
                        rows=rows, total_be=total)
                    if storage.is_configured():
                        storage.save_be_version(month_val, week_val, be_bytes, {
                            "sheet": sheet, "month": month_val, "week": week_val,
                            "uploaded": datetime.now().isoformat()})
                    st.success(f"BE loaded: {label} · {week_val} · "
                               f"{len({r['distNorm'] for r in rows})} distributors · "
                               f"{data.fmt(total)} MT")
                    st.rerun()

    if ag is None:
        st.info("Load a BE file to enable plan-vs-actual comparison.")
    else:
        b1, b2, b3, b4 = st.columns(4)
        b1.metric("BE total (MT)", data.fmt(ag.tot_be))
        b2.metric("Matched actuals (MT)", data.fmt(ag.matched_act))
        b3.metric("Open pipeline (MT)", data.fmt(ag.matched_pipe))
        b4.metric("Gap to BE (MT)", data.fmt(ag.tot_be - ag.matched_act))

        days_in_month = ((ag.be_month_end - ag.be_month_start).days + 1)
        dist_set = ag.dist_has_be
        daily = data.invoiced_daily_map(
            filtered, ag.be_month_start, ag.be_month_end, inv_index,
            filter_fn=lambda r: be_logic.be_eligible(r, dist_set))
        st.subheader("Daily invoice trajectory — BE month")
        st.plotly_chart(
            plots.be_trajectory(daily, ag.tot_be, be.month_label, days_in_month),
            use_container_width=True)

        st.subheader("Unmatched BE distributors (no actuals)")
        if ag.unmatched_be:
            um = pd.DataFrame([{"Distributor": u["dist"], "Region": u["region"],
                                "BE MT": round(u["qty"])} for u in
                               sorted(ag.unmatched_be, key=lambda x: -x["qty"])])
            st.dataframe(um, use_container_width=True, hide_index=True)
        else:
            st.caption("Every BE distributor has matching activity.")

with tab_dr:
    st.info("Drill-down tab — TODO (iteration). Data layer ready via "
            "`data.aggregate()`; assemble the expandable hierarchy here.")

with tab_oh:
    st.subheader("Orders in Hand")
    oh = filtered[filtered["_pend"] > 0]
    o1, o2, o3 = st.columns(3)
    o1.metric("Pending release (MT)", data.fmt(oh["_pend"].sum()))
    o2.metric("Pending invoice (MT)", data.fmt(filtered["_pendInv"].sum()))
    sc = filtered[filtered["_scShort"]]
    o3.metric("Short-closed lines", f"{len(sc):,}", f"{data.fmt(sc['_pendOrig'].sum())} MT")

with tab_cp:
    st.info("Period compare tab — TODO (iteration). Use `data.apply_filters()` with "
            "two date windows and diff the aggregates.")

with tab_sc:
    st.info("Scheme analysis tab — TODO (iteration).")

with tab_ln:
    st.subheader("Line items")
    cols = ["_d", "_oid", "_dn", "_pt", "_sta", "_st", "_gr", "_dia", "_fm",
            "_q", "_rq", "_iq", "_cm"]
    view = filtered[cols].rename(columns={
        "_d": "Date", "_oid": "Order ID", "_dn": "Distributor", "_pt": "Type",
        "_sta": "Status", "_st": "Ship to", "_gr": "Grade", "_dia": "Dia",
        "_fm": "Form", "_q": "Qty MT", "_rq": "Rel MT", "_iq": "Inv MT", "_cm": "CM"})
    st.download_button("⇩ CSV", view.to_csv(index=False).encode("utf-8"),
                       file_name="line_items.csv", mime="text/csv")
    st.dataframe(view, use_container_width=True, hide_index=True, height=520)


# ─── Universal drill-down drawer (rendered last so it floats above) ──────────
drawer.render()
