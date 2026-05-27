"""JSW Order Intelligence Dashboard — Streamlit entry point.

Sidebar filters, KPI strip, and tabs. The data layer (data.py / be_logic.py) is
complete; some tabs are marked TODO for the iteration session (see README).
"""
from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
import streamlit as st

import be_logic
import data
import plots
import supabase_io

st.set_page_config(page_title="JSW One — Order Intelligence", layout="wide")


# --------------------------------------------------------------------------- #
# Data loading (cached on raw bytes)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner="Parsing Excel...")
def parse_order_workbook(file_bytes: bytes):
    """Read the Order + Invoice sheets and return (enriched_df, inv_index)."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    inv_index = {}
    inv_sheet = next((s for s in xls.sheet_names if s.lower() == "invoice"), None) \
        or next((s for s in xls.sheet_names if "invoice" in s.lower()), None)
    if inv_sheet:
        inv_df = pd.read_excel(xls, sheet_name=inv_sheet)
        inv_index = data.build_invoice_index(inv_df.to_dict("records"))

    order_sheet = next((s for s in xls.sheet_names if "order" in s.lower()), xls.sheet_names[0])
    order_df = pd.read_excel(xls, sheet_name=order_sheet)
    enriched = data.enrich(order_df.to_dict("records"), inv_index)
    return enriched, inv_index


@st.cache_data(show_spinner="Parsing BE sheet...")
def list_be_sheets(file_bytes: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


@st.cache_data(show_spinner=False)
def parse_be(file_bytes: bytes, sheet: str) -> list[dict]:
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None)
    aoa = raw.where(pd.notna(raw), "").values.tolist()
    return be_logic.parse_be_sheet(aoa)


# --------------------------------------------------------------------------- #
# Session state
# --------------------------------------------------------------------------- #
if "be_version" not in st.session_state:
    st.session_state.be_version = None


def get_data():
    """Return (df, inv_index) from an uploaded file or Supabase, else (None, None)."""
    up = st.session_state.get("order_upload")
    if up is not None:
        file_bytes = up.getvalue()
        if supabase_io.is_configured():
            supabase_io.save_order_file(file_bytes)
        return parse_order_workbook(file_bytes)
    cached = supabase_io.load_order_file()
    if cached:
        return parse_order_workbook(cached)
    return None, None


# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #
hdr_l, hdr_r = st.columns([3, 1])
with hdr_l:
    st.title("JSW One — Order Intelligence")
with hdr_r:
    if supabase_io.is_configured():
        st.caption("Supabase: connected")
    else:
        st.caption("Supabase: not configured (no persistence)")

st.file_uploader("Load order Excel (.xlsx)", type=["xlsx", "xls"], key="order_upload")

df, inv_index = get_data()

if df is None or not len(df):
    st.info("Upload your Excel file. The dashboard reads the **Order** sheet "
            "(and an **Invoice** sheet if present) and builds all metrics.")
    st.stop()


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
kpis = data.compute_kpis(filtered)
period = data.get_active_period(f)

# Invoiced-in-period (invoice-date scoped, ignores order-date filter)
nd = data.apply_non_date_filters(df, f)
if period:
    inv_in_period = float(
        nd.apply(lambda r: data.invoiced_in_range(r, period["from"], period["to"], inv_index),
                 axis=1).sum()) if len(nd) else 0.0
else:
    inv_in_period = float(nd["_iq"].sum()) if len(nd) else 0.0

be = st.session_state.be_version
ag = be_logic.be_actuals_agg(filtered, be, inv_index) if be else None

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Ordered (MT)", data.fmt(kpis.ordered), f"{kpis.line_items:,} lines")
k2.metric("Released (MT)", data.fmt(kpis.released),
          f"{round(kpis.released / kpis.ordered * 100) if kpis.ordered else 0}% of ordered")
k3.metric("Invoiced (MT)", data.fmt(kpis.invoiced),
          f"{round(kpis.invoiced / kpis.ordered * 100) if kpis.ordered else 0}% of ordered")
k4.metric("Invoiced in period (MT)", data.fmt(inv_in_period),
          period["label"] if period else "All time")
if ag:
    gap = ag.tot_be - ag.matched_act
    k5.metric(f"BE gap — {be.month_label}", data.fmt(gap),
              f"{data.fmt(ag.matched_act)} / {data.fmt(ag.tot_be)} MT")
else:
    k5.metric("BE gap", "—", "load a BE file")


# --------------------------------------------------------------------------- #
# Tabs
# --------------------------------------------------------------------------- #
tab_ov, tab_be, tab_dr, tab_oh, tab_cp, tab_sc, tab_ln = st.tabs(
    ["Overview", "Vs BE", "Drill-down", "Orders in Hand",
     "Period compare", "Scheme analysis", "Line items"])

with tab_ov:
    c1, c2 = st.columns([3, 1])
    gran = c1.radio("Granularity", ["day", "week", "month", "year"], index=2,
                    horizontal=True, key="gran")
    ctype = c2.radio("Chart", ["line", "bar"], horizontal=True, key="ctype")
    st.subheader("Order trend by channel")
    st.plotly_chart(plots.channel_trend(filtered, gran, ctype), use_container_width=True)

    g1, g2 = st.columns(2)
    with g1:
        st.subheader("Top 10 ship-to states")
        st.plotly_chart(plots.top_states(filtered), use_container_width=True)
    with g2:
        st.subheader("Grade mix")
        st.plotly_chart(plots.grade_mix(filtered), use_container_width=True)

    g3, g4 = st.columns(2)
    with g3:
        st.subheader("Top 10 distributors")
        st.plotly_chart(plots.top_distributors(filtered), use_container_width=True)
    with g4:
        st.subheader("Dispatch plant mix")
        st.plotly_chart(plots.plant_mix(filtered), use_container_width=True)

    st.subheader("India heat map — Ship-to state")
    metric_opt = st.selectbox("Metric", ["Ordered MT", "Released MT", "Invoiced MT", "Line count"])
    metric_col = {"Ordered MT": "_q", "Released MT": "_rq", "Invoiced MT": "_iq"}.get(metric_opt)
    if metric_col:
        sv = filtered.groupby("_st")[metric_col].sum().reset_index()
    else:
        sv = filtered.groupby("_st").size().reset_index(name="value")
        sv = sv.rename(columns={0: "value"})
    sv.columns = ["state", "value"]
    st.plotly_chart(plots.india_map(sv, metric_opt), use_container_width=True)

with tab_be:
    st.subheader("Vs Best Estimate")
    with st.expander("Load / replace BE file", expanded=be is None):
        be_up = st.file_uploader("BE Excel", type=["xlsx", "xls"], key="be_upload")
        if be_up is not None:
            be_bytes = be_up.getvalue()
            sheets = list_be_sheets(be_bytes)
            prio = [s for s in sheets if any(t in s.lower() for t in ("distributor be", "be week"))]
            ordered = prio + [s for s in sheets if s not in prio]
            sheet = st.selectbox("BE sheet", ordered)
            month_val = st.text_input("BE for month (YYYY-MM)",
                                      value=datetime.now().strftime("%Y-%m"))
            week = st.text_input("Week label (optional)", value="")
            if st.button("Load BE"):
                rows = parse_be(be_bytes, sheet)
                if not rows:
                    st.error("Could not extract any BE rows from this sheet.")
                else:
                    label, y, m0 = be_logic.month_label_from_value(month_val)
                    total = sum(r["total"] for r in rows)
                    st.session_state.be_version = be_logic.BeVersion(
                        upload_date=datetime.now().isoformat(), month_label=label,
                        month_y=y, month_m=m0, week=week.strip(), sheet=sheet,
                        rows=rows, total_be=total)
                    if supabase_io.is_configured():
                        supabase_io.save_be_file(be_bytes)
                    st.success(f"BE loaded: {label} · {len({r['distNorm'] for r in rows})} "
                               f"distributors · {data.fmt(total)} MT")
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
