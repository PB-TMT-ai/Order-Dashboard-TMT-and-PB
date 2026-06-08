"""BE Output Report — the monthly AOP-vs-BE / order-book / invoicing operating
report (new tab).

Single-row report (Business/Category dimension excluded) for the month of the
loaded BE version. Scope = like the BE tab: TMT + Fe 550/550D + channels
{Retail, Self-stocking, Project-thru-Dist}. Order book uses the un-invoiced
backlog definition; invoicing is invoice-date attributed via the invoice index.

AOP figures come from two uploaded workbooks (Board / Internal), summing the
TMT business lines (JSW ONE TMT + One Helix TMT). Order BE and the Realistic-BE
override are manual inputs held in session state (reset on reload).
"""
from __future__ import annotations

import calendar
import io
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

import be_logic
from data import cl, invoiced_in_range

# AOP rows that make up the TMT report scope ("all TMT lines").
_AOP_TMT_ROWS = ("jsw one tmt", "one helix tmt")
_MONTH_FMT = "%b-%y"  # e.g. "Apr-26"


# ── month / fiscal helpers ───────────────────────────────────────────────────
def _month_start(y: int, m: int) -> datetime:
    return datetime(y, m, 1)


def _month_end(y: int, m: int) -> datetime:
    last = calendar.monthrange(y, m)[1]
    return datetime(y, m, last, 23, 59, 59)


def _quarter_months(y: int, m: int) -> list[tuple[int, int]]:
    """The three calendar (year, month) of the fiscal quarter containing (y, m).
    Fiscal year starts in April."""
    fy_start = y if m >= 4 else y - 1
    fm = (m - 4) % 12                 # Apr->0 … Mar->11
    qf = (fm // 3) * 3                # first fiscal-month of the quarter
    out = []
    for f in (qf, qf + 1, qf + 2):
        cm = (f + 3) % 12 + 1
        cy = fy_start if f <= 8 else fy_start + 1
        out.append((cy, cm))
    return out


# ── AOP parsing ──────────────────────────────────────────────────────────────
def parse_aop(file_bytes: bytes) -> tuple[dict[tuple[int, int], float], str]:
    """Parse an AOP workbook → {(year, month): MT} summed over the TMT rows.

    Locates the 'Volume (mt)' block, reads the month header row (Apr-26 …),
    and sums the JSW ONE TMT + One Helix TMT rows across the month columns.
    Returns (mapping, error). On failure mapping is empty and error is set.
    """
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet = next((s for s in xls.sheet_names if "aop" in s.lower()),
                     xls.sheet_names[0])
        raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
    except Exception as e:  # noqa: BLE001
        return {}, f"Could not read AOP workbook: {e}"

    nrows, ncols = raw.shape
    # Find the 'Volume (mt)' marker row.
    vol_row = None
    for i in range(nrows):
        for j in range(min(ncols, 4)):
            if "volume" in str(raw.iat[i, j] or "").lower():
                vol_row = i
                break
        if vol_row is not None:
            break
    if vol_row is None:
        return {}, "Could not find a 'Volume (mt)' section in the AOP sheet."

    # Month header columns: scan a few rows around the Volume marker for tokens
    # like 'Apr-26'. Use the row with the most parseable month cells.
    best_cols: dict[int, tuple[int, int]] = {}
    for i in range(max(0, vol_row - 2), min(nrows, vol_row + 2)):
        cand: dict[int, tuple[int, int]] = {}
        for j in range(ncols):
            dt = _parse_month_cell(raw.iat[i, j])
            if dt is not None:
                cand[j] = dt
        if len(cand) > len(best_cols):
            best_cols = cand
    if not best_cols:
        return {}, "Could not find month columns (e.g. 'Apr-26') in the AOP sheet."

    # Sum the TMT rows below the volume marker.
    out: dict[tuple[int, int], float] = {}
    for i in range(vol_row, min(nrows, vol_row + 12)):
        label = str(raw.iat[i, 0] or "").strip().lower()
        if label in _AOP_TMT_ROWS:
            for j, (yy, mm) in best_cols.items():
                val = _num(raw.iat[i, j])
                out[(yy, mm)] = out.get((yy, mm), 0.0) + val
    if not out:
        return {}, ("Found the Volume block but no TMT rows "
                    "(expected 'JSW ONE TMT' / 'One Helix TMT').")
    return out, ""


def _parse_month_cell(v: object) -> tuple[int, int] | None:
    s = str(v or "").strip()
    if not s:
        return None
    for fmt in (_MONTH_FMT, "%b-%Y", "%B-%y", "%B-%Y"):
        try:
            d = datetime.strptime(s, fmt)
            return d.year, d.month
        except ValueError:
            continue
    return None


def _num(v: object) -> float:
    if v is None:
        return 0.0
    try:
        return float(str(v).replace(",", "").strip() or 0)
    except (ValueError, TypeError):
        return 0.0


# ── order-book / invoicing primitives ────────────────────────────────────────
def _scope(df: pd.DataFrame) -> pd.DataFrame:
    if not len(df):
        return df
    return df[
        (df["_pt"] == "TMT")
        & df["_gr"].astype(str).str.lower().str.replace(" ", "", regex=False)
            .isin(("fe550", "fe550d"))
        & df["_ch"].isin(["rt", "ss", "pd"])
    ].copy()


def _inv_before(row, S: datetime, inv_index) -> float:
    """Invoiced qty for an order strictly before date S (invoice-date based)."""
    e = inv_index.get(row["_oid"])
    iq = row["_iq"]
    if e is None or e.total_qty == 0:
        d = row["_d"]
        return float(iq) if (iq > 0 and pd.notna(d) and d < S) else 0.0
    q = sum(inv["qty"] for inv in e.invs
            if inv["date"] is not None and inv["date"] < S)
    return q * (iq / e.total_qty)


def opening_book(scope: pd.DataFrame, S: datetime, inv_index) -> float:
    """Un-invoiced backlog at the start of month (date S): orders booked before S
    minus what was invoiced before S."""
    tot = 0.0
    for _, r in scope.iterrows():
        if pd.notna(r["_d"]) and r["_d"] < S:
            tot += max(float(r["_q"]) - _inv_before(r, S, inv_index), 0.0)
    return tot


def orders_booked(scope: pd.DataFrame, S: datetime, E: datetime) -> float:
    """Ordered MT booked within [S, E] (by order date)."""
    d = scope["_d"]
    win = scope[d.notna() & (d >= S) & (d <= E)]
    return float(win["_q"].sum()) if len(win) else 0.0


def invoiced_split(scope: pd.DataFrame, S: datetime, E: datetime,
                   inv_index) -> tuple[float, float]:
    """Invoiced in [S, E] split into (from_opening, from_current): orders booked
    before S vs booked within the month."""
    from_open = from_cur = 0.0
    for _, r in scope.iterrows():
        q = invoiced_in_range(r, S, E, inv_index)
        if q <= 0:
            continue
        if pd.notna(r["_d"]) and r["_d"] < S:
            from_open += q
        else:
            from_cur += q
    return from_open, from_cur


def total_invoiced(scope: pd.DataFrame, S: datetime, E: datetime,
                   inv_index) -> float:
    return float(sum(invoiced_in_range(r, S, E, inv_index)
                     for _, r in scope.iterrows()))


# ── report builder ───────────────────────────────────────────────────────────
def build_report(filtered: pd.DataFrame, be, ag, inv_index,
                 now: datetime, aop_board: dict, aop_internal: dict,
                 order_be: float) -> tuple[dict, dict]:
    """Return (columns, meta). columns is an ordered {label: value} dict for the
    one-row report; meta carries day counts / labels for captions."""
    sy, sm = be.month_y, be.month_m + 1            # selected month (1-indexed)
    m_start = _month_start(sy, sm)
    m_end = _month_end(sy, sm)
    asof = min(now, m_end)                          # MTD cut-off
    scope = _scope(filtered)

    qmonths = _quarter_months(sy, sm)               # 3 (y, m) of the quarter
    qlabels = [datetime(y, m, 1).strftime("%b") for (y, m) in qmonths]

    cols: dict[str, float] = {}

    # AOP Vs BE
    q_board = sum(aop_board.get((y, m), 0.0) for (y, m) in qmonths)
    q_internal = sum(aop_internal.get((y, m), 0.0) for (y, m) in qmonths)
    cols["Q AOP Board"] = q_board
    cols["Q AOP Internal"] = q_internal
    # QTD actual (completed quarter months) + selected-month BE
    qtd_actual = 0.0
    for (y, m) in qmonths:
        if (y, m) == (sy, sm):
            continue
        if _month_start(y, m) < m_start:            # completed month
            qtd_actual += total_invoiced(scope, _month_start(y, m),
                                         _month_end(y, m), inv_index)
    cols["QTD + BE (Actual+BE)"] = qtd_actual + float(ag.tot_be)
    cols["AOP Board (month)"] = aop_board.get((sy, sm), 0.0)
    cols["AOP Internal (month)"] = aop_internal.get((sy, sm), 0.0)
    cols["BE (month, W1)"] = float(ag.tot_be)

    # Order book — Opening + Current orders per quarter month
    opening_sel = 0.0
    for (y, m), lbl in zip(qmonths, qlabels):
        s = _month_start(y, m)
        ob = opening_book(scope, s, inv_index)
        cols[f"Opening OB {lbl}"] = ob
        if (y, m) == (sy, sm):
            opening_sel = ob
    for (y, m), lbl in zip(qmonths, qlabels):
        s, e = _month_start(y, m), _month_end(y, m)
        e_eff = asof if (y, m) == (sy, sm) else e
        cols[f"Orders {lbl}"] = orders_booked(scope, s, e_eff)
    cols["Order BE (month)"] = float(order_be)

    # Invoicing
    for (y, m), lbl in zip(qmonths, qlabels):
        if (y, m) == (sy, sm) or _month_start(y, m) >= m_start:
            continue
        cols[f"Invoiced {lbl} (full)"] = total_invoiced(
            scope, _month_start(y, m), _month_end(y, m), inv_index)
    from_open, from_cur = invoiced_split(scope, m_start, asof, inv_index)
    cols["Inv from Opening"] = from_open
    cols["Inv from Current"] = from_cur
    inv_mtd = from_open + from_cur
    cols["Inv MTD"] = inv_mtd

    # Rates & closing
    days_in_month = calendar.monthrange(sy, sm)[1]
    days_elapsed = max(1, min(days_in_month, (asof.date() - m_start.date()).days + 1))
    rem_days = max(0, days_in_month - days_elapsed)
    orders_sel = cols.get(f"Orders {datetime(sy, sm, 1).strftime('%b')}", 0.0)
    cols["Invoice DRR req (MT/day)"] = (
        (float(ag.tot_be) - inv_mtd) / rem_days if rem_days > 0 else 0.0)
    cols["Order DRR req (MT/day)"] = (
        (float(order_be) - orders_sel) / rem_days if rem_days > 0 else 0.0)
    cols["Closing Order Book"] = opening_sel + float(order_be) - float(ag.tot_be)

    # Standalone — last-year same month invoiced
    cols["LY same month (invoiced)"] = total_invoiced(
        scope, _month_start(sy - 1, sm), _month_end(sy - 1, sm), inv_index)

    meta = {"month_label": be.month_label, "qlabels": qlabels,
            "days_in_month": days_in_month, "days_elapsed": days_elapsed,
            "rem_days": rem_days, "scope_rows": len(scope)}
    return cols, meta


# ── per-distributor report ───────────────────────────────────────────────────
def build_dist_report(filtered: pd.DataFrame, be, ag, inv_index,
                      now: datetime, order_be_map: dict) -> tuple[pd.DataFrame, dict]:
    """Per-distributor BE Output table (same grain as the Vs BE tab).

    One row per distributor that either has BE or has in-scope activity. Every
    order-book / invoicing figure is computed on that distributor's own scope
    rows; Order BE is a per-distributor manual override held in session state.
    Columns mirror the total report: BE, Opening OB, Orders (MTD), invoicing
    split, Order BE, Closing OB, the two DRRs and last-year same month.
    """
    sy, sm = be.month_y, be.month_m + 1            # selected month (1-indexed)
    m_start = _month_start(sy, sm)
    m_end = _month_end(sy, sm)
    asof = min(now, m_end)                          # MTD cut-off
    scope = _scope(filtered)
    ly_start, ly_end = _month_start(sy - 1, sm), _month_end(sy - 1, sm)

    days_in_month = calendar.monthrange(sy, sm)[1]
    days_elapsed = max(1, min(days_in_month, (asof.date() - m_start.date()).days + 1))
    rem_days = max(0, days_in_month - days_elapsed)

    groups = dict(tuple(scope.groupby("_dnN"))) if len(scope) else {}
    dist_ns = set(ag.atom_lookup) | set(groups)
    empty = scope.iloc[0:0]

    rows: list[dict] = []
    for distN in dist_ns:
        atom = ag.atom_lookup.get(distN)
        g = groups.get(distN, empty)
        be_qty = float(atom["qty"]) if atom else 0.0
        if atom:
            name, state = atom["dist"], (atom.get("state") or "—")
        elif len(g):
            name = cl(g["_dn"].iloc[0]) or "Direct"
            state = cl(g["_st"].iloc[0]).title() or "—"
        else:
            name, state = distN, "—"

        opening = opening_book(g, m_start, inv_index) if len(g) else 0.0
        orders = orders_booked(g, m_start, asof) if len(g) else 0.0
        from_open, from_cur = (invoiced_split(g, m_start, asof, inv_index)
                               if len(g) else (0.0, 0.0))
        inv_mtd = from_open + from_cur
        ly = total_invoiced(g, ly_start, ly_end, inv_index) if len(g) else 0.0
        order_be = float(order_be_map.get(distN, 0.0))

        rows.append({
            "Distributor": name,
            "State": state,
            "BE": be_qty,
            "Opening OB": opening,
            "Orders (MTD)": orders,
            "Inv from Opening": from_open,
            "Inv from Current": from_cur,
            "Inv MTD": inv_mtd,
            "Order BE (override)": order_be,
            "Closing OB": opening + order_be - be_qty,
            "Invoice DRR req": (be_qty - inv_mtd) / rem_days if rem_days > 0 else 0.0,
            "Order DRR req": (order_be - orders) / rem_days if rem_days > 0 else 0.0,
            "LY same month": ly,
            "_distNorm": distN,
            "_hasBE": atom is not None,
        })

    cols = ["Distributor", "State", "BE", "Opening OB", "Orders (MTD)",
            "Inv from Opening", "Inv from Current", "Inv MTD",
            "Order BE (override)", "Closing OB", "Invoice DRR req",
            "Order DRR req", "LY same month", "_distNorm", "_hasBE"]
    out = pd.DataFrame(rows, columns=cols)
    if len(out):
        out = out.sort_values("BE", ascending=False).reset_index(drop=True)
    meta = {"month_label": be.month_label, "days_in_month": days_in_month,
            "days_elapsed": days_elapsed, "rem_days": rem_days,
            "scope_rows": len(scope)}
    return out, meta


# ── render ───────────────────────────────────────────────────────────────────
def render(filtered, df_cancelled, be, ag, inv_index, now,
           mt_col, pct_col, chart_header, kpi_card,
           kpi_view=None, open_drawer=None) -> None:
    st.markdown(
        '<div class="chart-title">BE Output Report</div>'
        '<div class="chart-sub">Monthly AOP-vs-BE, order book and invoicing for '
        'the loaded BE month. Scope: TMT + Fe 550/550D + Retail / Self-stocking '
        '/ Project-thru-Dist. Cancelled orders excluded.</div>',
        unsafe_allow_html=True)

    if ag is None or be is None:
        st.info("Load a BE file on the **Vs BE** tab — the report month and BE "
                "figures come from it.")
        return

    # Inputs: AOP uploads + manual Order BE.
    with st.expander("Inputs — AOP files & Order BE", expanded=False):
        c1, c2, c3 = st.columns(3)
        board_up = c1.file_uploader("Board AOP (.xlsx)", type=["xlsx"],
                                    key="aop_board_up")
        internal_up = c2.file_uploader("Internal AOP (.xlsx)", type=["xlsx"],
                                       key="aop_internal_up")
        order_be = c3.number_input("Order BE (month, MT)", min_value=0.0,
                                   value=float(st.session_state.get("_bo_order_be", 0.0)),
                                   step=100.0, key="bo_order_be_input")
        st.session_state["_bo_order_be"] = order_be

    aop_board = _load_aop(board_up, "_aop_board_cache")
    aop_internal = _load_aop(internal_up, "_aop_internal_cache")
    if board_up is None or internal_up is None:
        st.caption("⬆ Upload Board & Internal AOP files to populate the AOP "
                   "columns (the rest of the report works without them).")

    cols, meta = build_report(filtered, be, ag, inv_index, now,
                              aop_board, aop_internal, order_be)

    with st.container(border=True):
        report_df = pd.DataFrame([cols])
        chart_header(
            f"Report — {meta['month_label']}",
            f"MTD as of {now:%d %b %Y} · {meta['days_elapsed']}/"
            f"{meta['days_in_month']} days elapsed · {meta['rem_days']} remaining "
            f"· {meta['scope_rows']:,} in-scope order lines.",
            csv_df=report_df, csv_name="be_output_report.csv", key="bo_report")
        cfg = {}
        for c in report_df.columns:
            cfg[c] = (pct_col(c) if "DRR" in c else mt_col(c))
        st.dataframe(report_df, use_container_width=True, hide_index=True,
                     column_config=cfg)

    # ── Per-distributor BE Output (same grain as the Vs BE tab) ───────────────
    _render_dist_table(filtered, be, ag, inv_index, now, meta,
                       mt_col, chart_header, kpi_view, open_drawer)

    # Item-5 trio at report (total) level.
    with st.container(border=True):
        mom = be_logic.be_mom(filtered, ag, inv_index)
        realistic_auto = float(mom["Realistic BE (auto)"].sum()) if len(mom) else 0.0
        ov = st.session_state.get("_bo_realistic_override")
        c1, c2 = st.columns([1, 2])
        ov_in = c1.number_input(
            "Realistic BE override (MT, blank=auto)", min_value=0.0,
            value=float(ov) if ov is not None else 0.0, step=100.0,
            key="bo_realistic_override_input",
            help="0 keeps the auto 3-month-average Realistic BE.")
        st.session_state["_bo_realistic_override"] = ov_in if ov_in > 0 else None
        realistic_used = ov_in if ov_in > 0 else realistic_auto
        var = float(ag.tot_be) - realistic_used
        chart_header("Realistic BE & Volume at Risk (total)",
                     "Realistic BE auto = sum of per-distributor 3-month-average "
                     "actuals. Volume at Risk = BE − Realistic BE (used).",
                     key="bo_trio")
        cards = [
            kpi_card("k-or", "BE (total)", f"{ag.tot_be:,.0f}", meta["month_label"]),
            kpi_card("k-inp", "Realistic BE (auto)", f"{realistic_auto:,.0f}",
                     "3-month average"),
            kpi_card("k-in", "Realistic BE (used)", f"{realistic_used:,.0f}",
                     "override" if ov_in > 0 else "auto"),
            kpi_card("k-gap", "Volume at Risk", f"{var:,.0f}",
                     "BE − Realistic", value_cls="dn" if var > 0 else "up"),
        ]
        st.markdown(
            '<div class="kpi-row" style="grid-template-columns:repeat(4,1fr);">'
            + "".join(cards) + "</div>", unsafe_allow_html=True)


def _render_dist_table(filtered, be, ag, inv_index, now, meta,
                       mt_col, chart_header, kpi_view, open_drawer) -> None:
    """Per-distributor table with an editable Order BE column + drill-down,
    mirroring the Vs BE tab's BE-vs-Actuals editor."""
    with st.container(border=True):
        ov_map = st.session_state.setdefault("_bo_order_be_map", {})
        tbl, _ = build_dist_report(filtered, be, ag, inv_index, now, ov_map)
        chart_header(
            "BE Output — per distributor",
            "Opening order book, MTD orders & invoicing split, Closing OB and the "
            "Invoice/Order DRR required to hit BE — per distributor. Edit Order BE "
            "(override) to drive Closing OB and Order DRR; pick a distributor to drill.",
            csv_df=tbl.drop(columns=["_distNorm", "_hasBE"]),
            csv_name="be_output_by_distributor.csv", key="bo_dist_table")

        if not len(tbl):
            st.info("No in-scope distributors for this BE month yet.")
            return

        fc1, fc2 = st.columns([2, 3])
        search = fc1.text_input("Search distributor", "",
                                key="bo_tbl_search").strip().lower()
        states = sorted(s for s in tbl["State"].unique() if s and s != "—")
        sel_states = fc2.multiselect("State filter", states, default=[],
                                     key="bo_tbl_states")
        view = tbl.copy()
        if search:
            view = view[view["Distributor"].astype(str).str.lower().str.contains(search)]
        if sel_states:
            view = view[view["State"].isin(sel_states)]
        view = view.reset_index(drop=True)

        editor_cols = ["Distributor", "State", "BE", "Opening OB", "Orders (MTD)",
                       "Inv from Opening", "Inv from Current", "Inv MTD",
                       "Order BE (override)", "Closing OB", "Invoice DRR req",
                       "Order DRR req", "LY same month"]
        colcfg = {c: mt_col(c) for c in editor_cols if c not in ("Distributor", "State")}
        colcfg["Order BE (override)"] = st.column_config.NumberColumn(
            "Order BE (override)", format="%.0f",
            help="Per-distributor Order BE for the month. Drives Closing OB "
                 "(Opening + Order BE − BE) and Order DRR. Resets on reload.")
        disabled = [c for c in editor_cols if c != "Order BE (override)"]
        edited = st.data_editor(
            view[editor_cols], use_container_width=True, hide_index=True,
            height=460, column_config=colcfg, disabled=disabled,
            key="bo_dist_editor")

        # Persist edited Order BE back into the session map; rerun once on change
        # so Closing OB / Order DRR refresh.
        ov_series = edited["Order BE (override)"]
        changed = False
        for pos, distN in enumerate(view["_distNorm"].values):
            val = ov_series.iloc[pos]
            new = float(val) if pd.notna(val) else 0.0
            if new > 0:
                if ov_map.get(distN) != new:
                    ov_map[distN] = new
                    changed = True
            elif distN in ov_map:
                del ov_map[distN]
                changed = True
        if changed:
            st.rerun()

        # Drill-down to the shared drawer (matches the Vs BE tab).
        if kpi_view is None or open_drawer is None:
            return
        picks = ["—"] + view["Distributor"].astype(str).tolist()
        drill_pick = st.selectbox("Drill into a distributor's orders", picks,
                                  key="bo_drill_pick")
        if drill_pick and drill_pick != "—":
            row = view[view["Distributor"].astype(str) == drill_pick].iloc[0]
            dist_rows = filtered[filtered["_dnN"] == row["_distNorm"]]
            if drill_pick != st.session_state.get("_bo_tbl_last"):
                st.session_state["_bo_tbl_last"] = drill_pick
                open_drawer(
                    f"{row['Distributor']} — orders", kpi_view(dist_rows),
                    subtitle=f"{row['State']} · BE {row['BE']:,.0f} · "
                             f"Inv MTD {row['Inv MTD']:,.0f} · "
                             f"Closing OB {row['Closing OB']:,.0f}",
                    summary=[("BE", f"{row['BE']:,.0f}"),
                             ("Opening OB", f"{row['Opening OB']:,.0f}"),
                             ("Orders MTD", f"{row['Orders (MTD)']:,.0f}"),
                             ("Inv MTD", f"{row['Inv MTD']:,.0f}"),
                             ("Order BE", f"{row['Order BE (override)']:,.0f}"),
                             ("Closing OB", f"{row['Closing OB']:,.0f}")],
                    filename=f"{str(row['Distributor'])[:30].replace(' ', '_')}_orders.csv")


def _load_aop(upload, cache_key: str) -> dict:
    """Parse an AOP upload once and cache by content; return the {(y,m): MT} map."""
    if upload is None:
        return {}
    raw = upload.getvalue()
    cached = st.session_state.get(cache_key)
    if cached and cached.get("size") == len(raw):
        return cached["map"]
    mapping, err = parse_aop(raw)
    if err:
        st.warning(f"AOP parse: {err}")
    st.session_state[cache_key] = {"size": len(raw), "map": mapping}
    return mapping
