"""Best Estimate (BE) Excel parsing and plan-vs-actual aggregation.

BE matching is distributor-only: all BE rows for a distributor (across states,
grades, categories) are clubbed into a single number, and any eligible order
from that distributor counts as its actuals wherever it ships.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

import pandas as pd

from data import (InvoiceEntry, MOS, SHORT_CLOSE_DAYS, cl, invoiced_in_range,
                  norm_name, num)


@dataclass
class BeVersion:
    upload_date: str
    month_label: str
    month_y: int
    month_m: int  # 0-indexed month
    week: str
    sheet: str
    rows: list[dict]
    total_be: float


def parse_be_sheet(df_aoa: list[list[Any]]) -> list[dict]:
    """Parse a BE sheet given as an array-of-arrays (header=None read).

    Locates the 'Distributor Name' header row, then the Retail-PTR and
    Distributor-Project group columns for Fe 550 / Fe 550D.
    """
    header_row = dist_col = reg_col = state_col = -1
    for i in range(min(len(df_aoa), 15)):
        row = df_aoa[i]
        if not row:
            continue
        for j, cell in enumerate(row):
            v = str(cell or "").strip().lower()
            if v in ("distributor name", "distributor"):
                header_row, dist_col = i, j
                break
        if header_row >= 0:
            hdr = df_aoa[header_row]
            for j, cell in enumerate(hdr):
                v = str(cell or "").strip().lower()
                if v == "region":
                    reg_col = j
                if v == "state":
                    state_col = j
            break

    if header_row < 0 or dist_col < 0:
        raise ValueError('Could not find "Distributor Name" header row. Check sheet structure.')
    if state_col < 0:
        raise ValueError('Could not find "State" column in BE sheet.')

    group_row = df_aoa[header_row - 1] if header_row > 0 else []
    headers = df_aoa[header_row]
    r_fe550 = r_fe550d = p_fe550 = p_fe550d = -1
    last_group = ""
    for j in range(len(headers)):
        grp = str(group_row[j] if j < len(group_row) else "" or "").strip()
        if grp:
            last_group = grp
        cur = str(headers[j] or "").strip().lower()
        lgl = last_group.lower()
        is_retail = ("retail" in lgl and "ptr" in lgl) and "total" not in lgl
        is_proj = "distributor" in lgl and "project" in lgl and "total" not in lgl
        if is_retail:
            if cur in ("fe 550", "fe550") and r_fe550 < 0:
                r_fe550 = j
            elif cur in ("fe 550d", "fe550d") and r_fe550d < 0:
                r_fe550d = j
        if is_proj:
            if cur in ("fe 550", "fe550") and p_fe550 < 0:
                p_fe550 = j
            elif cur in ("fe 550d", "fe550d") and p_fe550d < 0:
                p_fe550d = j

    def cell(row: list, idx: int) -> Any:
        return row[idx] if 0 <= idx < len(row) else ""

    out: list[dict] = []
    for i in range(header_row + 1, len(df_aoa)):
        row = df_aoa[i]
        if not row:
            continue
        dn = str(cell(row, dist_col) or "").strip()
        if not dn:
            continue
        dnl = dn.lower()
        if dnl.startswith("total ") or dnl == "grand total" or dnl.startswith("grand total"):
            continue
        region = str(cell(row, reg_col) or "").strip() if reg_col >= 0 else ""
        if region.lower() == "total" or not region:
            continue
        state = str(cell(row, state_col) or "").strip()
        if not state:
            continue
        r_f = num(cell(row, r_fe550)) if r_fe550 >= 0 else 0
        r_d = num(cell(row, r_fe550d)) if r_fe550d >= 0 else 0
        p_f = num(cell(row, p_fe550)) if p_fe550 >= 0 else 0
        p_d = num(cell(row, p_fe550d)) if p_fe550d >= 0 else 0
        if r_f + r_d + p_f + p_d <= 0:
            continue
        out.append({
            "distributor": dn, "distNorm": norm_name(dn), "region": region,
            "state": state, "stateU": state.upper(),
            "retail_fe550": r_f, "retail_fe550d": r_d,
            "project_fe550": p_f, "project_fe550d": p_d,
            "retail_total": r_f + r_d, "project_total": p_f + p_d,
            "total": r_f + r_d + p_f + p_d,
        })
    return out


def flatten_be_atomic(be_rows: list[dict]) -> list[dict]:
    """Club BE rows: ONE entry per distributor (summed across states/grades/cats).

    Preserves the distributor's primary state (first non-empty seen) so the
    BE-vs-Actuals table and India gap map can render state-level views.
    """
    grp: dict[str, dict] = {}
    for b in be_rows:
        row = grp.get(b["distNorm"])
        if row is None:
            row = {"dist": b["distributor"], "distNorm": b["distNorm"],
                   "region": b["region"], "state": b.get("state", ""),
                   "qty": 0.0, "retail_be": 0.0, "project_be": 0.0}
            grp[b["distNorm"]] = row
        if not row["state"] and b.get("state"):
            row["state"] = b["state"]
        retail = b.get("retail_fe550", 0) + b.get("retail_fe550d", 0)
        project = b.get("project_fe550", 0) + b.get("project_fe550d", 0)
        row["retail_be"] += retail
        row["project_be"] += project
        row["qty"] += retail + project
    return [a for a in grp.values() if a["qty"] > 0]


def order_cat(row: pd.Series | dict) -> str | None:
    ot = row["_ot"]
    if ot in ("Retailer", "Self-stocking"):
        return "Retail+PTR"
    if ot == "Project":
        return "Project-thru-Dist"
    return None


def be_bucket(ch: str) -> str:
    """Map an order channel to a BE bucket: Retail (rt+ss) vs Project (pd).

    Mirrors the BE sheet's 'Retail-PTR' vs 'Distributor-Project' column groups.
    """
    return "Retail" if ch in ("rt", "ss") else "Project"


def order_grade(row: pd.Series | dict) -> str | None:
    gr = str(row["_gr"] or "").lower().replace(" ", "")
    if gr == "fe550":
        return "Fe 550"
    if gr == "fe550d":
        return "Fe 550D"
    return None


@dataclass
class BeAggregate:
    atomic: list[dict]
    atom_lookup: dict[str, dict]
    orders_by_key: dict[str, list[dict]]
    tot_be: float
    matched_act: float
    matched_pipe: float
    unmatched_be: list[dict]
    dist_has_be: set[str]
    be_month_start: datetime
    be_month_end: datetime
    # Eligible actuals from distributors with NO BE (Target 0). Tracked so the
    # BE-tab actuals reconcile with the invoiced figure for the same scope.
    nobe_orders_by_key: dict[str, list[dict]] = field(default_factory=dict)
    nobe_act: float = 0.0
    nobe_pipe: float = 0.0

    @property
    def total_act(self) -> float:
        """All eligible invoiced actuals in the BE month (BE + no-BE distributors)."""
        return self.matched_act + self.nobe_act

    @property
    def total_pipe(self) -> float:
        return self.matched_pipe + self.nobe_pipe


def be_actuals_agg(df: pd.DataFrame, be: BeVersion,
                   inv_index: dict[str, InvoiceEntry],
                   today: datetime | None = None) -> BeAggregate | None:
    """Aggregate actuals (invoiced + pipeline) against a BE version.

    Eligibility: TMT product, grade Fe 550/550D, channel in {Retail, Self-stocking,
    Project-thru-Dist}, and distributor present in BE. State is ignored (clubbed BE).
    """
    if be is None:
        return None
    if today is None:
        today = datetime.now()

    atomic = flatten_be_atomic(be.rows)
    atom_lookup = {a["distNorm"]: a for a in atomic}
    dist_has_be = set(atom_lookup.keys())

    be_month_start = datetime(be.month_y, be.month_m + 1, 1)
    next_month = (datetime(be.month_y, be.month_m + 2, 1)
                  if be.month_m < 11 else datetime(be.month_y + 1, 1, 1))
    be_month_end = next_month - timedelta(seconds=1)

    orders_by_key: dict[str, list[dict]] = {}
    nobe_orders_by_key: dict[str, list[dict]] = {}
    matched_act = matched_pipe = 0.0
    nobe_act = nobe_pipe = 0.0
    for _, r in df.iterrows():
        if r["_pt"] != "TMT":
            continue
        if order_grade(r) is None:
            continue
        if r["_ch"] not in ("rt", "ss", "pd"):
            continue
        # Eligible by scope. Distributors in BE feed the plan-vs-actual gap;
        # eligible distributors with NO BE are tracked separately so the total
        # actuals still reconcile with the invoiced figure for this scope.
        has_be = r["_dnN"] in dist_has_be
        inv_qty = invoiced_in_range(r, be_month_start, be_month_end, inv_index)
        pipe_eff = 0.0
        if r["_sta"] != "Cancelled":
            pipe = max(r["_q"] - r["_iq"], 0.0)
            if pipe > 0:
                if pipe < 5:
                    pipe_eff = 0.0
                elif pd.notna(r["_d"]):
                    days_old = (today - r["_d"]).days
                    pipe_eff = 0.0 if days_old > SHORT_CLOSE_DAYS else pipe
                else:
                    pipe_eff = pipe
        if inv_qty <= 0 and pipe_eff <= 0:
            continue
        entry = {"r": r, "invQty": inv_qty, "pipe": pipe_eff}
        if has_be:
            matched_act += inv_qty
            matched_pipe += pipe_eff
            orders_by_key.setdefault(r["_dnN"], []).append(entry)
        else:
            nobe_act += inv_qty
            nobe_pipe += pipe_eff
            nobe_orders_by_key.setdefault(r["_dnN"], []).append(entry)

    active = set(orders_by_key.keys())
    unmatched_be = [a for a in atomic if a["distNorm"] not in active]
    tot_be = sum(a["qty"] for a in atomic)

    return BeAggregate(
        atomic=atomic, atom_lookup=atom_lookup, orders_by_key=orders_by_key,
        tot_be=tot_be, matched_act=matched_act, matched_pipe=matched_pipe,
        unmatched_be=unmatched_be, dist_has_be=dist_has_be,
        be_month_start=be_month_start, be_month_end=be_month_end,
        nobe_orders_by_key=nobe_orders_by_key, nobe_act=nobe_act,
        nobe_pipe=nobe_pipe,
    )


def be_eligible(row: pd.Series | dict, dist_has_be: set[str]) -> bool:
    """Predicate matching the BE actuals eligibility (used for trajectory scope)."""
    return (row["_pt"] == "TMT"
            and order_grade(row) is not None
            and row["_ch"] in ("rt", "ss", "pd")
            and row["_dnN"] in dist_has_be)


def month_label_from_value(month_val: str) -> tuple[str, int, int]:
    """'YYYY-MM' -> (label, year, month0)."""
    y, m = (int(x) for x in month_val.split("-"))
    return f"{MOS[m - 1]} {y}", y, m - 1


# ─── Phase 4 helpers — BE-vs-Actuals table, comparison, state gap ──────────
def adjusted_be(be_qty: float, today: datetime,
                month_start: datetime, month_end: datetime) -> float:
    """Pro-rate BE through (today − 1).

    For past months → returns full BE (month is complete).
    For the current month → BE × (days_elapsed_through_yesterday / days_in_month).
    For future months → returns 0.
    """
    if today.date() > month_end.date():
        return float(be_qty)
    if today.date() <= month_start.date():
        return 0.0
    days_in_month = (month_end - month_start).days + 1
    days_elapsed = (today.date() - month_start.date()).days  # through yesterday
    days_elapsed = max(0, min(days_in_month, days_elapsed))
    return float(be_qty) * days_elapsed / days_in_month


def be_table(df: pd.DataFrame, ag: BeAggregate,
             today: datetime | None = None) -> pd.DataFrame:
    """Per-distributor BE-vs-Actuals table.

    Includes every distributor with BE (matched + unmatched) plus distributors
    with eligible orders but no BE (Target=0). Columns:
      Distributor, State, BE, Adjusted BE, Actuals, Absolute gap,
      Adjusted gap, Pending release, Pending invoice, Pending pipeline.
    """
    if today is None:
        today = datetime.now()

    # Eligible activity universe: TMT + Fe 550/550D + retail/SS/PD
    elig = df[
        (df["_pt"] == "TMT")
        & df["_gr"].astype(str).str.lower().str.replace(" ", "", regex=False)
            .isin(("fe550", "fe550d"))
        & df["_ch"].isin(["rt", "ss", "pd"])
    ].copy() if len(df) else df

    rows: list[dict] = []

    # 1) Every BE distributor (matched + unmatched in actuals)
    for atom in ag.atomic:
        distN = atom["distNorm"]
        orders = ag.orders_by_key.get(distN, [])
        actuals = sum(o["invQty"] for o in orders)
        retail_act = sum(o["invQty"] for o in orders
                         if be_bucket(o["r"]["_ch"]) == "Retail")
        project_act = float(actuals) - retail_act
        dist_rows = elig[elig["_dnN"] == distN] if len(elig) else elig
        pend_rel = float(dist_rows["_pend"].sum()) if len(dist_rows) else 0.0
        pend_inv = float(dist_rows["_pendInv"].sum()) if len(dist_rows) else 0.0
        be_qty = float(atom["qty"])
        adj_be = adjusted_be(be_qty, today,
                             ag.be_month_start, ag.be_month_end)
        rows.append({
            "Distributor": atom["dist"],
            "State": atom.get("state", "") or "—",
            "BE": be_qty,
            "Retail BE": float(atom.get("retail_be", 0.0)),
            "Project BE": float(atom.get("project_be", 0.0)),
            "Adjusted BE": adj_be,
            "Actuals": float(actuals),
            "Retail Act": retail_act,
            "Project Act": project_act,
            "Absolute gap": float(actuals) - be_qty,
            "Adjusted gap": float(actuals) - adj_be,
            "Pending release": pend_rel,
            "Pending invoice": pend_inv,
            "Pending pipeline": pend_rel + pend_inv,
            "_distNorm": distN,
            "_hasBE": True,
        })

    # 2) Distributors with eligible orders but NO BE — Target 0.
    # Actuals use the SAME proportional invoice-date attribution as the BE
    # distributors (via ag.nobe_orders_by_key), so the table's total Actuals
    # equals ag.total_act and reconciles with the invoiced KPI for this scope.
    if len(elig):
        no_be = elig[~elig["_dnN"].isin(ag.dist_has_be)]
        for distN, g in no_be.groupby("_dnN"):
            orders = ag.nobe_orders_by_key.get(distN, [])
            actuals_in_month = float(sum(o["invQty"] for o in orders))
            retail_act = sum(o["invQty"] for o in orders
                             if be_bucket(o["r"]["_ch"]) == "Retail")
            project_act = actuals_in_month - retail_act
            pend_rel = float(g["_pend"].sum())
            pend_inv = float(g["_pendInv"].sum())
            rows.append({
                "Distributor": cl(g["_dn"].iloc[0]) or "Direct",
                "State": cl(g["_st"].iloc[0]).title() or "—",
                "BE": 0.0,
                "Retail BE": 0.0,
                "Project BE": 0.0,
                "Adjusted BE": 0.0,
                "Actuals": actuals_in_month,
                "Retail Act": retail_act,
                "Project Act": project_act,
                "Absolute gap": actuals_in_month,
                "Adjusted gap": actuals_in_month,
                "Pending release": pend_rel,
                "Pending invoice": pend_inv,
                "Pending pipeline": pend_rel + pend_inv,
                "_distNorm": distN,
                "_hasBE": False,
            })
    out = pd.DataFrame(rows)
    if not len(out):
        return out
    return out.sort_values("BE", ascending=False).reset_index(drop=True)


def _month_bounds(y: int, m0: int, offset: int) -> tuple[datetime, datetime, str]:
    """Bounds + label for the month `offset` months before (y, m0), m0 0-indexed.
    offset 0 = that month, 1 = prior month, …  Label like \"Jun'26\"."""
    idx = y * 12 + m0 - offset
    yy, mm = divmod(idx, 12)  # mm 0-indexed
    start = datetime(yy, mm + 1, 1)
    end = ((datetime(yy, mm + 2, 1) if mm < 11 else datetime(yy + 1, 1, 1))
           - timedelta(seconds=1))
    return start, end, start.strftime("%b'%y")


def be_mom(df: pd.DataFrame, ag: BeAggregate,
           inv_index: dict[str, InvoiceEntry], months: int = 3) -> pd.DataFrame:
    """Per-distributor invoiced actuals split Retail/Project for the BE month and
    the prior (months-1) months, attributed by invoice date.

    Returns a frame keyed by `_distNorm` with columns "<MonLabel> Retail" /
    "<MonLabel> Project" (most recent month first) plus "Realistic BE (auto)" =
    mean of each month's total (Retail+Project) actuals.
    """
    elig = df[
        (df["_pt"] == "TMT")
        & df["_gr"].astype(str).str.lower().str.replace(" ", "", regex=False)
            .isin(("fe550", "fe550d"))
        & df["_ch"].isin(["rt", "ss", "pd"])
    ].copy() if len(df) else df

    be_m0 = ag.be_month_start.month - 1  # 0-indexed
    be_y = ag.be_month_start.year
    windows = [_month_bounds(be_y, be_m0, off) for off in range(months)]
    labels = [w[2] for w in windows]
    base = {f"{lbl} {b}": 0.0 for lbl in labels for b in ("Retail", "Project")}

    data_map: dict[str, dict] = {}
    if len(elig):
        for _, r in elig.iterrows():
            bucket = be_bucket(r["_ch"])
            rec = data_map.setdefault(r["_dnN"], dict(base))
            for start, end, lbl in windows:
                q = invoiced_in_range(r, start, end, inv_index)
                if q:
                    rec[f"{lbl} {bucket}"] += q

    cols = (["_distNorm"]
            + [f"{lbl} {b}" for lbl in labels for b in ("Retail", "Project")]
            + ["Realistic BE (auto)"])
    rows = []
    for distN, rec in data_map.items():
        month_tot = [rec[f"{lbl} Retail"] + rec[f"{lbl} Project"] for lbl in labels]
        realistic = sum(month_tot) / len(month_tot) if month_tot else 0.0
        rows.append({"_distNorm": distN, **rec, "Realistic BE (auto)": realistic})
    return pd.DataFrame(rows, columns=cols) if rows else pd.DataFrame(columns=cols)


def be_state_gap(df: pd.DataFrame, ag: BeAggregate, today: datetime,
                 mode: str = "abs") -> pd.DataFrame:
    """Aggregate the BE-vs-Actuals table to (state, value, be, actuals).

    `mode`: "abs" → value = Actuals − AdjBE (MT). "pct" → value = (gap/BE)*100,
    NaN when BE = 0.
    """
    table = be_table(df, ag, today)
    if not len(table):
        return pd.DataFrame(columns=["state", "value", "be", "actuals"])
    g = table.groupby("State").agg(
        be=("BE", "sum"), actuals=("Actuals", "sum"),
        adj_be=("Adjusted BE", "sum")).reset_index()
    g = g[g["State"].astype(str).str.strip() != ""]
    g = g[g["State"] != "—"]
    if mode == "pct":
        g["value"] = (g["actuals"] - g["be"]) / g["be"].replace(0, float("nan")) * 100.0
    else:
        g["value"] = g["actuals"] - g["adj_be"]
    return g.rename(columns={"State": "state"})[
        ["state", "value", "be", "actuals"]]


def be_compare_table(rows_a: list[dict], rows_b: list[dict]) -> pd.DataFrame:
    """Side-by-side comparison of two BE versions at distributor level."""
    a_atomic = {r["distNorm"]: r for r in flatten_be_atomic(rows_a)}
    b_atomic = {r["distNorm"]: r for r in flatten_be_atomic(rows_b)}
    keys = set(a_atomic) | set(b_atomic)
    rows: list[dict] = []
    for k in keys:
        a = a_atomic.get(k, {})
        b = b_atomic.get(k, {})
        be_a = float(a.get("qty", 0.0))
        be_b = float(b.get("qty", 0.0))
        rows.append({
            "Distributor": (b.get("dist") or a.get("dist") or ""),
            "State": (b.get("state") or a.get("state") or "—"),
            "BE A": be_a, "BE B": be_b,
            "Δ (B-A)": be_b - be_a,
            "% change": ((be_b - be_a) / be_a * 100.0) if be_a else float("nan"),
        })
    out = pd.DataFrame(rows)
    if not len(out):
        return out
    return out.sort_values("Δ (B-A)", key=lambda s: s.abs(),
                           ascending=False).reset_index(drop=True)
