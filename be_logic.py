"""Best Estimate (BE) Excel parsing and plan-vs-actual aggregation.

BE matching is distributor-only: all BE rows for a distributor (across states,
grades, categories) are clubbed into a single number, and any eligible order
from that distributor counts as its actuals wherever it ships.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any

import pandas as pd

from data import InvoiceEntry, MOS, invoiced_in_range, norm_name, num


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
    """Club BE rows: ONE entry per distributor (summed across states/grades/cats)."""
    grp: dict[str, dict] = {}
    for b in be_rows:
        row = grp.get(b["distNorm"])
        if row is None:
            row = {"dist": b["distributor"], "distNorm": b["distNorm"],
                   "region": b["region"], "qty": 0.0}
            grp[b["distNorm"]] = row
        row["qty"] += (b.get("retail_fe550", 0) + b.get("retail_fe550d", 0)
                       + b.get("project_fe550", 0) + b.get("project_fe550d", 0))
    return [a for a in grp.values() if a["qty"] > 0]


def order_cat(row: pd.Series | dict) -> str | None:
    ot = row["_ot"]
    if ot in ("Retailer", "Self-stocking"):
        return "Retail+PTR"
    if ot == "Project":
        return "Project-thru-Dist"
    return None


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
    matched_act = matched_pipe = 0.0
    for _, r in df.iterrows():
        if r["_pt"] != "TMT":
            continue
        if order_grade(r) is None:
            continue
        if r["_ch"] not in ("rt", "ss", "pd"):
            continue
        if r["_dnN"] not in dist_has_be:
            continue
        inv_qty = invoiced_in_range(r, be_month_start, be_month_end, inv_index)
        pipe_eff = 0.0
        if r["_sta"] != "Cancelled":
            pipe = max(r["_q"] - r["_iq"], 0.0)
            if pipe > 0:
                if pipe < 5:
                    pipe_eff = 0.0
                elif pd.notna(r["_d"]):
                    days_old = (today - r["_d"]).days
                    pipe_eff = 0.0 if days_old > 60 else pipe
                else:
                    pipe_eff = pipe
        if inv_qty <= 0 and pipe_eff <= 0:
            continue
        matched_act += inv_qty
        matched_pipe += pipe_eff
        orders_by_key.setdefault(r["_dnN"], []).append(
            {"r": r, "invQty": inv_qty, "pipe": pipe_eff})

    active = set(orders_by_key.keys())
    unmatched_be = [a for a in atomic if a["distNorm"] not in active]
    tot_be = sum(a["qty"] for a in atomic)

    return BeAggregate(
        atomic=atomic, atom_lookup=atom_lookup, orders_by_key=orders_by_key,
        tot_be=tot_be, matched_act=matched_act, matched_pipe=matched_pipe,
        unmatched_be=unmatched_be, dist_has_be=dist_has_be,
        be_month_start=be_month_start, be_month_end=be_month_end,
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
