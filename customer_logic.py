"""Customer (distributor/account) buying-pattern analytics.

Operates on the enriched order frame from data.enrich() (the rejected/cancelled
qty is already netted out of each order's quantity).
The customer key is the distributor display name `_dn` (blank → "Direct").
Provides RFM segmentation, product mix, reorder cadence / churn flags and
month-over-month growth — consumed by the Customer tab in app.py.
"""
from __future__ import annotations

from datetime import datetime

import pandas as pd

from plots import form_label

# Channel code -> readable label (mirrors data.CHANNEL_LABELS ordering).
_CH_LABEL = {"rt": "Retail", "ss": "Self-stocking",
             "pdir": "Project (direct)", "pd": "Project thru Dist"}


def _score(series: pd.Series, higher_is_better: bool, q: int = 3) -> pd.Series:
    """Quantile score 1..q (q = best). Falls back to the middle bucket when the
    data can't be split into q quantiles (too few distinct values)."""
    if not len(series):
        return pd.Series(dtype=int)
    try:
        binned = pd.qcut(series.rank(method="first"), q, labels=False) + 1
    except (ValueError, IndexError):
        binned = pd.Series((q + 1) // 2, index=series.index)
    if not higher_is_better:
        binned = q + 1 - binned
    return binned.astype(int)


def _segment(r: int, f: int) -> str:
    """Coarse RFM segment from Recency and Frequency scores (1..3)."""
    if r >= 3 and f >= 3:
        return "Champion"
    if r >= 3 and f <= 2:
        return "New / Promising"
    if r <= 2 and f >= 3:
        return "At risk"
    if r == 1 and f == 1:
        return "Dormant"
    return "Needs attention"


def rfm(df: pd.DataFrame, today: datetime | None = None) -> pd.DataFrame:
    """Per-customer Recency (days since last order), Frequency (distinct orders)
    and Monetary (total ordered MT), with 1..3 scores and a segment label."""
    if today is None:
        today = datetime.now()
    if not len(df):
        return pd.DataFrame(columns=[
            "Customer", "Recency (days)", "Frequency", "Monetary (MT)",
            "R", "F", "M", "Segment"])
    g = df.groupby("_dn").agg(
        last_order=("_d", "max"),
        Frequency=("_oid", "nunique"),
        Monetary=("_q", "sum"),
    ).reset_index()
    g["Recency (days)"] = (pd.Timestamp(today) - g["last_order"]).dt.days
    g = g.rename(columns={"_dn": "Customer", "Monetary": "Monetary (MT)"})
    g["R"] = _score(g["Recency (days)"], higher_is_better=False)
    g["F"] = _score(g["Frequency"], higher_is_better=True)
    g["M"] = _score(g["Monetary (MT)"], higher_is_better=True)
    g["Segment"] = [_segment(r, f) for r, f in zip(g["R"], g["F"])]
    cols = ["Customer", "Recency (days)", "Frequency", "Monetary (MT)",
            "R", "F", "M", "Segment"]
    return g[cols].sort_values("Monetary (MT)", ascending=False).reset_index(drop=True)


def mix(df: pd.DataFrame, dim: str) -> pd.DataFrame:
    """Per-customer share of ordered MT by a dimension.

    dim ∈ {"grade", "form", "channel"}. Returns wide table: Customer + one
    column per dimension value holding that value's % of the customer's MT.
    """
    if not len(df):
        return pd.DataFrame(columns=["Customer"])
    d = df.copy()
    if dim == "grade":
        d["_k"] = d["_gr"].astype(str)
    elif dim == "form":
        d["_k"] = d["_fm"].map(form_label)
    else:  # channel
        d["_k"] = d["_ch"].map(lambda c: _CH_LABEL.get(c, c))
    g = d.groupby(["_dn", "_k"])["_q"].sum().reset_index()
    piv = g.pivot(index="_dn", columns="_k", values="_q").fillna(0.0)
    tot = piv.sum(axis=1).replace(0, pd.NA)
    pct = piv.div(tot, axis=0).mul(100.0).fillna(0.0)
    pct.insert(0, "Total MT", piv.sum(axis=1))
    pct = pct.reset_index().rename(columns={"_dn": "Customer"})
    return pct.sort_values("Total MT", ascending=False).reset_index(drop=True)


def cadence(df: pd.DataFrame, today: datetime | None = None,
            dormant_days: int = 60) -> pd.DataFrame:
    """Per-customer reorder cadence and churn flags.

    Mean/median gap (days) between consecutive order dates, last order, recency,
    and a status: Dormant (recency > dormant_days) / At risk (recency > median
    gap and > 0.5*dormant_days) / Active.
    """
    if today is None:
        today = datetime.now()
    if not len(df):
        return pd.DataFrame(columns=[
            "Customer", "Orders", "Mean gap (days)", "Median gap (days)",
            "Last order", "Recency (days)", "Status"])
    rows = []
    for dn, g in df.groupby("_dn"):
        dates = g["_d"].dropna().sort_values()
        n = dates.nunique()
        gaps = dates.drop_duplicates().diff().dropna().dt.days
        mean_gap = float(gaps.mean()) if len(gaps) else 0.0
        med_gap = float(gaps.median()) if len(gaps) else 0.0
        last = dates.max() if len(dates) else pd.NaT
        rec = (pd.Timestamp(today) - last).days if pd.notna(last) else None
        if rec is None:
            status = "Unknown"
        elif rec > dormant_days:
            status = "Dormant"
        elif med_gap and rec > med_gap and rec > dormant_days / 2:
            status = "At risk"
        else:
            status = "Active"
        rows.append({
            "Customer": dn, "Orders": int(g["_oid"].nunique()),
            "Mean gap (days)": round(mean_gap, 1),
            "Median gap (days)": round(med_gap, 1),
            "Last order": last, "Recency (days)": rec, "Status": status,
        })
    out = pd.DataFrame(rows)
    return out.sort_values("Recency (days)", ascending=False).reset_index(drop=True)


def mom_growth(df: pd.DataFrame, months: int = 6) -> pd.DataFrame:
    """Per-customer monthly ordered MT for the last `months`, plus MoM % change
    between the two most recent months present in the data."""
    if not len(df) or df["_m"].dropna().empty:
        return pd.DataFrame(columns=["Customer"])
    all_months = sorted(m for m in df["_m"].dropna().unique() if m)
    keep = all_months[-months:]
    d = df[df["_m"].isin(keep)]
    g = d.groupby(["_dn", "_m"])["_q"].sum().reset_index()
    piv = (g.pivot(index="_dn", columns="_m", values="_q")
           .reindex(columns=keep).fillna(0.0))
    if len(keep) >= 2:
        prev, last = piv[keep[-2]], piv[keep[-1]]
        piv["MoM %"] = ((last - prev) / prev.replace(0, pd.NA) * 100.0)
        piv["MoM %"] = piv["MoM %"].fillna(0.0)
    piv = piv.reset_index().rename(columns={"_dn": "Customer"})
    sort_col = keep[-1]
    return piv.sort_values(sort_col, ascending=False).reset_index(drop=True)
