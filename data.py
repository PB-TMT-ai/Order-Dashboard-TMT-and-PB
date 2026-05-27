"""Order Excel parsing, enrichment, channel classification, and filtering.

Faithful Python port of the data layer from the HTML dashboard. Business rules
(channel classification, P&T vs TMT, short-close, invoice-date attribution) are
preserved exactly.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any, Callable

import pandas as pd

# --- Source column names in the Order sheet (mirror of K{} in the HTML) ---
K = {
    "date": "Opportunity date",
    "qty": "Quantity",
    "rqty": "Release Qty",
    "iqty": "Invoiced Qty",
    "cqty": "total cancelled qty",
    "status": "Order Status",
    "cm": "CM name",
    "sts": "Ship to State",
    "stc": "Ship to city",
    "gr": "Grade",
    "dia": "Diameter mm",
    "frm": "Form",
    "ptm": "Payment Terms",
    "dlm": "Initial delivery Mode",
    "ot": "Order Type",
    "dis": "Distributor (Yes/No)",
    "ws": "ad_ship_to_address_type_c",
    "sfs": "Ship_From_State",
    "bGST": "Bill to - GST",
    "dname": "Distributor Name",
    "oid": "Order ID",
}

# GST state code -> state name
SC_ = {
    1: "Jammu & Kashmir", 2: "Himachal Pradesh", 3: "Punjab", 4: "Chandigarh",
    5: "Uttarakhand", 6: "Haryana", 7: "Delhi", 8: "Rajasthan", 9: "Uttar Pradesh",
    10: "Bihar", 11: "Sikkim", 12: "Arunachal Pradesh", 13: "Nagaland",
    14: "Manipur", 15: "Mizoram", 16: "Tripura", 17: "Meghalaya", 18: "Assam",
    19: "West Bengal", 20: "Jharkhand", 21: "Odisha", 22: "Chhattisgarh",
    23: "Madhya Pradesh", 24: "Gujarat", 27: "Maharashtra", 29: "Karnataka",
    30: "Goa", 32: "Kerala", 33: "Tamil Nadu", 34: "Puducherry", 36: "Telangana",
    37: "Andhra Pradesh",
}

CHANNEL_LABELS = {
    "rt": "Retail",
    "ss": "Self-stocking",
    "pdir": "Project (direct)",
    "pd": "Project thru Dist",
}

MOS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

_NUM_STRIP = re.compile(r"[,\s ₹$€£]")
_NAME_PUNCT = re.compile(r"[\.\,\(\)\-\'\"\s]+")
_NAME_TOKENS = re.compile(r"\b(PVT|PRIVATE|LTD|LIMITED|LLP|CO|COMPANY|INDIA)\b")


def num(v: Any) -> float:
    """Parse a numeric cell, stripping thousand separators and currency symbols."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if pd.isna(v) else float(v)
    s = str(v).strip()
    if not s or s in ("-", "—") or s.lower() == "nan":
        return 0.0
    cleaned = _NUM_STRIP.sub("", s)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_date(v: Any) -> datetime | None:
    """Parse a date cell. Accepts datetimes, ISO strings, and d/m/y strings."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return None if pd.isna(v) else v
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    s = str(v).strip()
    if not s or s in ("nan", "NaT", "0", "None"):
        return None
    dt = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if pd.notna(dt):
        return dt.to_pydatetime()
    parts = s.split("/")
    if len(parts) == 3:
        try:
            return datetime(int(parts[2]), int(parts[1]), int(parts[0]))
        except (ValueError, IndexError):
            return None
    return None


def cl(v: Any) -> str:
    """Clean a string cell, treating NaN/NaT/None sentinels as empty."""
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s in ("nan", "NaT", "None") else s


def norm_name(s: Any) -> str:
    """Normalise a distributor name for matching (uppercase, drop suffixes)."""
    up = str(s or "").upper()
    up = _NAME_PUNCT.sub(" ", up)
    up = _NAME_TOKENS.sub("", up)
    return up.strip()


def is_pt(row: dict) -> bool:
    """P&T detection: Nippon CM or grade starting/containing 'yst'."""
    cm = str(row.get(K["cm"], "") or "").lower()
    gr = str(row.get(K["gr"], "") or "").lower().replace(" ", "")
    return "nippon" in cm or gr.startswith("yst") or "yst" in gr


def bill_state(row: dict) -> str:
    g = str(row.get(K["bGST"], "") or "").strip()
    try:
        code = int(g[:2])
    except (ValueError, IndexError):
        return ""
    return SC_.get(code, "")


def week_key(d: datetime | None) -> str:
    if not d:
        return ""
    iso = d.isocalendar()
    return f"{iso[0]}-W{iso[1]:02d}"


def order_channel(ot: str, dis: str) -> str:
    ot = str(ot or "").strip()
    dis = str(dis or "").strip()
    if ot == "Self-stocking":
        return "ss"
    if ot == "Project" and dis == "Yes":
        return "pd"
    if ot == "Project":
        return "pdir"
    return "rt"  # Retailer + blank/unrecognized -> Retail


@dataclass
class InvoiceEntry:
    total_qty: float = 0.0
    invs: list[dict] = field(default_factory=list)
    last_date: datetime | None = None
    first_date: datetime | None = None


def build_invoice_index(inv_rows: list[dict]) -> dict[str, InvoiceEntry]:
    """Index invoice rows by Order ID, accumulating per-invoice dates and qty."""
    index: dict[str, InvoiceEntry] = {}
    for inv in inv_rows:
        oid = str(inv.get("Order ID", "") or "").strip()
        if not oid:
            continue
        date = parse_date(inv.get("Invoice date"))
        qty = num(inv.get("Invoiced qty"))
        num_id = str(inv.get("Invoice number", "") or "").strip()
        entry = index.get(oid)
        if entry is None:
            entry = InvoiceEntry()
            index[oid] = entry
        entry.invs.append({"date": date, "qty": qty, "num": num_id})
        entry.total_qty += qty
        if date:
            if entry.last_date is None or date > entry.last_date:
                entry.last_date = date
            if entry.first_date is None or date < entry.first_date:
                entry.first_date = date
    return index


def enrich(order_rows: list[dict], inv_index: dict[str, InvoiceEntry],
           today: datetime | None = None) -> pd.DataFrame:
    """Enrich raw order rows into a DataFrame mirroring the HTML's `_` fields."""
    if today is None:
        today = datetime.now()
    today = today.replace(hour=0, minute=0, second=0, microsecond=0)

    recs: list[dict] = []
    for r in order_rows:
        d = parse_date(r.get(K["date"]))
        q = num(r.get(K["qty"]))
        rq = num(r.get(K["rqty"]))
        iq = num(r.get(K["iqty"]))
        cq = num(r.get(K["cqty"]))
        sta = cl(r.get(K["status"]))
        ot = cl(r.get(K["ot"]))
        dis = cl(r.get(K["dis"]))
        oid = cl(r.get(K["oid"]))
        dn = cl(r.get(K["dname"])) or "Direct"

        raw_pend = 0.0 if sta == "Cancelled" else max(q - rq - cq, 0.0)
        pend = raw_pend
        sc_short = False
        if raw_pend > 0:
            if raw_pend < 5:
                pend = 0.0
                sc_short = True
            elif d is not None:
                days_old = (today - d).days
                if days_old > 60:
                    pend = 0.0
                    sc_short = True

        e = inv_index.get(oid)
        recs.append({
            "_d": d,
            "_q": q, "_rq": rq, "_iq": iq, "_cq": cq,
            "_pt": "P&T" if is_pt(r) else "TMT",
            "_w": week_key(d), "_m": d.strftime("%Y-%m") if d else "",
            "_y": str(d.year) if d else "",
            "_st": cl(r.get(K["sts"])), "_ct": cl(r.get(K["stc"])),
            "_gr": cl(r.get(K["gr"])), "_dia": cl(r.get(K["dia"])),
            "_fm": cl(r.get(K["frm"])), "_p2": cl(r.get(K["ptm"])),
            "_dl": cl(r.get(K["dlm"])), "_cm": cl(r.get(K["cm"])),
            "_ws": cl(r.get(K["ws"])), "_dis": dis,
            "_dn": dn, "_dnN": norm_name(dn),
            "_sf": cl(r.get(K["sfs"])), "_bs": bill_state(r),
            "_sta": sta, "_ot": ot, "_oid": oid,
            "_ch": order_channel(ot, dis),
            "_pendOrig": raw_pend, "_pend": pend, "_scShort": sc_short,
            "_pendInv": max(rq - iq, 0.0),
            "_invDateLast": e.last_date if e else None,
            "_invDateFirst": e.first_date if e else None,
            "_orderTotInv": e.total_qty if e else 0.0,
        })
    return pd.DataFrame.from_records(recs)


def invoiced_in_range(row: pd.Series | dict, from_date: datetime, to_date: datetime,
                      inv_index: dict[str, InvoiceEntry]) -> float:
    """Invoiced qty attributed to [from_date, to_date], proportionally allocated."""
    oid = row["_oid"]
    e = inv_index.get(oid)
    iq = row["_iq"]
    if e is None or e.total_qty == 0:
        d = row["_d"]
        if iq > 0 and d is not None and from_date <= d <= to_date:
            return iq
        return 0.0
    q = 0.0
    for inv in e.invs:
        if inv["date"] is None:
            continue
        if from_date <= inv["date"] <= to_date:
            q += inv["qty"]
    return q * (iq / e.total_qty)


def invoiced_in_range_series(df: pd.DataFrame, from_date: datetime, to_date: datetime,
                             inv_index: dict[str, InvoiceEntry]) -> pd.Series:
    return df.apply(lambda r: invoiced_in_range(r, from_date, to_date, inv_index), axis=1)


def invoiced_daily_map(df: pd.DataFrame, from_date: datetime, to_date: datetime,
                       inv_index: dict[str, InvoiceEntry],
                       filter_fn: Callable[[pd.Series], bool] | None = None) -> dict[str, float]:
    """Map of YYYY-MM-DD -> invoiced MT (proportionally allocated) within range."""
    m: dict[str, float] = {}
    for _, r in df.iterrows():
        if filter_fn and not filter_fn(r):
            continue
        e = inv_index.get(r["_oid"])
        if e is None or e.total_qty == 0:
            continue
        share = r["_iq"] / e.total_qty
        for inv in e.invs:
            if inv["date"] is None:
                continue
            if from_date <= inv["date"] <= to_date:
                k = inv["date"].strftime("%Y-%m-%d")
                m[k] = m.get(k, 0.0) + inv["qty"] * share
    return dict(sorted(m.items()))


# --- Filters ---

# Filter keys that map to multi-select columns
MULTI_FILTERS = {
    "ot": "_ot", "dn": "_dn", "sts": "_st", "stc": "_ct", "bts": "_bs",
    "gr": "_gr", "dia": "_dia", "frm": "_fm", "ptm": "_p2", "dlm": "_dl",
    "cm": "_cm", "ws": "_ws", "stat": "_sta",
}


def default_filters() -> dict:
    return {
        "pt": "All", "dm": "lx", "lx": 9999, "mo": "", "yr": "", "df": "", "dt": "",
        "dis": "All",
        **{k: [] for k in MULTI_FILTERS},
    }


def get_active_period(f: dict, today: datetime | None = None) -> dict | None:
    """Return {from, to, label} for the active sidebar date selection, or None for All."""
    if today is None:
        today = datetime.now()
    today = today.replace(hour=23, minute=59, second=59, microsecond=999000)
    dm = f.get("dm", "lx")
    if dm == "lx":
        lx = int(f.get("lx", 9999))
        if lx >= 9999:
            return None
        return {"from": today - timedelta(days=lx), "to": today, "label": f"Last {lx} days"}
    if dm == "mo" and f.get("mo"):
        y, m = (int(x) for x in f["mo"].split("-"))
        frm = datetime(y, m, 1)
        to = (datetime(y, m + 1, 1) if m < 12 else datetime(y + 1, 1, 1)) - timedelta(seconds=1)
        return {"from": frm, "to": to, "label": f"{MOS[m - 1]} {y}"}
    if dm == "yr" and f.get("yr"):
        y = int(f["yr"])
        return {"from": datetime(y, 1, 1), "to": datetime(y, 12, 31, 23, 59, 59),
                "label": f"Year {y}"}
    if dm == "rg":
        frm = datetime.fromisoformat(f["df"]) if f.get("df") else None
        to = datetime.fromisoformat(f["dt"] + "T23:59:59") if f.get("dt") else None
        return {"from": frm, "to": to,
                "label": f"{f.get('df') or '…'} → {f.get('dt') or '…'}"}
    return None


def _multi_mask(df: pd.DataFrame, f: dict) -> pd.Series:
    mask = pd.Series(True, index=df.index)
    if f.get("pt", "All") != "All":
        mask &= df["_pt"] == f["pt"]
    if f.get("dis", "All") != "All":
        mask &= df["_dis"] == f["dis"]
    for key, col in MULTI_FILTERS.items():
        vals = f.get(key) or []
        if vals:
            mask &= df[col].isin(set(vals))
    return mask


def apply_non_date_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    """All sidebar filters EXCEPT the date range (used for invoice-in-period KPI)."""
    return df[_multi_mask(df, f)]


def apply_filters(df: pd.DataFrame, f: dict, today: datetime | None = None) -> pd.DataFrame:
    """Apply the full sidebar filter set (date range + all multi-selects)."""
    mask = _multi_mask(df, f)

    period = get_active_period(f, today)
    if period and (period["from"] is not None or period["to"] is not None):
        d = df["_d"]
        has_date = d.notna()
        if period["from"] is not None:
            mask &= (~has_date) | (d >= period["from"])
        if period["to"] is not None:
            mask &= (~has_date) | (d <= period["to"])
    return df[mask]


@dataclass
class Kpis:
    ordered: float
    released: float
    invoiced: float
    line_items: int
    ch_or: dict
    ch_re: dict
    ch_in: dict


def compute_kpis(df: pd.DataFrame) -> Kpis:
    channels = ["rt", "ss", "pdir", "pd"]
    ch_or = {c: 0.0 for c in channels}
    ch_re = {c: 0.0 for c in channels}
    ch_in = {c: 0.0 for c in channels}
    if len(df):
        g = df.groupby("_ch")
        for c in channels:
            if c in g.groups:
                sub = g.get_group(c)
                ch_or[c] = float(sub["_q"].sum())
                ch_re[c] = float(sub["_rq"].sum())
                ch_in[c] = float(sub["_iq"].sum())
    return Kpis(
        ordered=float(df["_q"].sum()) if len(df) else 0.0,
        released=float(df["_rq"].sum()) if len(df) else 0.0,
        invoiced=float(df["_iq"].sum()) if len(df) else 0.0,
        line_items=len(df),
        ch_or=ch_or, ch_re=ch_re, ch_in=ch_in,
    )


def aggregate(df: pd.DataFrame, group_col: str, metric_col: str, top: int | None = None,
              ascending: bool = False) -> pd.DataFrame:
    """Group by a single dimension and sum a metric. Used by charts and drill tables."""
    if not len(df):
        return pd.DataFrame(columns=[group_col, metric_col])
    out = (df.groupby(group_col)[metric_col].sum()
           .reset_index().sort_values(metric_col, ascending=ascending))
    if top:
        out = out.head(top)
    return out


def fmt(v: float) -> str:
    """Indian-style integer formatting (lakh/crore grouping)."""
    n = round(float(v or 0))
    sign = "-" if n < 0 else ""
    s = str(abs(n))
    if len(s) <= 3:
        return sign + s
    head, tail = s[:-3], s[-3:]
    head = re.sub(r"(\d)(?=(\d\d)+$)", r"\1,", head)
    return f"{sign}{head},{tail}"
