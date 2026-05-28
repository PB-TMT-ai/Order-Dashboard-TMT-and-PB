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

import numpy as np
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


# --- Vectorised column helpers (used by enrich / build_invoice_index) ---

def _col(df: pd.DataFrame, key: str) -> pd.Series:
    """Source column by logical key, or an empty Series if absent."""
    name = K[key]
    if name in df.columns:
        return df[name]
    return pd.Series([None] * len(df), index=df.index, dtype="object")


def _num_series(s: pd.Series) -> pd.Series:
    """Vectorised num(): strip separators/symbols, coerce to float, NaN -> 0."""
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0.0).astype(float)
    cleaned = s.astype(str).str.strip().str.replace(
        r"[,\s ₹$€£]", "", regex=True)
    return pd.to_numeric(cleaned, errors="coerce").fillna(0.0).astype(float)


def _date_series(s: pd.Series) -> pd.Series:
    """Vectorised parse_date(): returns a datetime64 Series (NaT for missing).

    Uses per-value inference (`format="mixed"`) with day-first preference
    (matches Indian D-M-Y data). Without `format="mixed"`, pandas locks onto a
    single inferred format from the bulk of values and silently coerces every
    differently-formatted row to NaT — which previously made rows with
    "28-4-25" disappear when most rows were ISO.
    """
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    txt = s.astype(str).str.strip()
    txt = txt.where(~txt.isin(["", "nan", "NaT", "0", "None"]), other=None)
    return pd.to_datetime(txt, errors="coerce", format="mixed", dayfirst=True)


def _cl_series(s: pd.Series) -> pd.Series:
    """Vectorised cl(): clean string, sentinel values -> ''."""
    txt = s.astype(str).str.strip()
    return txt.where(~txt.isin(["nan", "NaT", "None"]), other="")


def _to_obj_dates(dt: pd.Series) -> list:
    return [None if pd.isna(x) else x for x in dt]


def build_invoice_index(inv_df: pd.DataFrame) -> dict[str, InvoiceEntry]:
    """Index invoice rows by Order ID, accumulating per-invoice dates and qty."""
    index: dict[str, InvoiceEntry] = {}
    if inv_df is None or not len(inv_df):
        return index
    oids = inv_df["Order ID"].astype(str).str.strip() if "Order ID" in inv_df else \
        pd.Series([""] * len(inv_df))
    dates = _to_obj_dates(_date_series(inv_df["Invoice date"])) if "Invoice date" in inv_df \
        else [None] * len(inv_df)
    qtys = _num_series(inv_df["Invoiced qty"]).tolist() if "Invoiced qty" in inv_df \
        else [0.0] * len(inv_df)
    nums = inv_df["Invoice number"].astype(str).str.strip().tolist() \
        if "Invoice number" in inv_df else [""] * len(inv_df)

    for oid, date, qty, num_id in zip(oids.tolist(), dates, qtys, nums):
        if not oid:
            continue
        entry = index.get(oid)
        if entry is None:
            entry = InvoiceEntry()
            index[oid] = entry
        entry.invs.append({"date": date, "qty": qty, "num": num_id})
        entry.total_qty += qty
        if date is not None:
            if entry.last_date is None or date > entry.last_date:
                entry.last_date = date
            if entry.first_date is None or date < entry.first_date:
                entry.first_date = date
    return index


def enrich(order_df: pd.DataFrame, inv_index: dict[str, InvoiceEntry],
           today: datetime | None = None) -> pd.DataFrame:
    """Enrich raw order rows into a DataFrame mirroring the HTML's `_` fields.

    Fully vectorised — preserves the per-row business rules exactly.
    """
    if not isinstance(order_df, pd.DataFrame):
        order_df = pd.DataFrame(order_df)
    if today is None:
        today = datetime.now()
    today = today.replace(hour=0, minute=0, second=0, microsecond=0)
    today_ts = pd.Timestamp(today)
    n = len(order_df)
    idx = order_df.index

    dt = _date_series(_col(order_df, "date"))
    q = _num_series(_col(order_df, "qty"))
    rq = _num_series(_col(order_df, "rqty"))
    iq = _num_series(_col(order_df, "iqty"))
    cq = _num_series(_col(order_df, "cqty"))
    sta = _cl_series(_col(order_df, "status"))
    ot = _cl_series(_col(order_df, "ot"))
    dis = _cl_series(_col(order_df, "dis"))
    oid = _cl_series(_col(order_df, "oid"))
    dn = _cl_series(_col(order_df, "dname")).replace("", "Direct")

    # P&T detection
    cm_l = _col(order_df, "cm").astype(str).str.lower()
    gr_l = _col(order_df, "gr").astype(str).str.lower().str.replace(" ", "", regex=False)
    pt_mask = (cm_l.str.contains("nippon", na=False)
               | gr_l.str.contains("yst", na=False))
    pt = np.where(pt_mask, "P&T", "TMT")

    # Date-derived keys
    has_date = dt.notna()
    y = pd.Series("", index=idx)
    m = pd.Series("", index=idx)
    w = pd.Series("", index=idx)
    if has_date.any():
        dd = dt[has_date]
        y.loc[has_date] = dd.dt.year.astype(int).astype(str)
        m.loc[has_date] = dd.dt.strftime("%Y-%m")
        iso = dd.dt.isocalendar()
        w.loc[has_date] = (iso["year"].astype(int).astype(str) + "-W"
                           + iso["week"].astype(int).astype(str).str.zfill(2))

    # Channel
    ch = np.full(n, "rt", dtype=object)
    ch[ot.values == "Self-stocking"] = "ss"
    proj = ot.values == "Project"
    ch[proj & (dis.values == "Yes")] = "pd"
    ch[proj & (dis.values != "Yes")] = "pdir"

    # Bill-to state from GST code
    gst = _col(order_df, "bGST").astype(str).str.strip()
    code = pd.to_numeric(gst.str[:2], errors="coerce")
    bs = code.map(lambda c: SC_.get(int(c), "") if pd.notna(c) else "")

    # Normalised distributor name
    dn_norm = (dn.str.upper().str.replace(_NAME_PUNCT, " ", regex=True)
               .str.replace(_NAME_TOKENS, "", regex=True).str.strip())

    # Pending + short-close
    raw_pend = np.where(sta.values == "Cancelled", 0.0,
                        np.maximum(q.values - rq.values - cq.values, 0.0))
    days_old = (today_ts - dt).dt.days
    short_small = (raw_pend > 0) & (raw_pend < 5)
    short_old = (raw_pend >= 5) & has_date.values & (days_old.values > 60)
    sc_short = short_small | short_old
    pend = np.where(sc_short, 0.0, raw_pend)
    pend_inv = np.maximum(rq.values - iq.values, 0.0)

    # Invoice context (per order id)
    tot_inv = oid.map(lambda o: inv_index[o].total_qty if o in inv_index else 0.0)
    last_d = pd.to_datetime(
        oid.map(lambda o: inv_index[o].last_date if o in inv_index else None),
        errors="coerce")
    first_d = pd.to_datetime(
        oid.map(lambda o: inv_index[o].first_date if o in inv_index else None),
        errors="coerce")

    return pd.DataFrame({
        "_d": dt.to_numpy(), "_q": q.values, "_rq": rq.values, "_iq": iq.values,
        "_cq": cq.values, "_pt": pt, "_w": w.values, "_m": m.values, "_y": y.values,
        "_st": _cl_series(_col(order_df, "sts")).values,
        "_ct": _cl_series(_col(order_df, "stc")).values,
        "_gr": _cl_series(_col(order_df, "gr")).values,
        "_dia": _cl_series(_col(order_df, "dia")).values,
        "_fm": _cl_series(_col(order_df, "frm")).values,
        "_p2": _cl_series(_col(order_df, "ptm")).values,
        "_dl": _cl_series(_col(order_df, "dlm")).values,
        "_cm": _cl_series(_col(order_df, "cm")).values,
        "_ws": _cl_series(_col(order_df, "ws")).values,
        "_dis": dis.values, "_dn": dn.values, "_dnN": dn_norm.values,
        "_sf": _cl_series(_col(order_df, "sfs")).values, "_bs": bs.values,
        "_sta": sta.values, "_ot": ot.values, "_oid": oid.values, "_ch": ch,
        "_pendOrig": raw_pend, "_pend": pend, "_scShort": sc_short,
        "_pendInv": pend_inv, "_invDateLast": last_d.to_numpy(),
        "_invDateFirst": first_d.to_numpy(), "_orderTotInv": tot_inv.values,
    })


def invoiced_in_range(row: pd.Series | dict, from_date: datetime, to_date: datetime,
                      inv_index: dict[str, InvoiceEntry]) -> float:
    """Invoiced qty attributed to [from_date, to_date], proportionally allocated."""
    oid = row["_oid"]
    e = inv_index.get(oid)
    iq = row["_iq"]
    if e is None or e.total_qty == 0:
        d = row["_d"]
        if iq > 0 and pd.notna(d) and from_date <= d <= to_date:
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


def invoiced_in_period(df: pd.DataFrame, from_date: datetime, to_date: datetime,
                       inv_index: dict[str, InvoiceEntry]) -> pd.Series:
    """Vectorised invoice-in-range attribution over a DataFrame.

    Equivalent to applying invoiced_in_range() per row, but fast: one pass over
    the invoice index, then vectorised allocation across order rows.
    """
    if not len(df):
        return pd.Series([], dtype=float)
    # Per-order allocation factor = (qty invoiced within range) / (order total invoiced)
    factor: dict[str, float] = {}
    for oid, e in inv_index.items():
        if e.total_qty == 0:
            continue
        s = 0.0
        for inv in e.invs:
            d = inv["date"]
            if d is not None and from_date <= d <= to_date:
                s += inv["qty"]
        if s:
            factor[oid] = s / e.total_qty

    oids = df["_oid"]
    iq = df["_iq"].to_numpy(dtype=float)
    has_entry = oids.map(lambda o: o in inv_index and inv_index[o].total_qty > 0).to_numpy()
    fac = oids.map(lambda o: factor.get(o, 0.0)).to_numpy(dtype=float)

    d = df["_d"]
    in_range = (d.notna() & (d >= from_date) & (d <= to_date)).to_numpy()
    fallback = np.where((~has_entry) & (iq > 0) & in_range, iq, 0.0)
    return pd.Series(np.where(has_entry, fac * iq, fallback), index=df.index)


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
