"""PB TMT Update -> daily mail pack (plant-wise & grade-wise MTD snapshot).

Builds the daily "PB TMT Update — <d Mon yyyy>" report from one input workbook:

  * PB_TMT_Update_<dd-Mon-yyyy>.png   designed snapshot (the image that goes in the mail)
  * PB_TMT_Update_<dd-Mon-yyyy>.html  same snapshot as a web page
  * PB_TMT_Update_<dd-Mon-yyyy>.xlsx  same numbers as a formatted sheet (values, no formulas)
  * PB_TMT_Update_<dd-Mon-yyyy>.eml   unsent Outlook draft: PNG inline in the body + xlsx
                                      attached; double-click -> Outlook -> Send

Sections: header (invoiced MTD + open-order serving gap line) · 01 Key metrics
(12 cards with per-grade splits) · 02 Plant-wise by grade · 03 Grade-wise summary.

Usage:
  python scripts/pb_mtd_mail.py template <input.xlsx>        # blank input workbook
  python scripts/pb_mtd_mail.py build <input.xlsx> --out <dir>
        [--to a@x.com,b@y.com] [--cc ...] [--sender ...] [--no-png]

Input workbook:
  Meta  : as_on, prepared_by, latest_be, do_released, exp_btr_comp, exp_closing
          (the last three are only fallbacks when the Plant rows leave them blank)
  Plant : Plant, Grade, BE, Exp Orders, Orders MTD, Invoiced, Pending to Serve,
          Physical Inv, Exp BTR Comp, Exp Closing, Inventory Issue, PO Issued,
          Production MTD, DO Released
          Plant blank = same as row above. Rows with no numbers at all are dropped.

Derived, never typed:
  Pending to Serve (if blank) = max(Orders MTD − Invoiced, 0)
  Net to Serve                = Physical Inv − Pending to Serve
  Invoice % of BE, all totals, grade splits, grade-wise summary, header gap line.
"""
from __future__ import annotations

import argparse
import html
import math
import re
import sys
from dataclasses import dataclass, field
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import formatdate, make_msgid
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─── Column contracts ────────────────────────────────────────────────────────
PLANT_COLS = [
    "Plant", "Grade", "BE", "Exp Orders", "Orders MTD", "Invoiced", "Pending to Serve",
    "Physical Inv", "Exp BTR Comp", "Exp Closing", "Inventory Issue", "PO Issued",
    "Production MTD", "DO Released",
]
REQUIRED = ["Plant", "Grade", "BE", "Orders MTD", "Invoiced", "Physical Inv"]
TEXT_COLS = {"Plant", "Grade", "Inventory Issue"}
NUM_COLS = [c for c in PLANT_COLS if c not in TEXT_COLS] + ["Net to Serve"]
# Older sheet headers still accepted
ALIASES = {"Orders": "Orders MTD", "PO Prod": "Production MTD", "Critical Dia": "Inventory Issue",
           "Expected Orders": "Exp Orders", "Exp. Orders": "Exp Orders",
           "Exp. BTR Comp": "Exp BTR Comp", "Exp. Closing": "Exp Closing"}
META_KEYS = {
    "as_on": "Report date (dd-mm-yyyy)",
    "prepared_by": "Prepared by (footer name)",
    "latest_be": "Latest BE (MT) — blank = sum of plant BE",
    "do_released": "DO released (MT) — only if the Plant 'DO Released' column is blank",
    "exp_btr_comp": "Expected BTR completion (MT) — only if Plant column is blank",
    "exp_closing": "Expected closing inventory (MT) — only if Plant column is blank",
}
META_TEXT = {"as_on", "prepared_by"}
GRADES = ["FE 550", "FE 550D", "One Helix"]
GRADE_SHORT = {"FE 550": "550", "FE 550D": "550D", "ONE HELIX": "Helix"}

# Net to Serve / Exp Closing pills: <0 red · 0..green-from amber · >= green-from green
GREEN_FROM_MT = {"Net to Serve": 1000.0, "Exp Closing": 750.0}
PILL_RED, PILL_AMBER, PILL_GREEN = ("FBE1E1", "B42318"), ("FDF1D6", "8A6100"), ("E3F4E8", "1E7B3A")

NAVY, ORANGE = "0E2A47", "E07A2E"


def pill_colours(v: float | None, col: str) -> tuple[str, str] | None:
    if v is None or pd.isna(v):
        return None
    if v < 0:
        return PILL_RED
    return PILL_GREEN if v >= GREEN_FROM_MT[col] else PILL_AMBER


# ─── Model ───────────────────────────────────────────────────────────────────
@dataclass
class Report:
    as_on: date
    meta: dict[str, Any]
    plants: pd.DataFrame          # PLANT_COLS + Net to Serve
    grades: pd.DataFrame          # index = grade, NUM_COLS sums
    total: dict[str, float | None]
    cards: list[tuple[str, float | None, str, str]]   # label, value, kind(mt|pct), split html
    gap_line: str
    warnings: list[str] = field(default_factory=list)

    @property
    def stamp(self) -> str:
        return self.as_on.strftime("%d-%b-%Y")

    @property
    def day(self) -> str:
        return f"{self.as_on.day} {self.as_on.strftime('%b %Y')}"


def _num(v: Any) -> float | None:
    if v is None or (isinstance(v, str) and v.strip() in ("", "-", "–", "—")):
        return None
    if isinstance(v, str):
        neg = v.strip().startswith("(") and v.strip().endswith(")")
        v = v.replace(",", "").replace("(", "").replace(")", "").strip()
        try:
            f = float(v)
        except ValueError:
            return None
        return -f if neg else f
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(f) else f


def _parse_date(v: Any) -> date:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    dt = pd.to_datetime(str(v).strip(), dayfirst=True, errors="coerce")
    if pd.isna(dt):
        raise ValueError(f"Meta.as_on is not a date: {v!r}")
    return dt.date()


def read_input(path: Path) -> tuple[date, dict[str, Any], pd.DataFrame]:
    wb = load_workbook(path, data_only=True)
    for s in ("Meta", "Plant"):
        if s not in wb.sheetnames:
            raise ValueError(f"Input workbook needs a '{s}' sheet (has {wb.sheetnames})")
    raw = {str(r[0]).strip(): r[1] for r in wb["Meta"].iter_rows(min_row=2, values_only=True)
           if r and r[0] is not None}
    if not raw.get("as_on"):
        raise ValueError("Meta.as_on is required")
    meta: dict[str, Any] = {k: _num(raw.get(k)) for k in META_KEYS if k not in META_TEXT}
    meta["prepared_by"] = str(raw.get("prepared_by") or "").strip()

    rows = list(wb["Plant"].iter_rows(values_only=True))
    header = [ALIASES.get(str(c).strip(), str(c).strip()) if c is not None else "" for c in rows[0]]
    missing = [c for c in REQUIRED if c not in header]
    if missing:
        raise ValueError(f"Sheet 'Plant' missing columns: {missing}")
    data = {c: [r[header.index(c)] if c in header and header.index(c) < len(r) else None
                for r in rows[1:]] for c in PLANT_COLS}
    df = pd.DataFrame(data)
    df = df[df.apply(lambda r: any(v not in (None, "") for v in r), axis=1)].reset_index(drop=True)
    return _parse_date(raw["as_on"]), meta, df


# ─── Derivations ─────────────────────────────────────────────────────────────
def _sum(s: pd.Series) -> float | None:
    return None if s.isna().all() else float(s.sum())


def _mt(v: float | None) -> str:
    """Report number: blank -> en dash, negatives in (parentheses)."""
    if v is None or pd.isna(v):
        return "–"
    t = f"{abs(v):,.0f}"
    return f"({t})" if round(v) < 0 else t


def _pct(v: float | None) -> str:
    return "–" if v is None or pd.isna(v) else f"{v * 100:.0f}%"


def _signed(v: float) -> str:
    return f"+{v:,.0f}" if v >= 0 else f"-{abs(v):,.0f}"


def _grade_key(g: str) -> int:
    up = g.upper()
    return next((i for i, x in enumerate(GRADES) if x.upper() == up), 99)


def _short_plant(name: str) -> str:
    """'Ambashakti Udyog – Gwalior' -> 'Gwalior', 'Real Ispat' -> 'Real'."""
    if "–" in name or " - " in name:
        return re.split(r"\s*[–-]\s*", name)[-1]
    return name.split()[0]


def derive(as_on: date, meta: dict[str, Any], plants: pd.DataFrame) -> Report:
    warnings: list[str] = []
    p = plants.copy()
    p["Plant"] = p["Plant"].replace("", None).ffill().astype(str).str.strip()
    p["Grade"] = p["Grade"].fillna("").astype(str).str.strip()
    p["Inventory Issue"] = p["Inventory Issue"].fillna("").astype(str).str.strip()
    vals = [c for c in PLANT_COLS if c not in TEXT_COLS]
    for c in vals:
        p[c] = p[c].map(_num).astype(float)
    p = p[p[vals].fillna(0).ne(0).any(axis=1)].reset_index(drop=True)   # drop all-blank rows

    calc = (p["Orders MTD"].fillna(0) - p["Invoiced"].fillna(0)).clip(lower=0)
    blank = p["Pending to Serve"].isna() & p["Orders MTD"].notna()
    p.loc[blank, "Pending to Serve"] = calc[blank]
    p.loc[p["Pending to Serve"] == 0, "Pending to Serve"] = None
    p["Net to Serve"] = p["Physical Inv"].fillna(0) - p["Pending to Serve"].fillna(0)
    p.loc[p["Physical Inv"].isna() & p["Pending to Serve"].isna(), "Net to Serve"] = None

    g = p.groupby("Grade", sort=False)[NUM_COLS].agg(_sum)
    g = g.loc[sorted(g.index, key=_grade_key)]
    total = {c: _sum(p[c]) for c in NUM_COLS}

    be = meta.get("latest_be") or total["BE"]
    if meta.get("latest_be") and total["BE"] and abs(meta["latest_be"] - total["BE"]) > 0.5:
        warnings.append(f"Meta.latest_be {meta['latest_be']:,.0f} ≠ plant BE total {total['BE']:,.0f}")
    for col, key in (("DO Released", "do_released"), ("Exp BTR Comp", "exp_btr_comp"),
                     ("Exp Closing", "exp_closing")):
        if total[col] is None and meta.get(key) is not None:
            total[col] = meta[key]            # card shows the total, no grade split
    for col in ("Exp Orders", "Exp BTR Comp", "Exp Closing", "DO Released", "PO Issued",
                "Production MTD"):
        if p[col].isna().all():
            warnings.append(f"'{col}' is blank for every row"
                            + (" — card uses Meta total" if total[col] is not None else " — shown as –"))
    if not p["Inventory Issue"].any():
        warnings.append("Inventory Issue is blank for every row")

    def split(col: str, *, pct: bool = False) -> str:
        if p[col].isna().all():
            return ""
        out = []
        for grade, r in g.iterrows():
            v = (r["Invoiced"] / r["BE"] if r["BE"] else None) if pct else r[col]
            if v is None or pd.isna(v) or (not pct and round(v) == 0):
                continue
            out.append(f"{GRADE_SHORT.get(grade.upper(), grade)} <b>{_pct(v) if pct else _mt(v)}</b>")
        return " · ".join(out)

    cards = [
        ("Latest BE (MT)", be, "mt", split("BE")),
        ("Expected Orders", total["Exp Orders"], "mt", split("Exp Orders")),
        ("Orders MTD", total["Orders MTD"], "mt", split("Orders MTD")),
        ("Invoiced MTD", total["Invoiced"], "mt", split("Invoiced")),
        ("Invoice % of BE", (total["Invoiced"] or 0) / be if be else None, "pct",
         split("Invoiced", pct=True)),
        ("Pending Orders to Serve", total["Pending to Serve"], "mt", split("Pending to Serve")),
        ("DO Released", total["DO Released"], "mt", split("DO Released")),
        ("PO Issued (H1+H2)", total["PO Issued"], "mt", split("PO Issued")),
        ("Production MTD", total["Production MTD"], "mt", split("Production MTD")),
        ("Expected BTR Comp", total["Exp BTR Comp"], "mt", split("Exp BTR Comp")),
        ("Physical Inventory", total["Physical Inv"], "mt", split("Physical Inv")),
        ("Expected Closing Inv", total["Exp Closing"], "mt", split("Exp Closing")),
    ]

    # Header line: total gap, then each grade with open orders — short (worst 2 plants) / covered
    parts = [f"Open-order serving gap <b>{_mt(total['Net to Serve'])} MT</b>"]
    open_g = g[g["Pending to Serve"].fillna(0) > 0]
    for grade, r in open_g.sort_values("Net to Serve").iterrows():     # shortfalls first
        net = r["Net to Serve"] or 0
        if net < 0:
            worst = p[(p["Grade"] == grade) & (p["Net to Serve"] < 0)].nsmallest(2, "Net to Serve")
            who = ", ".join(f"{html.escape(_short_plant(n))} {_signed(v)}"
                            for n, v in zip(worst["Plant"], worst["Net to Serve"]))
            parts.append(f"{html.escape(grade)} short {abs(net):,.0f} MT" + (f" ({who})" if who else ""))
        else:
            parts.append(f"{html.escape(grade)} covered +{net:,.0f} MT")
    return Report(as_on, meta, p, g, total, cards, " &nbsp;·&nbsp; ".join(parts), warnings)


# ─── Snapshot HTML (→ PNG) ───────────────────────────────────────────────────
_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{background:#E9EBEE;font-family:Inter,'Segoe UI',Arial,sans-serif;color:#1B2A3A;
 -webkit-font-smoothing:antialiased;font-variant-numeric:tabular-nums}
#card{width:1400px;margin:24px;background:#fff;box-shadow:0 8px 30px rgba(10,30,60,.12)}
.hd{background:#0E2A47;color:#fff;padding:34px 38px 28px;display:flex;
 justify-content:space-between;align-items:flex-end;border-bottom:3px solid #E07A2E;gap:24px}
.eyebrow{font-size:11.5px;letter-spacing:.14em;font-weight:600;color:#B8C6D6;white-space:nowrap}
.hd h1{font-size:30px;font-weight:700;margin:8px 0 6px;letter-spacing:-.01em;white-space:nowrap}
.hd .sub{font-size:13px;color:#C9D4E0;white-space:nowrap}
.hero{text-align:right}.hero .big{font-size:44px;font-weight:700;line-height:1.05;margin-top:6px}
.hero .unit{font-size:11.5px;letter-spacing:.1em;font-weight:700;color:#F0A35E;margin-top:4px}
.hero .note{font-size:12.5px;color:#DCE4EC;margin-top:10px}.hero .note b{color:#F0A35E}
.bd{padding:26px 38px 30px}
.sec{display:flex;align-items:center;gap:10px;margin:24px 0 14px}.sec.first{margin-top:4px}
.num{background:#0E2A47;color:#fff;font-size:11.5px;font-weight:700;padding:3px 7px;border-radius:3px}
.sec h2{font-size:15px;font-weight:700;white-space:nowrap}.sec .rule{flex:1;height:1px;background:#DDE2E8}
.grid{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}
.kpi{border:1px solid #DDE2E8;border-radius:4px;padding:13px 14px 12px;background:#FBFCFD}
.kpi .l{font-size:10.5px;letter-spacing:.09em;font-weight:600;color:#6B7A8C;text-transform:uppercase}
.kpi .v{font-size:25px;font-weight:700;margin:12px 0 9px;color:#0E2A47}
.kpi .s{font-size:10.5px;color:#6B7A8C;white-space:nowrap;min-height:13px}.kpi .s b{color:#0E2A47;font-weight:600}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#0E2A47;color:#fff;font-size:10.5px;letter-spacing:.06em;font-weight:600;
 text-transform:uppercase;padding:10px 8px;text-align:right;white-space:nowrap}
th.t,td.t{text-align:left}
td{padding:9px 8px;border-bottom:1px solid #E6EAEF;text-align:right;white-space:nowrap}
td.p{font-weight:700;border-right:1px solid #E6EAEF;white-space:normal;width:170px}
td.g{color:#5A6878}td.gb{font-weight:700}
td.iss{font-size:10px;line-height:1.35;color:#5A6878;white-space:normal;width:290px;text-align:left}
tr.tot td{background:#0E2A47;color:#fff;font-weight:700;font-size:13px;border-top:2px solid #E07A2E;
 border-bottom:0;padding:12px 8px;white-space:normal}
.pill{display:inline-block;min-width:40px;text-align:center;padding:2px 7px;border-radius:3px;font-weight:700}
.dash{color:#A7B1BD}
.ft{background:#0E2A47;color:#9FB0C3;font-size:10.5px;letter-spacing:.1em;padding:14px 38px;
 display:flex;justify-content:space-between;text-transform:uppercase}.ft b{color:#F0A35E;font-weight:600}
"""


def _c(v: float | None) -> str:
    t = _mt(v)
    return '<span class="dash">–</span>' if t == "–" else t


def _pill(v: float | None, which: str) -> str:
    col = pill_colours(v, which)
    if col is None:
        return _c(v)
    return f'<span class="pill" style="background:#{col[0]};color:#{col[1]}">{_mt(v)}</span>'


def issue_text(txt: str) -> str:
    """'FE 550 (8MM-46T, 10MM-10T)' -> 'FE 550 · 8MM (46T) · 10MM (10T)'; other text as-is."""
    if not re.search(r"\d+\s*MM\s*-", txt):
        return txt
    t = re.sub(r"(\d+\s*MM)\s*-\s*(\d+\s*T?)", lambda m: f"{m[1]} [{m[2]}]", txt)
    t = t.replace("(", " · ").replace(")", "").replace(",", " · ")
    t = re.sub(r"\s*·\s*(·\s*)*", " · ", t).strip(" ·")
    return t.replace("[", "(").replace("]", ")")


def build_snapshot_html(rep: Report) -> str:
    e = html.escape
    t = rep.total
    o = ['<!DOCTYPE html><html><head><meta charset="utf-8">'
         f"<title>PB TMT Update — {e(rep.day)}</title>"
         '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">'
         f"<style>{_CSS}</style></head><body><div id=\"card\">",
         '<div class="hd"><div><div class="eyebrow">JSW ONE &nbsp;·&nbsp; PRIVATE BRANDS — TMT</div>'
         f"<h1>PB TMT Update — {e(rep.day)}</h1>"
         '<div class="sub">Plant-wise &amp; Grade-wise MTD Snapshot &nbsp;·&nbsp; All Plants (Total Qty, MT)</div></div>'
         f'<div class="hero"><div class="eyebrow">AS ON {e(rep.day.upper())} &nbsp;·&nbsp; MONTH-TO-DATE</div>'
         f'<div class="big">{_mt(t["Invoiced"])}</div><div class="unit">MT INVOICED MTD</div>'
         f'<div class="note">{rep.gap_line}</div></div></div><div class="bd">']

    def sec(n: str, title: str, first: bool = False) -> None:
        o.append(f'<div class="sec{" first" if first else ""}"><span class="num">{n}</span>'
                 f'<h2>{e(title)}</h2><span class="rule"></span></div>')

    sec("01", "Key Metrics — MTD Snapshot (All Plants)", first=True)
    o.append('<div class="grid">' + "".join(
        f'<div class="kpi"><div class="l">{e(l)}</div>'
        f'<div class="v">{_pct(v) if k == "pct" else _mt(v)}</div><div class="s">{s}</div></div>'
        for l, v, k, s in rep.cards) + "</div>")

    sec("02", "Plant-wise Performance — by Grade")
    cols = ["BE", "Exp Orders", "Orders MTD", "Invoiced", "Pending to Serve", "Physical Inv",
            "Exp BTR Comp"]
    heads = ["Plant", "Grade", "BE", "Exp. Orders", "Orders MTD", "Invoiced", "Pending to Serve",
             "Physical Inv", "Exp. BTR Comp", "Net to Serve", "Exp. Closing", "Inventory Issue"]
    o.append("<table><tr>" + "".join(
        f'<th class="{"t" if h in ("Plant", "Grade", "Inventory Issue") else ""}">{h}</th>'
        for h in heads) + "</tr>")
    p = rep.plants
    spans = p.groupby("Plant", sort=False).size().to_dict()
    seen: set[str] = set()
    for _, r in p.iterrows():
        o.append("<tr>")
        if r["Plant"] not in seen:
            seen.add(r["Plant"])
            o.append(f'<td class="p t" rowspan="{spans[r["Plant"]]}">{e(r["Plant"])}</td>')
        o.append(f'<td class="g t">{e(r["Grade"])}</td>'
                 + "".join(f"<td>{_c(r[c])}</td>" for c in cols)
                 + f"<td>{_pill(r['Net to Serve'], 'Net to Serve')}</td><td>{_pill(r['Exp Closing'], 'Exp Closing')}</td>"
                 + f'<td class="iss">{e(issue_text(r["Inventory Issue"]))}</td></tr>')
    o.append('<tr class="tot"><td class="t" colspan="2">GRAND TOTAL — ALL PLANTS</td>'
             + "".join(f"<td>{_mt(t[c])}</td>" for c in cols)
             + f"<td>{_mt(t['Net to Serve'])}</td><td>{_mt(t['Exp Closing'])}</td><td></td></tr></table>")

    sec("03", "Grade-wise Summary — All Plants")
    gcols = ["BE", "Exp Orders", "Orders MTD", "Invoiced", "Pending to Serve", "PO Issued",
             "Production MTD", "Exp BTR Comp", "Physical Inv"]
    gh = ["Grade", "BE", "Exp. Orders", "Orders MTD", "Invoiced", "Pending to Serve", "PO Issued",
          "Production MTD", "Exp. BTR Comp", "Physical Inv", "Net to Serve", "Exp. Closing"]
    o.append("<table><tr>" + "".join(
        f'<th class="{"t" if i == 0 else ""}">{h}</th>' for i, h in enumerate(gh)) + "</tr>")
    for grade, r in rep.grades.iterrows():
        o.append(f'<tr><td class="t gb">{e(grade)}</td>'
                 + "".join(f"<td>{_c(r[c])}</td>" for c in gcols)
                 + f"<td>{_pill(r['Net to Serve'], 'Net to Serve')}</td><td>{_pill(r['Exp Closing'], 'Exp Closing')}</td></tr>")
    o.append('<tr class="tot"><td class="t">TOTAL — ALL GRADES</td>'
             + "".join(f"<td>{_mt(t[c])}</td>" for c in gcols)
             + f"<td>{_mt(t['Net to Serve'])}</td><td>{_mt(t['Exp Closing'])}</td></tr></table>")

    who = rep.meta.get("prepared_by") or ""
    o.append(f'</div><div class="ft"><span>PB TMT Update &nbsp;·&nbsp; {e(rep.day.upper())} MTD</span>'
             f"<span>{'PREPARED BY <b>' + e(who.upper()) + '</b> &nbsp;·&nbsp; ' if who else ''}"
             "JSW ONE PLATFORMS LTD.</span></div></div></body></html>")
    return "".join(o)


def render_png(page_html: str, png: Path) -> Path:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        try:
            b = pw.chromium.launch()
        except Exception:  # pinned Playwright without its bundled browser
            exe = next(Path("/opt/pw-browsers").glob("chromium-*/chrome-linux*/chrome"), None)
            if exe is None:
                raise
            b = pw.chromium.launch(executable_path=str(exe))
        page = b.new_page(viewport={"width": 1460, "height": 900}, device_scale_factor=2)
        try:
            page.set_content(page_html, wait_until="networkidle", timeout=15000)
        except Exception:  # offline: web font unavailable, system fallback font is used
            page.set_content(page_html, wait_until="load")
        page.evaluate("document.fonts.ready")
        page.locator("#card").screenshot(path=str(png))
        b.close()
    return png


# ─── Excel (same sections, values only) ──────────────────────────────────────
_thin = Side(style="thin", color="D0D6DE")
_B = Border(bottom=_thin)
_W = "FFFFFF"
MTF = '#,##0;(#,##0);"–"'


def _cell(ws, r: int, c: int, v: Any, *, fmt: str | None = None, bold: bool = False,
          bg: str | None = None, color: str = "1B2A3A", size: int = 10, left: bool = False,
          wrap: bool = False) -> None:
    x = ws.cell(row=r, column=c)
    x.value = None if v is None or (isinstance(v, float) and pd.isna(v)) else v
    x.font = Font(name="Calibri", size=size, bold=bold, color=color)
    x.alignment = Alignment(horizontal="left" if left else "right", vertical="center",
                            wrap_text=wrap)
    x.border = _B
    if bg:
        x.fill = PatternFill("solid", fgColor=bg)
    if fmt:
        x.number_format = fmt


def _plain(s: str) -> str:
    return html.unescape(re.sub(r"<[^>]+>", "", s))


def build_xlsx(rep: Report, out: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "PB TMT Update"
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([26, 11, 10, 11, 11, 11, 12, 11, 11, 11, 11, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.merge_cells("A1:L1")
    _cell(ws, 1, 1, f"PB TMT Update — {rep.day}   ·   Plant-wise & Grade-wise MTD Snapshot (MT)",
          bold=True, bg=NAVY, color=_W, size=14, left=True)
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:L2")
    _cell(ws, 2, 1, f"Invoiced MTD {_mt(rep.total['Invoiced'])} MT   ·   " + _plain(rep.gap_line),
          bg=NAVY, color="F0A35E", size=10, left=True)

    r = 4
    _cell(ws, r, 1, "01  Key Metrics — MTD Snapshot (All Plants)", bold=True, size=12, left=True)
    for i, (label, v, kind, split) in enumerate(rep.cards):
        row, col = r + 1 + (i // 6) * 3, 1 + (i % 6) * 2
        for k in range(3):
            ws.merge_cells(start_row=row + k, start_column=col, end_row=row + k, end_column=col + 1)
        _cell(ws, row, col, label.upper(), bold=True, color="6B7A8C", size=8, left=True, bg="F4F6F8")
        _cell(ws, row + 1, col, v, fmt="0%" if kind == "pct" else MTF, bold=True, size=16,
              color=NAVY, left=True, bg="F4F6F8")
        _cell(ws, row + 2, col, _plain(split) or None, color="6B7A8C", size=8, left=True, bg="F4F6F8")
    r += 8

    _cell(ws, r, 1, "02  Plant-wise Performance — by Grade", bold=True, size=12, left=True)
    heads = ["Plant", "Grade", "BE", "Exp. Orders", "Orders MTD", "Invoiced", "Pending to Serve",
             "Physical Inv", "Exp. BTR Comp", "Net to Serve", "Exp. Closing", "Inventory Issue"]
    keys = ["BE", "Exp Orders", "Orders MTD", "Invoiced", "Pending to Serve", "Physical Inv",
            "Exp BTR Comp", "Net to Serve", "Exp Closing"]
    for c, h in enumerate(heads, 1):
        _cell(ws, r + 1, c, h.upper(), bold=True, bg=NAVY, color=_W, size=9, left=c in (1, 2, 12))
    r0 = r + 2
    for i, (_, row) in enumerate(rep.plants.iterrows()):
        rr = r0 + i
        _cell(ws, rr, 1, row["Plant"], bold=True, left=True)
        _cell(ws, rr, 2, row["Grade"], color="5A6878", left=True)
        for c, k in enumerate(keys, 3):
            col = pill_colours(row[k], k) if k in ("Net to Serve", "Exp Closing") else None
            _cell(ws, rr, c, row[k], fmt=MTF, bold=col is not None,
                  bg=col[0] if col else None, color=col[1] if col else "1B2A3A")
        _cell(ws, rr, 12, issue_text(row["Inventory Issue"]) or None, color="5A6878", size=8,
              left=True, wrap=True)
    names = list(rep.plants["Plant"])
    start = 0
    for i in range(1, len(names) + 1):
        if i == len(names) or names[i] != names[start]:
            if i - start > 1:
                ws.merge_cells(start_row=r0 + start, start_column=1, end_row=r0 + i - 1, end_column=1)
            start = i
    rt = r0 + len(names)
    _cell(ws, rt, 1, "GRAND TOTAL — ALL PLANTS", bold=True, bg=NAVY, color=_W, left=True)
    _cell(ws, rt, 2, None, bg=NAVY)
    for c, k in enumerate(keys, 3):
        _cell(ws, rt, c, rep.total[k], fmt=MTF, bold=True, bg=NAVY, color=_W)
    _cell(ws, rt, 12, None, bg=NAVY)

    r = rt + 2
    _cell(ws, r, 1, "03  Grade-wise Summary — All Plants", bold=True, size=12, left=True)
    gk = ["BE", "Exp Orders", "Orders MTD", "Invoiced", "Pending to Serve", "PO Issued",
          "Production MTD", "Exp BTR Comp", "Physical Inv", "Net to Serve", "Exp Closing"]
    gh = ["Grade", "BE", "Exp. Orders", "Orders MTD", "Invoiced", "Pending to Serve", "PO Issued",
          "Production MTD", "Exp. BTR Comp", "Physical Inv", "Net to Serve", "Exp. Closing"]
    for c, h in enumerate(gh, 1):
        _cell(ws, r + 1, c, h.upper(), bold=True, bg=NAVY, color=_W, size=9, left=c == 1)
    for i, (grade, row) in enumerate(rep.grades.iterrows()):
        rr = r + 2 + i
        _cell(ws, rr, 1, grade, bold=True, left=True)
        for c, k in enumerate(gk, 2):
            col = pill_colours(row[k], k) if k in ("Net to Serve", "Exp Closing") else None
            _cell(ws, rr, c, row[k], fmt=MTF, bold=col is not None,
                  bg=col[0] if col else None, color=col[1] if col else "1B2A3A")
    rr = r + 2 + len(rep.grades)
    _cell(ws, rr, 1, "TOTAL — ALL GRADES", bold=True, bg=NAVY, color=_W, left=True)
    for c, k in enumerate(gk, 2):
        _cell(ws, rr, c, rep.total[k], fmt=MTF, bold=True, bg=NAVY, color=_W)
    who = rep.meta.get("prepared_by")
    ws.cell(row=rr + 2, column=1).value = (f"PB TMT Update · {rep.day} MTD"
                                           + (f"   ·   Prepared by {who}" if who else "")
                                           + "   ·   JSW One Platforms Ltd.")
    ws.cell(row=rr + 2, column=1).font = Font(name="Calibri", size=8, color="7A8796")
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


# ─── .eml draft: PNG inline + xlsx attached ──────────────────────────────────
def build_eml(rep: Report, xlsx: Path, png: Path | None, out: Path, *,
              to: str = "", cc: str = "", sender: str = "") -> Path:
    msg = EmailMessage()
    msg["Subject"] = f"PB TMT Update — {rep.day} (MTD)"
    if sender:
        msg["From"] = sender
    if to:
        msg["To"] = to
    if cc:
        msg["Cc"] = cc
    msg["Date"] = formatdate(localtime=True)
    msg["X-Unsent"] = "1"          # Outlook opens it as an editable draft
    msg.set_content(f"Dear All,\n\nPlease find the PB TMT update as on {rep.day} "
                    "(Excel attached).\n\nRegards,")
    font = "font:14px Calibri,Arial,sans-serif;"
    cid = make_msgid(domain="pb-tmt-update")
    img = (f'<img src="cid:{cid[1:-1]}" width="1000" alt="PB TMT Update {html.escape(rep.day)}" '
           'style="display:block;width:1000px;max-width:100%;height:auto;border:0;">') if png else ""
    sign = f"<br>{html.escape(rep.meta['prepared_by'])}" if rep.meta.get("prepared_by") else ""
    msg.add_alternative(
        f'<html><body><p style="{font}">Dear All,<br><br>Please find below the PB TMT update as on '
        f"<b>{html.escape(rep.day)}</b> (Excel attached).</p>{img}"
        f'<p style="{font}">Regards,{sign}</p></body></html>', subtype="html")
    if png:
        msg.get_payload()[1].add_related(png.read_bytes(), maintype="image", subtype="png",
                                         cid=cid, filename=png.name)
    msg.add_attachment(xlsx.read_bytes(), maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       filename=xlsx.name)
    out.write_bytes(bytes(msg))
    return out


# ─── Input template ──────────────────────────────────────────────────────────
def write_template(out: Path, *, meta: dict[str, Any] | None = None,
                   plants: list[list[Any]] | None = None) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Meta"
    ws.append(["key", "value", "description"])
    meta = meta or {}
    for k, desc in META_KEYS.items():
        ws.append([k, meta.get(k), desc])
    s = wb.create_sheet("Plant")
    s.append(PLANT_COLS)
    for r in plants or []:
        s.append(r)
    for sh in wb.worksheets:
        for c in sh[1]:
            c.font = Font(bold=True, color=_W)
            c.fill = PatternFill("solid", fgColor=NAVY)
        for i in range(1, sh.max_column + 1):
            sh.column_dimensions[get_column_letter(i)].width = 15
    wb["Meta"].column_dimensions["C"].width = 64
    wb["Plant"].column_dimensions["A"].width = 26
    wb["Plant"].column_dimensions["K"].width = 60
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def build_pack(inp: Path, out_dir: Path, *, to: str = "", cc: str = "", sender: str = "",
               png: bool = True) -> tuple[Report, dict[str, Path]]:
    rep = derive(*read_input(inp))
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"PB_TMT_Update_{rep.stamp}"
    files: dict[str, Path] = {}
    page = build_snapshot_html(rep)
    files["html"] = out_dir / f"{base}.html"
    files["html"].write_text(page, encoding="utf-8")
    files["xlsx"] = build_xlsx(rep, out_dir / f"{base}.xlsx")
    if png:
        try:
            files["png"] = render_png(page, out_dir / f"{base}.png")
        except Exception as exc:  # no browser available: still produce the rest
            rep.warnings.append(f"PNG not rendered ({exc.__class__.__name__}); mail has no image")
    files["eml"] = build_eml(rep, files["xlsx"], files.get("png"), out_dir / f"{base}.eml",
                             to=to, cc=cc, sender=sender)
    return rep, files


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    t = sub.add_parser("template", help="write a blank input workbook")
    t.add_argument("path", type=Path)
    b = sub.add_parser("build", help="build png + html + xlsx + eml from an input workbook")
    b.add_argument("input", type=Path)
    b.add_argument("--out", type=Path, default=Path(".workspace/daily_mail"))
    b.add_argument("--to", default="")
    b.add_argument("--cc", default="")
    b.add_argument("--sender", default="")
    b.add_argument("--no-png", action="store_true", help="skip the PNG snapshot")
    a = ap.parse_args(argv)
    if a.cmd == "template":
        print(write_template(a.path))
        return 0
    rep, files = build_pack(a.input, a.out, to=a.to, cc=a.cc, sender=a.sender, png=not a.no_png)
    for w in rep.warnings:
        print(f"WARNING: {w}", file=sys.stderr)
    for k, pth in files.items():
        print(f"{k}: {pth}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
