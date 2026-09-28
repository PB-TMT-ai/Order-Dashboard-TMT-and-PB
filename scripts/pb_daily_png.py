"""PB MTD daily PNG — the "PB MTD DASHBOARD" Excel sheet, in the PB TMT Update design.

Uses ONLY the fields of the daily PB MTD Dashboard sheet (no Expected Orders etc.):
  header (Invoiced MTD hero) · 01 Key metrics (12 cards, grade splits) ·
  02 Plant-wise by grade · 03 Grade-wise summary.

Usage:
  python scripts/pb_daily_png.py template <input.xlsx>
  python scripts/pb_daily_png.py build <input.xlsx> [--out DIR]

Input workbook:
  Meta  : as_on, prepared_by, latest_be, dispatch_d1, btr, exp_closing, doh, ageing,
          prev_month_invoiced, production_mtd
  Plant : Plant, Grade, BE, Orders, Invoiced, Conf Pending Invoice, Pending Orders,
          Physical Inv, DOH, Ageing, Critical Dia       (Plant blank = row above)
  Grade : Grade, Production MTD, BTR, Exp Closing, DOH, Ageing
          (grade splits printed under the sheet's KPI tiles; optional)

Derived: all totals, Invoice % of BE, Pending to Serve (= Conf Pending + Pending Orders),
grade-wise summary. Plant rows are reconciled against Meta totals (warnings).
"""
from __future__ import annotations

import argparse
import html
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pb_mtd_mail import (  # noqa: E402  (shared design + helpers)
    _CSS, GRADE_SHORT, _grade_key, _mt, _num, _parse_date, _pct, render_png,
)

PLANT_COLS = ["Plant", "Grade", "BE", "Orders", "Invoiced", "Conf Pending Invoice",
              "Pending Orders", "Physical Inv", "DOH", "Ageing", "Critical Dia"]
QTY = ["BE", "Orders", "Invoiced", "Conf Pending Invoice", "Pending Orders", "Physical Inv"]
GRADE_COLS = ["Grade", "Production MTD", "BTR", "Exp Closing", "DOH", "Ageing"]
META_KEYS = {
    "as_on": "Report date (dd-mm-yyyy)",
    "prepared_by": "Prepared by (footer)",
    "latest_be": "Latest BE (MT)",
    "dispatch_d1": "Dispatch D-1 (MT)",
    "btr": "Balance to Produce - BTR (MT)",
    "exp_closing": "Expected Closing Inv (MT)",
    "doh": "Days of Inventory - DOH",
    "ageing": "Ageing (days)",
    "prev_month_invoiced": "Prev month invoiced MTD (MT)",
    "production_mtd": "Production MTD (MT)",
}
ALIASES = {"Conf. Pending Invoice": "Conf Pending Invoice",
           "Pending Orders (Non-Conf+SFDC)": "Pending Orders", "Ageing (Days)": "Ageing",
           "Physical inv": "Physical Inv", "OH": "One Helix"}
# Badge thresholds (repo blueprint pb-mtd-dashboard-daily): DOH ≤10 good · 11–20 warn · >20 bad;
# Ageing ≤15 good · 16–30 warn · >30 bad
GOOD, WARN, BAD = ("E3F4E8", "1E7B3A"), ("FDF1D6", "8A6100"), ("FBE1E1", "B42318")
DOH_BANDS = [(10, GOOD), (20, WARN), (float("inf"), BAD)]
AGEING_BANDS = [(15, GOOD), (30, WARN), (float("inf"), BAD)]


@dataclass
class Daily:
    as_on: date
    meta: dict[str, Any]
    plants: pd.DataFrame
    grades: pd.DataFrame
    total: dict[str, float | None]
    warnings: list[str] = field(default_factory=list)

    @property
    def day(self) -> str:
        return f"{self.as_on.day} {self.as_on.strftime('%b %Y')}"


def _sheet(wb, name: str, cols: list[str], required: list[str]) -> pd.DataFrame:
    rows = list(wb[name].iter_rows(values_only=True)) if name in wb.sheetnames else []
    if not rows:
        return pd.DataFrame(columns=cols)
    head = [ALIASES.get(str(c).strip(), str(c).strip()) if c is not None else "" for c in rows[0]]
    miss = [c for c in required if c not in head]
    if miss:
        raise ValueError(f"Sheet '{name}' missing columns: {miss}")
    df = pd.DataFrame({c: [r[head.index(c)] if c in head and head.index(c) < len(r) else None
                           for r in rows[1:]] for c in cols})
    return df[df.apply(lambda r: any(v not in (None, "") for v in r), axis=1)].reset_index(drop=True)


def load(path: Path) -> Daily:
    wb = load_workbook(path, data_only=True)
    raw = {str(r[0]).strip(): r[1] for r in wb["Meta"].iter_rows(min_row=2, values_only=True)
           if r and r[0] is not None}
    meta: dict[str, Any] = {k: _num(raw.get(k)) for k in META_KEYS if k not in ("as_on", "prepared_by")}
    meta["prepared_by"] = str(raw.get("prepared_by") or "").strip()
    as_on = _parse_date(raw.get("as_on"))
    warnings: list[str] = []

    p = _sheet(wb, "Plant", PLANT_COLS, ["Plant", "Grade", "BE", "Orders", "Invoiced", "Physical Inv"])
    p["Plant"] = p["Plant"].replace("", None).ffill().astype(str).str.strip()
    p["Grade"] = p["Grade"].fillna("").astype(str).str.strip().replace(ALIASES)
    p["Critical Dia"] = p["Critical Dia"].fillna("").astype(str).str.strip()
    nums = [c for c in PLANT_COLS if c not in ("Plant", "Grade", "Critical Dia")]
    for c in nums:
        p[c] = p[c].map(_num).astype(float)
    # Critical Dia is per plant (merged cell): keep it even if its row has no numbers
    dia = p.groupby("Plant", sort=False)["Critical Dia"].agg(lambda v: " ".join(x for x in v if x))
    p = p[p[nums].fillna(0).ne(0).any(axis=1)].reset_index(drop=True)
    p["Critical Dia"] = ""
    p.loc[~p["Plant"].duplicated(), "Critical Dia"] = p.loc[~p["Plant"].duplicated(), "Plant"].map(dia)
    p["Pending to Serve"] = p["Conf Pending Invoice"].fillna(0) + p["Pending Orders"].fillna(0)
    p.loc[p["Pending to Serve"] == 0, "Pending to Serve"] = None
    for i, r in p.iterrows():
        if pd.notna(r["Orders"]):
            gap = r["Orders"] - (r["Invoiced"] or 0) - (r["Pending to Serve"] or 0)
            if abs(gap) > 2:
                warnings.append(f"{r['Plant']} / {r['Grade']}: Orders ≠ Invoiced + Conf + Pending "
                                f"(off by {gap:+,.0f})")

    def s(col: pd.Series) -> float | None:
        return None if col.isna().all() else float(col.sum())

    cols = QTY + ["Pending to Serve"]
    g = p.groupby("Grade", sort=False)[cols].agg(s)
    gx = _sheet(wb, "Grade", GRADE_COLS, ["Grade"])
    gx["Grade"] = gx["Grade"].astype(str).str.strip().replace(ALIASES)
    for c in GRADE_COLS[1:]:
        gx[c] = gx[c].map(_num).astype(float)
    g = g.join(gx.set_index("Grade"), how="outer")
    g = g.loc[sorted(g.index, key=_grade_key)]
    g["Invoice %"] = (g["Invoiced"] / g["BE"]).where(g["BE"].fillna(0) > 0)
    g["Order %"] = (g["Orders"] / g["BE"]).where(g["BE"].fillna(0) > 0)

    total: dict[str, float | None] = {c: s(p[c]) for c in cols}
    for c, k in (("Production MTD", "production_mtd"), ("BTR", "btr"), ("Exp Closing", "exp_closing")):
        total[c] = meta.get(k) if meta.get(k) is not None else s(g[c])
        if meta.get(k) is not None and not g[c].isna().all() and abs(s(g[c]) - meta[k]) > 1:
            warnings.append(f"Grade {c} splits sum {s(g[c]):,.0f} ≠ Meta {meta[k]:,.0f}")
    be = meta.get("latest_be") or total["BE"]
    total["BE"] = be
    return Daily(as_on, meta, p, g, total, warnings)


def _c(v: float | None, dp: int = 0) -> str:
    if v is None or pd.isna(v):
        return '<span class="dash">–</span>'
    return f"{v:,.{dp}f}" if dp else _mt(v)


def _badge(v: float | None, bands: list[tuple[float, tuple[str, str]]], dp: bool = False) -> str:
    if v is None or pd.isna(v):
        return '<span class="dash">–</span>'
    bg, fg = next(c for lim, c in bands if v <= lim)
    return f'<span class="pill" style="background:#{bg};color:#{fg}">{_days(v) if dp else f"{v:,.0f}"}</span>'


def _days(v: float | None) -> str:
    if v is None or pd.isna(v):
        return "–"
    return f"{v:.1f}" if round(v, 1) != round(v) else f"{v:.0f}"


def _split(g: pd.DataFrame, col: str, fmt: str = "mt") -> str:
    out = []
    for grade, v in g[col].items():
        if v is None or pd.isna(v) or (fmt == "mt" and round(v) == 0):
            continue
        txt = _pct(v) if fmt == "pct" else (_days(v) if fmt == "days" else _mt(v))
        out.append(f"{GRADE_SHORT.get(str(grade).upper(), grade)} <b>{txt}</b>")
    return " · ".join(out)


def build_html(d: Daily) -> str:
    e = html.escape
    t, m, g = d.total, d.meta, d.grades
    be = t["BE"]
    prev = m.get("prev_month_invoiced")
    note = (f"Invoice % of BE <b>{_pct(t['Invoiced'] / be if be else None)}</b> &nbsp;·&nbsp; "
            f"Dispatch D-1 <b>{_mt(m.get('dispatch_d1'))} MT</b> &nbsp;·&nbsp; "
            f"Pending to Serve <b>{_mt(t['Pending to Serve'])} MT</b>"
            + (f" &nbsp;·&nbsp; Prev month invoiced (MTD) <b>{_mt(prev)} MT</b>" if prev else ""))
    cards = [
        ("Latest BE (MT)", _mt(be), _split(g, "BE")),
        ("Total Orders Received", _mt(t["Orders"]), _split(g, "Orders")),
        ("Invoiced MTD", _mt(t["Invoiced"]), _split(g, "Invoiced")),
        ("Invoice % of BE", _pct(t["Invoiced"] / be if be else None), _split(g, "Invoice %", "pct")),
        ("Pending Orders to Serve", _mt(t["Pending to Serve"]), _split(g, "Pending to Serve")),
        ("Dispatch D-1", _mt(m.get("dispatch_d1")),
         f"Prev month MTD inv <b>{_mt(prev)}</b>" if prev else ""),
        ("Physical Inventory", _mt(t["Physical Inv"]), _split(g, "Physical Inv")),
        ("Production MTD", _mt(t["Production MTD"]), _split(g, "Production MTD")),
        ("Balance to Produce (BTR)", _mt(t["BTR"]), _split(g, "BTR")),
        ("Expected Closing Inv", _mt(t["Exp Closing"]), _split(g, "Exp Closing")),
        ("Days of Inventory — DOH", _days(m.get("doh")), _split(g, "DOH", "days")),
        ("Ageing (Days)", _days(m.get("ageing")), _split(g, "Ageing", "days")),
    ]
    o = ['<!DOCTYPE html><html><head><meta charset="utf-8">'
         f"<title>PB MTD Dashboard — {e(d.day)}</title>"
         '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">'
         f"<style>{_CSS}td.dia{{font-size:10px;line-height:1.35;color:#5A6878;white-space:normal;"
         "width:300px;text-align:left;vertical-align:middle;border-left:1px solid #E6EAEF}</style>"
         '</head><body><div id="card">',
         '<div class="hd"><div><div class="eyebrow">JSW ONE &nbsp;·&nbsp; PRIVATE BRANDS — TMT</div>'
         f"<h1>PB MTD Dashboard — {e(d.day)}</h1>"
         f'<div class="sub">Plant-wise &amp; Grade-wise MTD Snapshot &nbsp;·&nbsp; '
         f'{e(d.as_on.strftime("%b-%y").upper())} &nbsp;·&nbsp; All Plants (MT)</div></div>'
         f'<div class="hero"><div class="eyebrow">AS ON {e(d.day.upper())} &nbsp;·&nbsp; MONTH-TO-DATE</div>'
         f'<div class="big">{_mt(t["Invoiced"])}</div><div class="unit">MT INVOICED MTD</div>'
         f'<div class="note">{note}</div></div></div><div class="bd">']

    def sec(n: str, title: str, first: bool = False) -> None:
        o.append(f'<div class="sec{" first" if first else ""}"><span class="num">{n}</span>'
                 f'<h2>{e(title)}</h2><span class="rule"></span></div>')

    sec("01", "Key Metrics — MTD Snapshot (All Plants)", first=True)
    o.append('<div class="grid">' + "".join(
        f'<div class="kpi"><div class="l">{e(l)}</div><div class="v">{v}</div>'
        f'<div class="s">{sp}</div></div>' for l, v, sp in cards) + "</div>")

    sec("02", "Plant-wise Performance (MT) — by Grade")
    heads = ["Plant", "Grade", "BE", "Orders", "Invoiced", "Conf. Pending Invoice",
             "Pending Orders (Non-Conf+SFDC)", "Physical Inv", "DOH", "Ageing (Days)", "Critical Dia"]
    o.append("<table><tr>" + "".join(
        f'<th class="{"t" if h in ("Plant", "Grade", "Critical Dia") else ""}">{h}</th>'
        for h in heads) + "</tr>")
    p = d.plants
    spans = p.groupby("Plant", sort=False).size().to_dict()
    dia = {k: "<br>".join(e(x) for x in v if x) for k, v in p.groupby("Plant", sort=False)["Critical Dia"]}
    seen: set[str] = set()
    for _, r in p.iterrows():
        first = r["Plant"] not in seen
        o.append("<tr>")
        if first:
            seen.add(r["Plant"])
            o.append(f'<td class="p t" rowspan="{spans[r["Plant"]]}">{e(r["Plant"])}</td>')
        o.append(f'<td class="g t">{e(r["Grade"])}</td>'
                 + "".join(f"<td>{_c(r[c])}</td>" for c in QTY)
                 + f"<td>{_badge(r['DOH'], DOH_BANDS)}</td><td>{_badge(r['Ageing'], AGEING_BANDS)}</td>"
                 + (f'<td class="dia" rowspan="{spans[r["Plant"]]}">{dia[r["Plant"]]}</td>' if first else "")
                 + "</tr>")
    o.append('<tr class="tot"><td class="t" colspan="2">GRAND TOTAL — ALL PLANTS</td>'
             + "".join(f"<td>{_mt(t[c])}</td>" for c in QTY)
             + f"<td>{_days(m.get('doh'))}</td><td>{_days(m.get('ageing'))}</td><td></td></tr></table>")

    sec("03", "Grade-wise Summary — All Plants")
    gh = ["Grade", "BE", "Orders", "Invoiced", "Invoice % of BE", "Conf. Pending Invoice",
          "Pending Orders", "Pending to Serve", "Physical Inv", "Production MTD", "BTR",
          "Exp. Closing", "DOH", "Ageing"]
    o.append("<table><tr>" + "".join(
        f'<th class="{"t" if i == 0 else ""}">{h}</th>' for i, h in enumerate(gh)) + "</tr>")
    for grade, r in g.iterrows():
        o.append(f'<tr><td class="t gb">{e(grade)}</td>'
                 + "".join(f"<td>{_c(r[c])}</td>" for c in ("BE", "Orders", "Invoiced"))
                 + f"<td>{_pct(r['Invoice %'])}</td>"
                 + "".join(f"<td>{_c(r[c])}</td>" for c in
                           ("Conf Pending Invoice", "Pending Orders", "Pending to Serve",
                            "Physical Inv", "Production MTD", "BTR", "Exp Closing"))
                 + f"<td>{_badge(r['DOH'], DOH_BANDS, dp=True)}</td>"
                 + f"<td>{_badge(r['Ageing'], AGEING_BANDS, dp=True)}</td></tr>")
    o.append('<tr class="tot"><td class="t">TOTAL — ALL GRADES</td>'
             + "".join(f"<td>{_mt(t[c])}</td>" for c in ("BE", "Orders", "Invoiced"))
             + f"<td>{_pct(t['Invoiced'] / be if be else None)}</td>"
             + "".join(f"<td>{_mt(t[c])}</td>" for c in
                       ("Conf Pending Invoice", "Pending Orders", "Pending to Serve",
                        "Physical Inv", "Production MTD", "BTR", "Exp Closing"))
             + f"<td>{_days(m.get('doh'))}</td><td>{_days(m.get('ageing'))}</td></tr></table>")

    who = m.get("prepared_by") or ""
    o.append(f'</div><div class="ft"><span>PB MTD Dashboard &nbsp;·&nbsp; {e(d.day.upper())} MTD</span>'
             f"<span>{'PREPARED BY <b>' + e(who.upper()) + '</b> &nbsp;·&nbsp; ' if who else ''}"
             "JSW ONE PLATFORMS LTD.</span></div></div></body></html>")
    return "".join(o)


def write_template(out: Path, *, meta: dict[str, Any] | None = None,
                   plants: list[list[Any]] | None = None,
                   grades: list[list[Any]] | None = None) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Meta"
    ws.append(["key", "value", "description"])
    for k, desc in META_KEYS.items():
        ws.append([k, (meta or {}).get(k), desc])
    for name, cols, rows in (("Plant", PLANT_COLS, plants), ("Grade", GRADE_COLS, grades)):
        sh = wb.create_sheet(name)
        sh.append(cols)
        for r in rows or []:
            sh.append(r)
    for sh in wb.worksheets:
        for c in sh[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0E2A47")
        for i in range(1, sh.max_column + 1):
            sh.column_dimensions[get_column_letter(i)].width = 16
    wb["Plant"].column_dimensions["K"].width = 70
    wb["Meta"].column_dimensions["C"].width = 40
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("template").add_argument("path", type=Path)
    b = sub.add_parser("build")
    b.add_argument("input", type=Path)
    b.add_argument("--out", type=Path, default=Path(".workspace/daily_mail"))
    a = ap.parse_args(argv)
    if a.cmd == "template":
        print(write_template(a.path))
        return 0
    d = load(a.input)
    base = a.out / f"PB_MTD_Dashboard_{d.as_on.strftime('%d-%b-%Y')}"
    a.out.mkdir(parents=True, exist_ok=True)
    page = build_html(d)
    base.with_suffix(".html").write_text(page, encoding="utf-8")
    render_png(page, base.with_suffix(".png"))
    for w in d.warnings:
        print(f"WARNING: {w}", file=sys.stderr)
    print(f"html: {base.with_suffix('.html')}\npng: {base.with_suffix('.png')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
