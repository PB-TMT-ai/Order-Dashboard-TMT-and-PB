"""Unit tests for scripts/pb_mtd_mail.py (run: python -m pytest tests/unit).

Golden case = the hand-built 25-Sep-2026 PB TMT Update: every derived number
below is what that report shows.
"""
from __future__ import annotations

import sys
from email import message_from_bytes
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))
import pb_mtd_mail as m  # noqa: E402

N = None
# Plant, Grade, BE, Exp Orders, Orders MTD, Invoiced, Pending, Physical, BTR, Exp Closing,
# Inventory Issue, PO Issued, Production MTD, DO Released
REF = [
    ["SKA Ispat", "FE 550", 450, 1000, 993, 547, 443, 228, N, 183, "FE 550 (8MM-46T, 10MM-10T)", N, N, N],
    [N, "FE 550D", N, N, N, 0, N, 7, N, 7, "", N, N, N],
    [N, "One Helix", 1400, 1400, 366, 368, N, 586, N, -448, "", N, N, N],
    ["API Ispat & Powertech", "FE 550", 6200, 6200, 4924, 3427, 1503, 3343, 1766, 2416, "", 21400, 10616, 843],
    ["Real Ispat", "FE 550D", 4900, 8700, 6392, 2823, 3544, 441, 1900, -3540, "", 21300, 12507, 1142],
    ["AIC Iron Industries", "FE 550", 700, 2200, 2159, 1305, 852, 345, 855, 1796, "", N, N, N],
    ["Ambashakti Udyog – Gwalior", "FE 550", N, N, N, 0, N, 37, N, 37, "", N, N, N],
    [N, "FE 550D", 9500, 12600, 12585, 7958, 4626, 446, 3048, -65, "", N, N, N],
    ["Ambashakti Industries", "FE 550", 6200, 6850, 6846, 5240, 1623, 1893, 2833, 3578, "", N, N, N],
    ["Aditya Industries", "FE 550", 1300, 1400, 1385, 733, 649, 1038, 346, 750, "", N, N, N],
    [N, "FE 550D", 1350, 3650, 3037, 2360, 685, 1045, 786, 449, "", N, N, N],
    [N, "One Helix", 1400, 1400, N, 0, N, 824, 500, -77, "", 900, 20, N],
    ["German Green Steel & Power", "FE 550D", 100, 150, 104, 92, 13, 216, N, 146, "", N, N, N],
    ["German Green Steel & Power", "FE 550", N, N, N, N, N, N, N, N, "", N, N, N],  # empty -> dropped
]


def _rep(tmp_path: Path, rows=REF, meta=None):
    inp = m.write_template(tmp_path / "in.xlsx", plants=rows,
                           meta=meta or {"as_on": "25-09-2026", "prepared_by": "Test"})
    return m.derive(*m.read_input(inp))


def test_golden_totals(tmp_path: Path) -> None:
    rep = _rep(tmp_path)
    t = rep.total
    assert (t["BE"], t["Exp Orders"], t["Orders MTD"], t["Invoiced"]) == (33500, 45550, 38791, 24853)
    assert (t["Pending to Serve"], t["Physical Inv"], t["Exp BTR Comp"]) == (13938, 10449, 12034)
    assert (t["Net to Serve"], t["Exp Closing"]) == (-3489, 5232)
    assert (t["PO Issued"], t["Production MTD"], t["DO Released"]) == (43600, 23143, 1985)
    assert len(rep.plants) == 13


def test_golden_grades_and_header(tmp_path: Path) -> None:
    rep = _rep(tmp_path)
    g = rep.grades
    assert list(g.index) == ["FE 550", "FE 550D", "One Helix"]
    assert g.loc["FE 550D", "Net to Serve"] == -6713 and g.loc["FE 550", "Exp Closing"] == 8760
    assert rep.gap_line.index("FE 550D short 6,713 MT (Gwalior -4,180, Real -3,103)") \
        < rep.gap_line.index("FE 550 covered +1,814 MT")
    assert "One Helix" not in rep.gap_line               # no open orders -> not listed
    cards = {label: (v, s) for label, v, _, s in rep.cards}
    assert round(cards["Invoice % of BE"][0], 2) == 0.74
    assert cards["Pending Orders to Serve"][1] == "550 <b>5,070</b> · 550D <b>8,868</b>"


def test_pending_computed_and_meta_fallback(tmp_path: Path) -> None:
    rows = [["P", "FE 550", 100, N, 80, 50, N, 20, N, N, "", N, N, N]]
    rep = _rep(tmp_path, rows, {"as_on": "01-10-2026", "exp_closing": 999})
    assert rep.plants.loc[0, "Pending to Serve"] == 30 and rep.plants.loc[0, "Net to Serve"] == -10
    assert rep.total["Exp Closing"] == 999
    assert any("Exp Orders" in w for w in rep.warnings)


def test_issue_text() -> None:
    assert m.issue_text("FE 550 (8MM-46T, 10MM-10T)") == "FE 550 · 8MM (46T) · 10MM (10T)"
    assert m.issue_text("FE 550 · 8MM (46T)") == "FE 550 · 8MM (46T)"


def test_pack_outputs(tmp_path: Path) -> None:
    inp = m.write_template(tmp_path / "in.xlsx", plants=REF, meta={"as_on": "25-09-2026"})
    rep, files = m.build_pack(inp, tmp_path / "out", png=False)
    assert {"html", "xlsx", "eml"} <= files.keys()
    msg = message_from_bytes(files["eml"].read_bytes())
    assert "25 Sep 2026" in msg["Subject"] and msg["X-Unsent"] == "1"
    assert [p.get_filename() for p in msg.walk() if p.get_filename()] == [files["xlsx"].name]
    assert "GRAND TOTAL — ALL PLANTS" in files["html"].read_text(encoding="utf-8")
