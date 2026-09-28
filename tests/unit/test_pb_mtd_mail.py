"""Unit tests for scripts/pb_mtd_mail.py (run: python -m pytest tests/unit)."""
from __future__ import annotations

import sys
from email import message_from_bytes
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))
import pb_mtd_mail as m  # noqa: E402

PLANTS = [
    ["Plant A", "FE 550", 1000, 1200, 800, 100, None, 300, 5, 12, "12MM-50", 1000, 700],
    [None, "FE 550D", 500, 400, 300, 50, 60, 100, 40, 50, "", None, None],
    ["Plant B", "FE 550", 700, 600, 500, None, 100, 200, 9, 8, "", 500, 250],
]
ZONES = [
    ["North", "Retail", 1000, 1000, 700, 100, 200],
    [None, "PTR", 500, 700, 500, 50, 50],
    ["Project", "", 700, 500, 400, None, 210],
]
META = {"dispatch_d1": 50, "btr": 100, "expected_closing_inv": 400, "doh": 7,
        "ageing": 20.5, "prev_month_invoiced": 1500}


def _build(tmp_path: Path, zones=ZONES):
    inp = m.write_template(tmp_path / "in.xlsx", as_on="28-09-2026", meta=META,
                           plants=PLANTS, zones=zones)
    return m.build_pack(inp, tmp_path / "out")


def test_totals_and_derivations(tmp_path: Path) -> None:
    rep, _ = _build(tmp_path)
    t = rep.plant_total
    assert (t["BE"], t["Orders"], t["Invoiced"]) == (2200, 2200, 1600)
    assert t["Pending Orders"] == 300 + 60 + 100          # blank pending computed
    assert t["PO Compliance"] == 950 / 1500
    assert rep.plants["Plant"].tolist() == ["Plant A", "Plant A", "Plant B"]
    kpi = dict((k, v) for k, v, _ in rep.kpi)
    assert kpi["Invoice % of BE"] == "73%" and kpi["Production MTD (MT)"] == "950"
    # pending 60 given but 400-300-50 = 50 -> flagged
    assert any("Pending Orders 60" in w for w in rep.warnings)


def test_zone_ties_without_unmapped_when_consistent(tmp_path: Path) -> None:
    rep, _ = _build(tmp_path)
    z = rep.zones
    assert "Unmapped" not in z["Zone"].tolist()
    assert z.iloc[-1]["Orders"] == rep.plant_total["Orders"]
    assert z.loc[z["Zone"] == "Retail+PTR", "Orders"].item() == 1700


def test_unmapped_row_when_zones_short(tmp_path: Path) -> None:
    rep, _ = _build(tmp_path, zones=ZONES[:2])
    un = rep.zones[rep.zones["Zone"] == "Unmapped"].iloc[0]
    assert un["Orders"] == 500 and un["BE"] == 700


def test_outputs_written(tmp_path: Path) -> None:
    _, files = _build(tmp_path)
    assert {"xlsx", "html", "eml"} <= files.keys()
    msg = message_from_bytes(files["eml"].read_bytes())
    assert "28-Sep-2026" in msg["Subject"] and msg["X-Unsent"] == "1"
    names = [p.get_filename() for p in msg.walk() if p.get_filename()]
    assert names == [files["xlsx"].name]
    assert "GRAND TOTAL" in files["html"].read_text(encoding="utf-8")
