"""Unit tests for scripts/pb_daily_png.py (daily sheet -> PNG, no Expected Orders)."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))
import pb_daily_png as d  # noqa: E402

N = None
PLANTS = [
    ["A", "FE 550", 100, 90, 60, 10, 20, 50, 11, 12, "FE 550 (8MM-5T)"],
    ["B", "FE 550", N, N, N, N, N, N, N, N, "FE 550D (12MM-9T)"],   # empty row, dia kept
    [N, "FE 550D", 50, 40, 30, N, 10, 20, 5, 7, ""],
]


def test_totals_and_rules(tmp_path: Path) -> None:
    inp = d.write_template(tmp_path / "i.xlsx", plants=PLANTS,
                           meta={"as_on": "28-09-2026", "btr": 30},
                           grades=[["FE 550", 5, 20, N, N, N], ["FE 550D", 3, 10, N, N, N]])
    r = d.load(inp)
    assert (r.total["Orders"], r.total["Invoiced"], r.total["Pending to Serve"]) == (130, 90, 40)
    assert r.total["BTR"] == 30 and r.total["Production MTD"] == 8
    assert len(r.plants) == 2 and r.plants.loc[1, "Critical Dia"] == "FE 550D (12MM-9T)"
    html = d.build_html(r)
    assert "Expected Orders" not in html and "Exp. Orders" not in html
    assert not r.warnings
