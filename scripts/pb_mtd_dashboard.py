#!/usr/bin/env python3
"""Render the PB MTD Dashboard one-pager (HTML + PNG) for a given day.

Wraps the `pb-mtd-dashboard` skill template so the dashboard can be produced
on a daily cadence with dated, non-clobbering output files.

Usage:
    python3 scripts/pb_mtd_dashboard.py <data.json> [--outdir DIR] [--date YYYY-MM-DD]
                                        [--template PATH]

Writes <outdir>/pb_mtd_dashboard_<date>.html and .png.

The Jinja template is the skill's `assets/template.html.jinja` (the approved
design). Override with --template or $PB_MTD_TEMPLATE. Never restyle here --
change the template in the skill so the design persists.
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import json
import os
from pathlib import Path

# Playwright ships a browser-build pin that can disagree with the image's
# pre-installed Chromium; fall back to whatever build is actually on disk.
_CHROMIUM_GLOBS = (
    "/opt/pw-browsers/chromium-*/chrome-linux/chrome",
    "/opt/pw-browsers/chromium_headless_shell-*/chrome-linux/headless_shell",
)


def find_template() -> Path:
    env = os.environ.get("PB_MTD_TEMPLATE")
    if env:
        return Path(env)
    home = Path.home() / ".claude" / "skills"
    hits = sorted(home.glob("**/pb-mtd-dashboard/assets/template.html.jinja"))
    if not hits:
        raise FileNotFoundError(
            "pb-mtd-dashboard template not found under ~/.claude/skills; "
            "pass --template or set $PB_MTD_TEMPLATE"
        )
    return hits[0]


def find_chromium() -> str | None:
    for pattern in _CHROMIUM_GLOBS:
        hits = sorted(glob.glob(pattern))
        if hits:
            return hits[-1]
    return None


def render(data_path: Path, out_base: Path, template_path: Path) -> tuple[Path, Path]:
    from jinja2 import Environment, FileSystemLoader

    data = json.loads(data_path.read_text())
    env = Environment(loader=FileSystemLoader(str(template_path.parent)))
    html = env.get_template(template_path.name).render(**data)

    out_html = out_base.with_suffix(".html")
    out_png = out_base.with_suffix(".png")
    out_html.parent.mkdir(parents=True, exist_ok=True)
    out_html.write_text(html)

    from playwright.sync_api import sync_playwright

    launch_kwargs = {}
    exe = find_chromium()
    if exe:
        launch_kwargs["executable_path"] = exe

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_kwargs)
        page = browser.new_page(
            viewport={"width": 1560, "height": 1000}, device_scale_factor=2
        )
        page.goto(f"file://{out_html.resolve()}")
        page.wait_for_timeout(350)
        page.screenshot(path=str(out_png), full_page=True)
        browser.close()

    return out_html, out_png


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("data", type=Path, help="dashboard data JSON (see data_schema.md)")
    ap.add_argument("--outdir", type=Path, default=Path(".workspace/pb_mtd"))
    ap.add_argument("--date", default=dt.date.today().isoformat(),
                    help="report date, YYYY-MM-DD (default: today)")
    ap.add_argument("--template", type=Path, default=None)
    args = ap.parse_args()

    template = args.template or find_template()
    out_base = args.outdir / f"pb_mtd_dashboard_{args.date}"
    out_html, out_png = render(args.data, out_base, template)
    print(f"Wrote {out_html}")
    print(f"Wrote {out_png}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
