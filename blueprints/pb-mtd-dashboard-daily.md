# Blueprint: PB MTD Dashboard — daily run

## Goal
Produce the branded PB (Private Brands — TMT) MTD Dashboard one-pager
(HTML + PNG) once per working day, from that morning's MTD numbers.

## Inputs Required
- source: the day's PB MTD Excel / dispatch MIS (or a pasted screenshot of it)
- report_date: YYYY-MM-DD — the "as on" date
- day_of_month / days_in_month: for the `AS ON … · DAY n/N` line

## Scripts to Use
1. `scripts/pb_mtd_dashboard.py` — fills the approved Jinja template and
   screenshots it to PNG. Writes dated files, so a daily run never clobbers
   yesterday's.

## Steps
1. Transcribe the day's numbers into a data JSON.
   - Shape: `pb-mtd-dashboard` skill → `references/data_schema.md`.
   - Start from the previous day's JSON in `.workspace/pb_mtd/` and edit
     values — the plant/grade and zone/type structure is stable month to
     month, only the numbers move.
2. **Reconcile before rendering** (see Edge Cases — this is the step that
   catches transcription errors):
   - plant rows must sum to the GRAND TOTAL row for BE, Orders, Invoiced,
     Conf. Pending, Pending Orders and Physical Inv;
   - zone rows must sum to `Retail+PTR`, and `Retail+PTR` + `Project` must
     equal the zone GRAND TOTAL.
   Fix any column that does not tie before rendering.
3. Render:
   ```bash
   python3 scripts/pb_mtd_dashboard.py .workspace/pb_mtd/pb_mtd_<date>.json \
       --date <YYYY-MM-DD>
   ```
4. Eyeball the PNG — confirm numbers landed in the right cells and that a
   long Critical Dia note has not overflowed its column.
5. Send the PNG (primary deliverable) and the HTML (so text can be tweaked).

## Update the header for each day
`as_of` (`AS ON 10 SEP 2026 · DAY 10/30`), `subtitle` (month), `hero_value`
(Invoiced MTD), `mom_note` and `footer_left` all carry the date — bump them
all, not just `as_of`.

## Edge Cases
- **A column does not tie to the grand total**: the transcription is wrong,
  not the sheet. Diff against the previous day's JSON — a single cell is
  usually the culprit.
- **Collapsed/hidden rows in the source Excel**: a screenshot of a sheet with
  grouped rows hides grade rows that the GRAND TOTAL still includes. Work
  from the file, not the screenshot, when a column will not tie.
- **No activity for a plant-grade or zone-type**: use the literal string
  `"–"` (en dash), matching the source Excel.
- **Zone with zero orders**: leave `order_pct` as `""` to render a plain `–`
  instead of a badge.
- **Badge thresholds**: DOH ≤10 good / 11–20 warn / >20 bad · Ageing ≤15 good
  / 16–30 warn / >30 bad · Order % 85–110 good, else warn/bad by direction.

## Known Issues
- **Playwright browser mismatch**: the image ships Chromium build 1194 but a
  freshly pip-installed Playwright pins a newer build and errors with
  `Executable doesn't exist at /opt/pw-browsers/chromium_headless_shell-<n>`.
  `scripts/pb_mtd_dashboard.py` resolves the on-disk binary itself, so use it
  rather than the skill's own `render.py`. Do **not** run `playwright install`.
- **Design changes** belong in the skill's `assets/template.html.jinja`, not
  in this repo — that is what keeps the look stable across runs.
