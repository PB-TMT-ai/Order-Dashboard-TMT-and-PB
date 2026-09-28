# Blueprint: PB MTD Dashboard — daily mail

## Goal
Produce the daily "PB MTD DASHBOARD — <MON-YY> (As on <date>)" mail: Outlook-ready
HTML body + formatted Excel attachment (+ optional PNG snapshot).

## Inputs Required
- input workbook (.xlsx) with sheets `Meta`, `Plant`, `Zone`.
  Create a blank one: `python scripts/pb_mtd_mail.py template PB_MTD_Input.xlsx`
  - `Meta`: as_on (dd-mm-yyyy), latest_be (blank = sum of plant BE), dispatch_d1,
    btr, expected_closing_inv, doh, ageing, prev_month_invoiced (MT)
  - `Plant`: one row per plant × grade. Leave Plant blank to repeat the row above
    (like merged cells). Pending Orders blank → Orders − Invoiced − Conf Pending.
  - `Zone`: North/Central/East/West × Retail/PTR rows + one `Project` row.
- optional: `--to`, `--cc`, `--sender` for the draft mail headers

## Scripts to Use
1. `scripts/pb_mtd_mail.py build <input.xlsx> --out <dir> [--png]`
   → `PB_MTD_Dashboard_<dd-Mon-yyyy>.xlsx`, `PB_MTD_Mail_<dd-Mon-yyyy>.html`,
   `PB_MTD_Mail_<dd-Mon-yyyy>.eml` (unsent draft with xlsx attached), optional `.png`
   (`--png`: designed snapshot — navy header, KPI cards with grade splits, plant table,
   grade-wise summary, zone table, footer; 2x. `Meta.prepared_by` fills the footer name).

## Steps
1. Update the input workbook with today's numbers (copy yesterday's, overwrite).
2. Run `build`; read every `WARNING:` line on stderr before sending.
3. Send: double-click the `.eml` → it opens in Outlook as a draft → Send.
   Or: open the `.html` in a browser, Ctrl+A, Ctrl+C, paste into a new mail, attach the `.xlsx`.

## Derived, never typed
KPI totals (from Plant sheet), Invoice % of BE, Production MTD (= Σ PO Prod),
PO Compliance (PO Prod / PO Issued), Order % vs BE, Retail+PTR subtotal, Grand totals.

## Edge Cases
- Zone rows don't sum to the plant grand total (> 5 MT on any column): an `Unmapped`
  row is added so the table still ties; a warning says by how much. Find the missing
  zone/channel in the source instead of leaving it.
- Pending Orders typed but ≠ Orders − Invoiced − Conf (> 2 MT): warning, typed value kept.
- Totals are sums of the typed rows: if the source rounds per row, totals can differ by
  a few MT from the source's own total. Paste unrounded values to avoid this.

## Colour rules (constants at top of script)
DOH ≤7 green · 8–20 amber · 21–30 orange · >30 red.
Ageing ≤10 green · 11–20 amber · 21–45 orange · >45 red.
Order % vs BE 90–110% green · 110–130% amber · >130% red · <90% no fill.
