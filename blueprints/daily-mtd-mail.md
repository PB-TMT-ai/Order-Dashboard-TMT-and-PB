# Blueprint: PB TMT Update — daily mail

## Goal
Produce the daily "PB TMT Update — <d Mon yyyy>" mail: designed PNG snapshot inline in an
Outlook draft, with the same numbers attached as Excel.

## Inputs Required
- input workbook (.xlsx), sheets `Meta` + `Plant`.
  Blank one: `python scripts/pb_mtd_mail.py template PB_TMT_Update_Input.xlsx`
  - `Meta`: as_on (dd-mm-yyyy), prepared_by, latest_be (blank = Σ plant BE); fallbacks used
    only when the Plant column is blank: do_released, exp_btr_comp, exp_closing
  - `Grade` (optional): per-grade Exp Orders / Exp BTR Comp / Exp Closing / DO Released —
    fills cards + grade summary when only grade-level figures exist (e.g. BTR note
    "FE550/FE550D/OH: 5,360/3,814/500")
  - `Plant` (one row per plant × grade; blank Plant = row above):
    Plant, Grade, BE, Exp Orders, Orders MTD, Invoiced, Pending to Serve, Physical Inv,
    Exp BTR Comp, Exp Closing, Inventory Issue, PO Issued, Production MTD, DO Released
- optional: `--to`, `--cc`, `--sender` for the draft headers

## Scripts to Use
1. `scripts/pb_mtd_mail.py build <input.xlsx> --out <dir>` →
   `PB_TMT_Update_<dd-Mon-yyyy>.png|.html|.xlsx|.eml`

## Steps
1. Copy yesterday's input, overwrite the numbers.
2. Run `build`; read every `WARNING:` line before sending.
3. Double-click the `.eml` → Outlook draft (PNG in body, xlsx attached) → add recipients → Send.

## Derived, never typed
- Pending to Serve (if blank) = max(Orders MTD − Invoiced, 0)
- Net to Serve = Physical Inv − Pending to Serve (verified on every row of the 25-Sep report)
- Invoice % of BE, all totals, per-grade card splits, grade-wise summary
- Header line: total serving gap, then each grade with open orders — "short X MT (worst 2
  plants)" first, then "covered +X MT"

## Not derivable — must be typed
Exp Orders, Exp BTR Comp, Exp Closing (not a fixed formula: most rows ≈ Physical + BTR −
(Exp Orders − Invoiced), but AIC / Gwalior / Ambashakti / SKA 550 differ), DO Released,
Inventory Issue. Blank → shown as "–".

## Colour rules (constants at top of script)
Net to Serve / Exp Closing pills: < 0 red · amber below green-from · green from 1,000 MT
(Net to Serve) / 750 MT (Exp Closing) — matches the 25-Sep reference.

## Edge Cases
- Inventory Issue accepts "FE 550 (8MM-46T, 10MM-10T)" or "FE 550 · 8MM (46T)" — normalised.
- Rows with no numbers (e.g. an unused grade) are dropped.
- No browser for PNG → warning; xlsx/html/eml still built (eml without image).

## Golden test
`tests/unit/test_pb_mtd_mail.py` rebuilds the 25-Sep report and asserts every total.

## Mapping from the "PB MTD DASHBOARD" Excel sheet
- Pending to Serve = Conf. Pending Invoice + Pending Orders (Non-Conf+SFDC) (= Orders − Invoiced)
- DO Released = Conf. Pending Invoice (assumption — confirm with the business)
- Exp BTR Comp = "Balance to Produce (BTR)"; Exp Closing = "Expected Closing Inv" (totals / grade notes)
- Inventory Issue = "Critical Dia"
- Plant names: Amba-Sikandrabad → Ambashakti Industries; Amba Gwalior → Ambashakti Udyog – Gwalior;
  Aditya → Aditya Industries; API → API Ispat & Powertech; SKA → SKA Ispat;
  German Steel → German Green Steel & Power; AIC → AIC Iron Industries
