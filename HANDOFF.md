# JSW Order Dashboard — Claude Code Handoff (v2, 2026-06-15)

**Purpose:** complete, self-contained context to continue the Streamlit JSW Order
Dashboard in a **new repo / new login**. Drop this file in the project root and
reference it at session start. It reflects the *current* shipped state (the
original port is complete; this supersedes the first HANDOFF).

**How to migrate:** copy the whole repo (all `.py` files + `.streamlit/`,
`requirements.txt`, `.env.example`) to the new account, recreate the secrets
(see §13), then read this file first.

---

## 0. TL;DR status

- The Streamlit port is **complete and in production shape**. All seven tabs are
  built and working: **Overview · Vs BE · Drill-down · Orders in Hand · Period
  compare · Scheme analysis · Line items**.
- Recent waves (all merged to `main`): Scheme **uplift** model, BE actuals
  **reconciliation** with the invoiced KPI, Overview **Form mix** donut.
- **One open bug** (not yet fixed): "Invoiced in period" reads **0** when the
  Invoice-sheet dates are day-first (`03-06-2026`) or Excel serials — see §8.1.
  Awaiting a sample file to fix safely.
- **Next wave not started:** W3 Secondary sales + opening stock — blocked on the
  daily secondary-sales data file (§9).

---

## 1. Project context

- **User:** JSW One, Private Brands business controls (TMT + Pipes & Tubes).
- **What:** distributor + order analytics dashboard for JSW One steel sales.
- **Why:** BE (Budget Estimate) review today; forecasting later (W5–W7).
- **History:** built first as a single-file HTML dashboard (Chart.js + XLSX.js +
  IndexedDB), then ported to **Streamlit + Plotly + Cloudflare R2** for
  multi-user deployment.
- **Deploy target:** Streamlit Community Cloud or any Python host; Cloudflare R2
  (S3-compatible) for cross-session file persistence (optional).

## 2. Repo layout (flat — all modules in project root)

```
app.py          # Streamlit entry: data load, sidebar filters, KPI strip, 7 tabs
data.py         # Order/Invoice parsing, enrich, channel/P&T/short-close, num(), filters, KPIs
be_logic.py     # BE Excel parsing, distributor clubbing, plan-vs-actual aggregation, BE table
plots.py        # Plotly figures (trend, rankings, mixes, India maps, BE trajectory)
drawer.py       # Universal right-side slide-over drill-down (st.dialog based)
theme.py        # JSW palette + Plotly template + CSS (pills, KPI cards)
storage_io.py   # Cloudflare R2 upload/download (boto3 / S3 API)
.streamlit/      # config.toml (theme)
requirements.txt # streamlit, pandas, numpy, openpyxl, python-calamine, plotly, boto3, python-dotenv
.env.example     # R2 config template
tests/           # placeholder dirs only (no real tests; validate via AppTest + synthetic dfs)
```

Tech stack: **Python 3.11**, Streamlit ≥1.40, pandas ≥2.0, Plotly ≥5.20,
boto3. No external services required to run (R2 is optional).

## 3. Data files & schemas

### 3.1 Order workbook (primary upload, `.xlsx`)
Single workbook with an **Order** sheet (any sheet whose name contains "order")
and an **Invoice** sheet (exact `Invoice`, else any name containing "invoice").
~59k order rows in production. Source column → internal field map lives in
`data.K` (exact header strings). Key columns: `Order ID`(`_oid`),
`Opportunity date`(`_d`), `Distributor Name`(`_dn`/`_dnN`), `Order Type`(`_ot`),
`Distributor (Yes/No)`(`_dis`), `Ship to State`(`_st`), `Grade`(`_gr`),
`Diameter mm`(`_dia`), `Form`(`_fm`), `Quantity`(`_q`), `Release Qty`(`_rq`),
`Invoiced Qty`(`_iq`), `total cancelled qty`(`_cq`), `Order Status`(`_sta`),
`Payment Terms`(`_p2`), `Initial delivery Mode`(`_dl`), `CM name`(`_cm`),
`Bill to - GST`(→`_bs` via first 2 GST digits → `data.SC_`).

### 3.2 Invoice sheet (same workbook)
4 columns: `Order ID`, `Invoice date`, `Invoiced qty`, `Invoice number`. One
Order ID may have many invoice rows (partial invoicing). Indexed by
`build_invoice_index` → `{oid: InvoiceEntry(total_qty, invs[{date,qty,num}],
first_date, last_date)}`. **⚠ Invoice-date parsing is the current bug — §8.1.**

### 3.3 BE workbook (separate upload)
Sheet `Distributor BE W3` (fallback regex `/distributor.*be|be.*week|
distributor.*w\d/i`). Row = distributor × state; columns grouped under
`Retail & PTR` and `Distributor through Project` for `Fe 550` / `Fe 550D`.
Parsed by `be_logic.parse_be_sheet` (carries `lastGroup` forward; skips
Total/Grand Total rows). **Comma-thousands fix is critical** — see `num()`.
Weekly versions stored in R2 at `be_versions/<month>_<week>.xlsx` (W1–W4).

## 4. Business logic invariants (LAW — confirm with user before changing)

- **Channels (4)** — `order_channel(ot, dis)`: `ss` if Self-stocking; `pd` if
  Project & dis=="Yes"; `pdir` if Project (else); `rt` (Retail) absorbs blanks
  and everything else. Sub-lines MUST sum to the headline. Do **not** add an
  "Other" bucket. Codes `rt`/`ss`/`pdir`/`pd` are used as dict keys & CSS — do
  not rename.
- **P&T vs TMT** — `is_pt`: CM contains "nippon" OR grade contains/starts "yst".
- **BE-eligible grade** — `order_grade`: `Fe 550` / `Fe 550D` only (lowercased,
  spaces stripped); else None.
- **Short-close** — pending `< 5 MT` → 0; order `> 60 days` old → 0. Applied
  uniformly (Orders-in-Hand, BE pipeline, drill tables).
- **Bill-to state** — first 2 digits of GST → `data.SC_`.
- **`num()`** — strips thousand commas / currency / NBSP; `"1,000"` → 1000.0
  (NOT 1). Use everywhere; never `float()` raw cells.
- **Invoice-date attribution** — `invoiced_in_range` / `invoiced_in_period`:
  an order's invoiced qty is attributed to the month(s) its **invoices** fall
  in, proportionally across the order's SKUs. April order invoiced in May counts
  for **May**.

## 5. BE matching — distributor-only (clubbed)

`flatten_be_atomic`: one atomic per distributor = sum across states/grades/
categories (keeps primary state for the map). Pal Cement (Punjab 500 + Haryana
1300 + HP 200) = one 2000 MT atomic; any eligible Pal order anywhere counts.
Actuals eligibility (`be_actuals_agg`): `_pt=="TMT"` AND `order_grade ∈
{Fe 550, Fe 550D}` AND `_ch ∈ {rt, ss, pd}` (pdir excluded) AND invoice-in-month
> 0 OR open pipeline > 0. Drill levels: region → distributor only.

## 6. Current state — what works (all merged)

- Sidebar: order + BE upload (R2 persist + auto-hydrate); product type;
  distributor flag; period (Last-X / Month / Year / Range); **13 multi-selects**
  (`data.MULTI_FILTERS`: order type, distributor, ship-to state/city, bill-to
  state, grade, diameter, form, payment terms, delivery, CM/plant, ship type,
  order status).
- KPI strip: Ordered / Released / Invoiced / Invoiced-in-period / BE gap, each
  with 4 channel sub-lines.
- **Overview:** channel trend, India choropleth + Top-10 sidekick (Ship/Bill
  toggle), Top-10 distributors, **Form mix** (Straight / U-bend / Fishbend) +
  Grade mix donuts. Per-card CSV; click-to-drill drawer everywhere.
- **Vs BE:** 6 KPI tiles incl. **Total eligible actuals**, daily invoice
  trajectory (clickable), BE-vs-Actuals table (search + state filter, row
  drill), state gap choropleth (abs/% toggle), unmatched-BE table, BE-A-vs-BE-B
  comparison; KPI drill buttons.
- **Drill-down:** up-to-4-level dimension picker → sunburst + pivot (row drill).
- **Orders in Hand:** pending KPIs + pending-by-distributor table.
- **Period compare:** two windows, A-vs-B KPIs, per-channel table, overlay trend.
- **Scheme analysis:** **Before/During/After uplift** (equal-length windows;
  Ordered/Invoiced metric; Uplift% & Sustained% per distributor; row drill).
- **Line items:** searchable table + CSV.

## 7. Recent changes (this engagement) — implementation notes

1. **Scheme uplift model** (`app.py tab_sc`): scheme start/end pickers; Before =
   `[start−L, start−1]`, After = `[end+1, end+L]`, `L = (end−start)+1` days;
   metric toggle Ordered (by `_d`) / Invoiced (`invoiced_in_period`); per-dist
   table Before/During/After + Uplift%/Sustained% (blank when Before=0).
2. **BE reconciliation (8.1 decision: include zero-BE distributors, channels
   rt+ss+pd, within BE month):** `be_actuals_agg` now also accumulates eligible
   actuals from distributors with **no BE** → `nobe_act`/`nobe_pipe`/
   `nobe_orders_by_key`; added `BeAggregate.total_act`/`total_pipe`; `be_table`
   no-BE branch uses the same proportional attribution (via
   `nobe_orders_by_key`), so **table Actuals == total_act == invoiced-in-BE-month
   for the eligible scope**. New "Total eligible actuals" KPI tile + drill.
   Gap-vs-BE tiles still use `matched_act` only (plan covers BE distributors).
   Fixed a latent crash: NaN ship-to state (`nan or ""` is truthy) → use
   `data.cl()`.
3. **Form mix** (`plots.form_label` + `plots.form_mix`): replaced the Overview
   "Order status mix" donut. `form_label` normalises Form cells to Straight /
   U-bend / Fishbend (case/space variants; unknowns kept as-is; blank →
   "Unspecified"). The card CSV groups by the same label.

## 8. Open issues / pending work

### 8.1 ⚠ BUG: "Invoiced in period" = 0 (HIGH — next to fix)
**Symptom:** with a Month filter, "Invoiced" shows e.g. 1,144 (from Order sheet
`_iq`) but "Invoiced in period" shows **0**.
**Root cause:** `data._date_series` parses Invoice dates with
`format="mixed"` (month-first default). Reproduced:
- `03-06-2026` / `03/06/2026` (Indian **day-first**, day ≤ 12) → read as **6
  March** → outside the month → 0.
- **Excel serial** numbers (`46176`) → **NaT** → 0.
Because today is early in the month, all day-1..12 invoices get thrown to other
months → the whole month reads 0.
**Fix plan (awaiting a sample file to confirm format):** make `_date_series`
(a) detect Excel serial numbers (numeric → `to_datetime(origin='1899-12-30',
unit='D')`), and (b) parse day-first dates correctly **without** breaking the
Order sheet, which the original HANDOFF says may be US-style (`5/8/26` = May 8).
Both sheets share this one parser, so confirm each sheet's real format before
flipping `dayfirst`. Validate against the user's samples + the synthetic
reproduction (see git history of this session).

### 8.2 Other pending (from original handoff, lower priority)
- State-name "unmapped" footer when a data state has no GeoJSON match.
- Manual "Refresh from storage" button (auto-hydrate on load already exists).
- BE weekly version selector is a selectbox (not the range slider once imagined).

## 9. Forward roadmap (locked waves, build in order)

- **W3 — Secondary sales + opening stock** *(next; BLOCKED on data file)*. Load a
  daily secondary-sales file → stock-in-hand per distributor = opening +
  cumulative primary − cumulative secondary; sell-through view. **Need from
  user:** the secondary-sales file (schema/sample) + opening-stock source. Not
  found in SharePoint search; user said "we will do later".
- **W4 — Price + scheme log** (JSW MoM prices, Steelmint NCR daily, prospective
  scheme logging). Makes Scheme analysis forward-looking, not just retroactive.
- **W5 — Distributor baselines** (24-mo history per distributor × order-type:
  seasonality, pacing, FY-end behaviour).
- **W6 — Forecast model** (Baseline × Pacing × Stock × Price × Scheme →
  per-distributor × order-type forecasts; rest-of-month/next-month/quarter/FY).
- **W7 — Override + scenario builder** (manual overrides + "what-if" price/scheme
  scenarios).

## 10. Code conventions

- Flat imports (`from data import num`), snake_case, `_`-prefixed derived
  DataFrame columns (`_q`, `_d`, `_ch`, …). Channel codes `rt/ss/pdir/pd` fixed.
- `@st.cache_data` for parse functions; parsed data + BE live in
  `st.session_state` (`order_data=(df, inv_index)`, `be_version`).
- Display: `data.fmt()` Indian-grouped integers, no decimals for MT.
- Drawer: `drawer.open_drawer(title, df, subtitle=…, summary=[(k,v)…],
  filename=…)`; `drawer.render()` is called once at the end of `app.py`.
- Don'ts: no "Other" channel bucket; don't drop blank order types (→ Retail);
  don't `float()` raw cells (use `num()`); don't drop proportional invoice
  allocation; BE state/grade/category are clubbed (don't re-split).

## 11. Validation (no real test suite)

Validate changes with: `python3 -m py_compile`, synthetic-DataFrame unit checks
of the touched logic, and a headless render via
`from streamlit.testing.v1 import AppTest` — inject
`at.session_state["order_data"]=(df, inv_index)` and `["be_version"]=BeVersion(…)`
before `at.run()`, then assert `at.exception is None` and that expected
markdown/labels are present. Reconciliation invariant: BE-tab total Actuals ==
`ag.total_act` == `invoiced_in_period(eligible, be_month)`.

## 12. Where to look in code

| Change | File | Symbol |
|---|---|---|
| Channel / P&T / short-close | data.py | `order_channel`, `is_pt`, `enrich` |
| Comma / number parsing | data.py | `num` |
| **Date parsing (bug 8.1)** | data.py | `_date_series`, `build_invoice_index` |
| Invoice-date attribution | data.py | `invoiced_in_range`, `invoiced_in_period` |
| Filters / KPIs | data.py | `apply_filters`, `MULTI_FILTERS`, `compute_kpis` |
| BE parse / clubbing | be_logic.py | `parse_be_sheet`, `flatten_be_atomic` |
| BE aggregation + reconciliation | be_logic.py | `be_actuals_agg`, `BeAggregate.total_act`, `be_table` |
| BE table / state gap / compare | be_logic.py | `be_table`, `be_state_gap`, `be_compare_table` |
| Charts / mixes / maps | plots.py | `channel_trend`, `form_mix`/`form_label`, `india_map`, `be_trajectory` |
| KPI strip | app.py | ~L342–381 |
| Tabs | app.py | `st.tabs([...])` (~L548) |
| Scheme uplift | app.py | `with tab_sc:` |
| R2 persistence | storage_io.py | `upload_*`, `download_*`, `is_configured` |

## 13. Migration checklist (new login / new repo)

1. Copy all files in §2 to the new repo.
2. Secrets / env (or `.streamlit/secrets.toml`): `R2_ENDPOINT` (or
   `R2_ACCOUNT_ID`), `R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY`, `R2_BUCKET`.
   Persistence is optional — the app runs without R2 (no cross-session sharing).
3. `pip install -r requirements.txt` → `streamlit run app.py`.
4. (If carrying data) copy the R2 bucket contents, incl. `be_versions/`.
5. Read this file, then fix §8.1 first.

---
_End of handoff. One feature at a time; validate against §11 before shipping._
