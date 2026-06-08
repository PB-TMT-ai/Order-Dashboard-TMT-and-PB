# JSW Order Dashboard — Streamlit

Python/Streamlit port of the HTML dashboard. Same business logic, Cloudflare R2 persistence.

## What's identical to the HTML version
- Channel classification (Retail / Self-stocking / Project-direct / Project-thru-Dist)
- P&T vs TMT detection (Nippon / "yst*" prefix)
- Short-close rules (pending <5 MT or >60 days → zero)
- BE parsing including the comma-formatted thousand fix
- Invoice-date attribution with proportional allocation
- **BE matching = distributor-only** (clubbed across state/grade/category) — Pal Cement → Goa counts as Pal's actuals
- BE eligibility: TMT grade + channels {Retail, Self-stocking, Project-thru-Dist}

## What changes by platform (heads up)
- Charts are Plotly, not Chart.js — hover/zoom feel different
- "Drawer" → right-side slide-over detail panel driven by chart/KPI/table clicks
- Persistence is via Cloudflare R2 (S3-compatible Storage), not browser IndexedDB
- Streamlit reruns the script on each filter change — first interaction may feel slower

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env
# Edit .env with your Cloudflare R2 endpoint + access keys + bucket name
# In Cloudflare: create an R2 bucket "jsw-dashboard" and an API token (Object Read & Write)
streamlit run app.py
```

R2 config (env vars or Streamlit secrets): `R2_ENDPOINT` (or `R2_ACCOUNT_ID`),
`R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY`, `R2_BUCKET`. Persistence is optional —
the app runs without it, just without cross-session/cross-user file sharing.

## File map

```
app.py           # Streamlit entry — sidebar filters, tabs, KPI strip, BE tab, line items
data.py          # Order Excel parsing, enrich, channel classifier, num() with comma-strip
be_logic.py      # BE Excel parsing, flattenBeAtomic (dist-only), beActualsAgg
plots.py         # Plotly: channel trend, top states/distributors, India map, BE trajectory
storage_io.py    # Upload/download Excel binaries to Cloudflare R2 (boto3 / S3 API)
```

## Tabs

`Overview · Vs BE · Drill-down · Orders in Hand · Period compare · Line items`

## Deploy

- **Streamlit Community Cloud**: connect this repo, add the `R2_*` values as secrets, deploy.
- **Internal server**: any Python host (Render/Railway/VM). Run `streamlit run app.py --server.port 8501`.

## Iterating with Claude Code

All scaffold `TODO:` extension points have shipped (PRs #5–#9):

- ✅ **Period Compare tab** — two date windows, A-vs-B KPI deltas, per-channel comparison, daily overlay chart
- ✅ **Drill-down tab** — multi-level dimension picker, Plotly sunburst, pivot table with row drill
- ✅ **Daily invoice trajectory chart** on the Vs BE tab (clickable to drill a day's invoiced lines)
- ✅ **Inline drill panel** — right-side slide-over drawer wired to KPI tiles, chart clicks, and table rows
- ✅ **CSV exports** on every table / chart card (`st.download_button` over the underlying frames)

The data layer is complete and the UI now exposes all of the above. New work
should extend from these tabs rather than re-implement the scaffold TODOs.
