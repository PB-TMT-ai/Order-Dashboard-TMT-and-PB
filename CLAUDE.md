# Project: Order-Dashboard-TMT-and-PB

## Core Principle

You operate as the decision-maker in a modular system. Your job is NOT to do everything yourself. Your job is to read instructions, pick the right tools, handle errors intelligently, and improve the system over time.

Why? 90% accuracy across 5 steps = 59% total success. Push repeatable work into tested scripts. You focus on decisions.

## System Architecture

**Blueprints (/blueprints)** - Step-by-step instructions in markdown. Goal, inputs, scripts to use, output, edge cases. Check here FIRST.

**Scripts (/scripts)** - Tested, deterministic code. Call these instead of writing from scratch.

**Workspace (/.workspace)** - Temp files. Never commit. Delete anytime.

## How You Operate

1. Check blueprints first - If one exists, follow it exactly
2. Use existing scripts - Only create new if nothing exists
3. Fail forward - Error → Fix → Test → Update blueprint → Add to LEARNINGS.md → System smarter
4. Ask before creating - Don't overwrite blueprints without asking

## Tech Stack

- Python 3.11
- Streamlit (UI, run with `streamlit run app.py`)
- pandas + NumPy (data layer)
- Plotly (charts)
- openpyxl / python-calamine (Excel parsing)
- boto3 → Cloudflare R2, S3-compatible (optional persistence)

No Node/React build. An unused Next.js scaffold lingers under `/src`
(`.gitkeep` placeholders + `src/lib/logger.ts`) — the live app is the
top-level Python modules. Ignore `/src` unless you are intentionally
reviving that scaffold.

## Project Structure

Top-level Python modules (the app):

- `app.py` — Streamlit entry: upload, sidebar filters, KPI strip, tabs (Overview, Vs BE, BE Output, Drill-down, Orders in Hand, Period compare, Customers, Line items)
- `data.py` — Order Excel parsing, `enrich()`, channel classifier, filters
- `be_logic.py` — BE Excel parsing, plan-vs-actual aggregation (`be_actuals_agg`, `be_table`, `be_mom`)
- `be_output.py` — BE Output Report: AOP parsing, order-book accounting, monthly report + per-distributor table
- `customer_logic.py` — distributor/account buying-pattern analytics (RFM, cadence, MoM)
- `drawer.py` — right-side slide-over detail panel (`open_drawer` / `render`)
- `plots.py` — Plotly figures
- `storage_io.py` — Cloudflare R2 persistence (optional; app runs without it)
- `theme.py` — JSW One brand theme / Plotly defaults

Supporting dirs:

- `/blueprints` — task SOPs (check FIRST)
- `/scripts` — tested automation scripts (currently empty)
- `/tests` — `unit/` + `integration/` (no runner wired up yet)
- `/.streamlit` — Streamlit config/theme
- `/.workspace` — temp files, gitignored, never commit
- `.env.example` — R2 env vars (`R2_ENDPOINT`/`R2_ACCOUNT_ID`, `R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY`, `R2_BUCKET`)

## Code Standards

- `from __future__ import annotations`; type hints + explicit return types
- Functional style; `@dataclass` for state containers (e.g. `DrawerState`, `BeAggregate`)
- Vectorise pandas work; preserve the documented business rules exactly (channel classification, P&T vs TMT, short-close, invoice-date attribution)
- Streamlit reruns top-to-bottom on every interaction — define shared values (e.g. `now`) before the tabs, and seed widget `session_state` rather than relying on `del key` + `value=` to reset a widget
- Keep number formatting consistent (% ≤1 dp, MT 0 dp)
- Run a boot smoke before declaring done: `AppTest.from_file("app.py").run()` should have 0 exceptions

## Error Protocol

1. Stop and read the full error
2. Isolate - which component/script failed
3. Fix and test
4. Document in LEARNINGS.md
5. Update relevant blueprint

## What NOT To Do

- Don't skip blueprint check
- Don't ignore errors and retry blindly
- Don't create files outside structure
- Don't write from scratch when blueprint exists
