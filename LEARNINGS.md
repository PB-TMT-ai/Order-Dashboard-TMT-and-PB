# Project Learnings

Track errors, solutions, and insights. System gets smarter with each entry.

## Format
Date | Component | Issue | Resolution | Insight

---

## Entries

(Add new entries at top)

2026-06-08 | data / app (Item 7 revised) | "Cancelled" was defined by the order-status flag (`_sta == "Cancelled"`), dropping whole lines. Business definition: cancelled = the rejected/cancelled QUANTITY ('total cancelled qty' column), independent of status — a line can be partially rejected while still Open. | Net the rejected qty into the active quantity at the single source (`data._q = max(gross − _cq, 0)`; keep `_qg` gross and `_cq` for reporting), drop the status-based split, and report `Σ _cq` as the Cancelled KPI. Pending simplifies to `max(_q − _rq, 0)`; removed the `_sta != "Cancelled"` guard in `be_logic`. | Prefer netting a quantity at the source frame over row-dropping — partial cancellations then flow correctly through every metric, and no view needs to know about status. Don't anchor a definition on a status flag when the source has an explicit quantity column.

2026-06-08 | app (BE tab) | Needed an editable "Realistic BE override" column AND click-to-drill on the same BE table, but `st.data_editor` has no `on_select` row-selection event (only `st.dataframe` does). | Render the table with `st.data_editor` (only the override column enabled), persist edits into a `session_state` map and `st.rerun()` once when they change so Realistic-used / Volume-at-Risk refresh; move drill-down to a separate `st.selectbox`. | You can't get both inline editing and row-selection from one Streamlit table widget — split editing (data_editor) from selection (dataframe/selectbox).

2026-06-08 | data / app (Item 7) | Cancelled orders were inflating Ordered MT and order metrics. | Split `_sta == "Cancelled"` out of the master frame once, right after load (keep `df_cancelled` for a KPI card); everything downstream uses the cancelled-excluded `df`. | `float('nan')`-style truthiness aside, do global exclusions at the single source frame, not per-view, so KPIs/BE/drill all stay consistent. Short-close age threshold also unified behind `data.SHORT_CLOSE_DAYS` (60→30).

2026-06-02 | be_logic.be_table | `(g["_st"].iloc[0] or "")` crashed with `AttributeError: 'float' object has no attribute 'title'` when a no-BE distributor's ship-to state was NaN. | Replaced with `data.cl(...)` which maps NaN/None/sentinels to "". | `float('nan')` is truthy in Python, so `nan or ""` returns `nan`, not `""`. Use the `cl()` helper for any cell that may be NaN.

2026-06-02 | be_logic / app (BE tab) | BE-tab actuals didn't reconcile with the Invoiced KPI: eligible distributors with no BE were excluded from `matched_act`, and the no-BE table branch used order-date attribution while BE distributors used proportional invoice-date. | `be_actuals_agg` now also accumulates eligible no-BE actuals (`nobe_act`/`nobe_pipe`/`nobe_orders_by_key`) with the SAME proportional attribution; added `BeAggregate.total_act`; `be_table` reads no-BE actuals from `nobe_orders_by_key`; added a "Total eligible actuals" KPI tile + drill. | Reconciliation rule: BE-tab total Actuals == invoiced-in-BE-month for scope (TMT, Fe550/550D, channels rt/ss/pd). Keep gap-vs-BE on `matched_act` only (plan covers BE distributors).
