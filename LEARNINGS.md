# Project Learnings

Track errors, solutions, and insights. System gets smarter with each entry.

## Format
Date | Component | Issue | Resolution | Insight

---

## Entries

(Add new entries at top)

2026-06-02 | be_logic.be_table | `(g["_st"].iloc[0] or "")` crashed with `AttributeError: 'float' object has no attribute 'title'` when a no-BE distributor's ship-to state was NaN. | Replaced with `data.cl(...)` which maps NaN/None/sentinels to "". | `float('nan')` is truthy in Python, so `nan or ""` returns `nan`, not `""`. Use the `cl()` helper for any cell that may be NaN.

2026-06-02 | be_logic / app (BE tab) | BE-tab actuals didn't reconcile with the Invoiced KPI: eligible distributors with no BE were excluded from `matched_act`, and the no-BE table branch used order-date attribution while BE distributors used proportional invoice-date. | `be_actuals_agg` now also accumulates eligible no-BE actuals (`nobe_act`/`nobe_pipe`/`nobe_orders_by_key`) with the SAME proportional attribution; added `BeAggregate.total_act`; `be_table` reads no-BE actuals from `nobe_orders_by_key`; added a "Total eligible actuals" KPI tile + drill. | Reconciliation rule: BE-tab total Actuals == invoiced-in-BE-month for scope (TMT, Fe550/550D, channels rt/ss/pd). Keep gap-vs-BE on `matched_act` only (plan covers BE distributors).
