# Construction CFO — Independent Review Packet

This packet is everything a reviewer needs to assess the project end to end: what it
is, how to run it, what it claims (and deliberately doesn't), the known limitations,
and the test suite. It is written to be handed to a reviewer who has **not** seen the
build.

---

## 1. What this is

A **Claude-guided local financial pipeline** for a small/mid construction contractor.
It reads a QuickBooks export + a structured project workbook and produces four Excel
reports a CFO would make: **job cost, retention, cash flow, and an expense audit**.

Design rule (governing everything): **Python does the math, SQLite is the source of
truth, Claude only orchestrates and classifies.** No number in any report comes from
an LLM. The classifier is the one LLM touch-point and its output is always validated
against the chart of accounts.

## 2. Architecture

```
QuickBooks export (CSV)  ─┐
project-input.xlsx        ├─► adapters ─► SQLite (canonical model) ─► reports (Excel)
A/R + A/P aging (CSV)     ─┘        │            │                          │
                                reconcile     classify (rule→learned→        job cost
                                  gate         commitment→LLM→abstain)       retention
                                                                            cash flow
                                                                            expense audit
```

- **Money** is stored as integer minor units; arithmetic in `Decimal`; floats are
  banned (a helper raises `TypeError` if handed one).
- **Transactions are lines, not headers** (one bill splits across projects/codes).
- **Dedup** is on business identity; the file hash lives on the import run.
- **Budgets/contracts/commitments are versioned**; pending never mutates approved.
- Schema: `scripts/schema.sql` (32 tables).

## 3. How to run it (full pipeline)

```bash
pip install -r requirements.txt          # openpyxl only

SR=integration-data/summit-ridge
DB=review.db

python scripts/cfo.py init         --db $DB --tenant "Summit Ridge Construction Pty Ltd" --force
python scripts/cfo.py ingest       --db $DB --source $SR/canonical/canonical-transactions.csv \
                                   --adapter canonical --projects $SR/project-input.xlsx \
                                   --control $SR/qbo-exports/control-totals.csv
python scripts/cfo.py load-workbook --db $DB --workbook $SR/project-input.xlsx
python scripts/cfo.py ingest-aging  --db $DB --ar $SR/qbo-exports/ar-aging.csv --ap $SR/qbo-exports/ap-aging.csv
python scripts/cfo.py classify      --db $DB          # optional: needed only for the QBO (uncoded) path
python scripts/cfo.py jobcost       --db $DB --out out-jobcost.xlsx
python scripts/cfo.py retention     --db $DB --out out-retention.xlsx
python scripts/cfo.py cashflow      --db $DB --out out-cashflow.xlsx
python scripts/cfo.py audit         --db $DB --out out-audit.xlsx
```

The **QBO messy path** (real "Transaction Detail by Account" grouped export) uses
`--adapter qbo` and requires the classify → review → apply loop to code the lines
first — see §7.

## 4. The review artifacts (pre-generated)

In `review-artifacts/` — the four reports run on the Summit Ridge dataset:

| File | What to look at |
|---|---|
| `1-job-cost.xlsx` | Sheets: Job Cost, Report Info (manifest/lineage), Transaction Detail (drill-down). Orange = over budget. |
| `2-retention.xlsx` | Est. client vs sub retention, release dates, Status (overdue in red). |
| `3-cash-flow.xlsx` | Summary (3 scenarios), Weekly (Expected), Assumptions (discloses mode + what's modeled). |
| `4-expense-audit.xlsx` | Duplicate Alerts, Unusual Costs, Vendor Concentration — all "REVIEW ALERTS, not confirmed". |

## 5. Headline results on Summit Ridge (a 6-project, 18-month synthetic contractor)

- **Reconciliation:** imported cost ties to the control report to the cent — AUD 6,986,858.47 ex-GST / 552,535.51 GST.
- **Job cost:** flags every project over budget on `Site Services` (~AUD 847k total overrun). Marked **DRAFT** because 2 lines (AUD 2,047) are unclassified → 99.96% coverage.
- **Retention:** EST AUD 460,569 owed by clients, 131,824 held from subs; **4 releases overdue**.
- **Cash flow:** with real A/R+A/P loaded (invoice-based) → **no shortfall, low +AUD 349k**. Without it (schedule-based estimate) → −AUD 586k. The gap is the point.
- **Audit:** 1 possible duplicate, 0 cost spikes (data is clean), 12 vendor-concentration flags.

## 6. Claims — what we DO and DON'T assert

**We assert:**
- Job cost is real and complete from posted transactions + the contractor's budgets.
- Every reported number is deterministic and traceable to source records.
- Reconciliation is enforced; unreconciled or under-classified reports are marked DRAFT.
- The classifier never invents a code; deterministic tiers auto-apply, the rest wait for review.
- Cash flow is an **invoice-based forecast** when A/R+A/P are loaded, a **schedule-based estimate** otherwise — the report says which.

**We deliberately DO NOT assert:**
- That retention figures are **certified** — they are **estimated** from % complete; release dates are **planned**, not from completion certificates. The report is titled "Estimated Retention Position".
- That cash flow is a full treasury forecast — not-yet-invoiced work is excluded; it's a 13-week plan.
- That audit alerts are confirmed errors/fraud — they are **review alerts** for a human.
- Universal accounting-platform support — v1 is QuickBooks Online CSV + a canonical template.

## 7. Known limitations & assumptions (per report)

- **Job cost:** only lines with both a project and a cost code are attributed; the rest show as an UNCLASSIFIED row (never hidden). Forecast-at-completion method is `actual + remaining approved budget`, shown per row.
- **Classification (QBO path):** commitment matches are **proposals**, not auto-applied. Rules auto-apply only after explicit approval or ≥2 consistent confirmations. Review workbooks are bound to the database (`db_uid` + batch token), idempotent, and stale-safe.
- **Retention:** estimated from budget-weighted % complete × contract retention %. No certified claims / retention-ledger / actual-release events (a QBO transaction export doesn't contain them).
- **Cash flow:** inflows/outflows from open A/R+A/P by due date (+ scenario delay); payroll, overhead, GST layered on. GST amount excluded if not provided. Equal weekly billing (claim frequency not applied). Scenario ordering is enforced monotonic (contractual ≥ expected ≥ conservative every week).
- **Audit:** robust median/MAD test, ≥5-sample minimum, positive cost lines only.

## 8. Test suite (44 tests, all passing)

Run any: `python tests/test_<name>.py`

| Suite | Tests | Covers |
|---|---|---|
| `test_phase0` | 3 | schema shape, money no-float rule, sample data reconciles |
| `test_phase1` | 8 | ingest + reconciliation gate, both adapters, idempotency, void/tax/unassigned |
| `test_phase2` | 9 | job cost math, planted overrun, cutoff filtering, versioned budgets, trust flag, idempotent load |
| `test_phase3` | 9 | classification stack, invented-code rejection, batch binding, idempotent/stale-safe apply, conservative promotion, precision vs committed truth |
| `test_phase45` | 10 | retention math + overdue flags + vendor dedup, cash flow running-balance + scenario monotonicity + shortfall, A/R+A/P mode switch + idempotency |
| `test_phase6` | 5 | duplicate detection, robust spike detection, no-false-fire guards, small-vendor guard, concentration |

There is also a committed truth file `integration-data/summit-ridge/expected/classification-truth.csv` (keyed by business identity) that drives the classifier precision regression.

## 9. Suggested focus for the reviewer

Areas most worth challenging:
1. **Retention correctness** — is "% complete × contract × retention%" a defensible estimate, and are the release-date and overdue rules right?
2. **Cash-flow model** — is excluding not-yet-invoiced work the right call? Is the scenario delay model (median / 2×median days late) sound?
3. **Classifier safety** — can any path auto-apply a wrong code, or apply a stale/foreign review workbook? (See `test_phase3`.)
4. **Reconciliation gate** — can an under-classified or unreconciled dataset ever produce a "trusted" report?
5. **Audit thresholds** — are MAD z=3.5, 5-sample minimum, and the 14-day duplicate window reasonable, and what's the false-positive rate on real data?

## 10. Prior review history

This project has already been through five independent review rounds; each was
incorporated. Notable hardening: transaction-line model, versioned budgets, the
reconciliation gate, integer-minor-units money, the rule→learned→commitment→abstain
classifier, batch-bound/idempotent/stale-safe review application, conservative rule
promotion, project-scoped code validation, retention overdue flags + vendor dedup,
cash-flow scenario monotonicity + honest relabeling, and A/R+A/P wiring to make the
cash forecast invoice-based. See `PROJECT-BRIEF.md` (Appendix A) and `SETUP.md` for
the full trail.

---

*Everything here runs offline against the committed synthetic dataset. No real
contractor data is included. The genuinely un-real pieces (certified retention
claims, actual retention-release events) are labeled as estimates, not hidden.*
