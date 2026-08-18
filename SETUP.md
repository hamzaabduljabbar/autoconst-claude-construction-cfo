# Construction CFO — Setup & Phase 0

The deterministic Python engine is the real product; the Claude skill orchestrates it.
Phase 0 stands up the foundation: the SQLite schema, the money discipline, a
reproducible sample dataset, and `init`/`check` commands. No reports yet — those
are Phase 2 onward.

## Requirements

- Python 3.10+
- `pip install -r requirements.txt` (openpyxl)

## Commands available now

```bash
# create the database from the canonical schema
python scripts/cfo.py init --db outputs/demo.db --tenant "Northwind Construction Pty Ltd"

# verify the database has all expected tables + meta
python scripts/cfo.py check --db outputs/demo.db

# (re)generate the sample dataset into sample-data/
python scripts/make_sample_data.py

# run the Phase 0 acceptance checks
python tests/test_phase0.py
```

## What Phase 0 delivers

| File | Purpose |
|---|---|
| `scripts/schema.sql` | The 29-relation canonical schema (§15 of the brief). Integer minor units for money, line-level classification, business-identity dedup key, versioned budgets/contracts/commitments, lineage + data-quality tables. |
| `scripts/money.py` | Money helper. Stores integer minor units, does arithmetic in `Decimal`, and **refuses to convert a float** so binary-float error can never enter a total. |
| `scripts/cfo.py` | CLI. `init` builds + stamps the DB; `check` verifies its shape. |
| `scripts/make_sample_data.py` | Generates an internally-consistent sample: canonical CSV, a QBO-flavoured export of the same data, control totals **computed from the lines**, and the structured project-input workbook. |
| `sample-data/` | The generated dataset (committed so Phase 1 has a fixture). |
| `tests/test_phase0.py` | Schema shape, money rules, and the sample-data reconciliation cross-check. |

## The sample scenario

`Northwind Construction` with three active projects (Northgate Warehouse,
Bayside Aged Care, Clearview Bridge). 11 transactions / 27 lines, deliberately
exercising the cases Phase 1's reconciliation gate must handle:

- a single bill split across **two projects plus a company-level line**
- **GST/tax lines kept separate** from cost (never folded in)
- a **credit note** (negative)
- a **voided** transaction (present, flagged, excluded from totals)
- a **vague-memo** line the classifier should abstain on (`UNCLASSIFIED`)

Control totals cross-foot: account totals and project totals each sum to the
same `cost_ex_tax` grand total. That is the target Phase 1 reconciles against.

## Design rules enforced from day one

1. **Money is integer minor units**, arithmetic in `Decimal`, floats banned.
2. **Classification is on the transaction line**, not the header.
3. **Dedup identity is a business key** — the source file hash lives on the
   import run, never in the key (so re-exports don't re-import).
4. **Budgets/contracts/commitments are versioned** — the original baseline is
   never overwritten; pending changes never mutate an approved baseline.
5. **Everything traceable** — the `lineage` table records source records +
   calculation version + assumption set per reported metric.

## Phase 1 — ingest + reconciliation gate (done)

Adapters normalize any supported export into one line shape, then an idempotent
importer loads it and a reconciliation gate proves the totals before anything is
trusted.

```bash
# clean canonical CSV
python scripts/cfo.py ingest --db outputs/demo.db \
  --source sample-data/canonical-transactions.csv --adapter canonical \
  --projects sample-data/project-input.xlsx \
  --control sample-data/control-totals.csv

# the SAME data as a messy real-world QBO 'Transaction Detail by Account' export
python scripts/cfo.py ingest --db outputs/demo.db \
  --source sample-data/qbo-transaction-detail-messy.csv --adapter qbo \
  --projects sample-data/project-input.xlsx \
  --control sample-data/control-totals.csv

# reconcile an already-imported db on its own
python scripts/cfo.py reconcile --db outputs/demo.db --control sample-data/control-totals.csv

# acceptance tests #1-7
python tests/test_phase1.py
```

Both fixtures reconcile to the identical **AUD 105,210** cost ex-tax / **9,071**
GST — proving the messy grouped QBO export reconstructs to the same truth as the
clean canonical file.

| File | Purpose |
|---|---|
| `scripts/adapters.py` | `parse_canonical` + `parse_qbo`. All platform mess (group-header accounts, `$`/`(...)` amounts, `MM/DD/YYYY`, subtotal & blank rows, voids as `$0.00`) lives here; both yield the same normalized line. |
| `scripts/ingest.py` | Idempotent line-level import (business-identity dedup, update-in-place on change, void/tax flagging) + the reconciliation gate. |
| `sample-data/qbo-transaction-detail-messy.csv` | A realistic QBO export used to harden the adapter without needing anyone's real books. |

**What the gate enforces:** account totals and project totals each reconcile to
the named control report; the account grand total must cross-foot to the project
grand total; tax is checked separately and never enters cost; voids are present
but excluded; unassigned-project cost stays visible. A single wrong figure marks
the run `unreconciled` and records a blocking `data_quality_issue`.

### The one real-data test worth doing now

The adapter is hardened against synthetic mess. The remaining unknown is whether
a *real* QBO "Transaction Detail by Account" CSV matches the shape `parse_qbo`
expects. Export one from a real (ideally de-identified) QBO account, drop it in
`inputs/`, and run the `qbo` adapter against it. Mismatches point straight at the
adapter — everything downstream is already proven against the fixture.

## Phase 2 — job cost report (done)

Loads the project-side input workbook into the schema, then produces a
deterministic job-cost report per project × cost code.

```bash
SR=integration-data/summit-ridge
python scripts/cfo.py init --db outputs/summit-ridge.db --tenant "Summit Ridge Construction Pty Ltd" --force
python scripts/cfo.py ingest --db outputs/summit-ridge.db --source $SR/canonical/canonical-transactions.csv \
  --adapter canonical --projects $SR/project-input.xlsx --control $SR/qbo-exports/control-totals.csv
python scripts/cfo.py load-workbook --db outputs/summit-ridge.db --workbook $SR/project-input.xlsx
python scripts/cfo.py jobcost --db outputs/summit-ridge.db --out outputs/summit-ridge-job-cost.xlsx

python tests/test_phase2.py
```

| File | Purpose |
|---|---|
| `scripts/load_workbook.py` | Loads cost codes, activities, progress, versioned budgets, contracts + changes, and commitments from the input workbook into the schema. |
| `scripts/jobcost.py` | Computes original vs approved budget, recognized actual (credits net, tax + voids excluded), committed + remaining, % complete (latest progress, with method), earned value, cost variance, forecast-at-completion (method shown), and OVER-BUDGET alerts on explicit thresholds. Writes a formatted, formula-injection-sanitized Excel sheet, titled `UNTRUSTED` if the import did not reconcile. |

**Report columns (§7.1):** Original Budget · Approved Budget · Committed ·
Recognized Actual · % Complete · Progress Method · Earned Value · Cost Variance
· FAC Method · Forecast @ Completion · Alert.

**Validated against the answer key:** every number is deterministic (no LLM); the
planted MCF (PRJ-002) labour overrun surfaces as a `01-50-00 Site Services` alert
(AUD 319,329 actual vs 89,600 approved), report actuals tie back to the tagged
cost lines in the DB.

**Hardening pass (after independent review):**
- `--as-of DATE` reporting cutoff now actually filters transactions, progress
  updates, and approved budget revisions (a July report cannot include August cost).
- Unclassified project cost is shown as an **UNCLASSIFIED row + coverage %** per
  project — never silently dropped (Summit Ridge: AUD 2,047.45, coverage 99.96%).
- `load-workbook` is **idempotent** (business-key clear-then-load) — re-running no
  longer doubles commitments/contracts/progress.
- Recognized actual **excludes non-cost account types** (income/asset/liability/
  equity) as a guard against mis-coded balance-sheet lines.
- The unsafe `remaining = committed − all actuals` figure was **removed** (it
  subtracted unrelated wages/materials); deferred until commitments can be linked.
- **Trust** now requires reconciled AND no blocking issues AND 100% coverage AND a
  full-period view — a date-cutoff view is analytical/DRAFT by definition.
- Two new sheets give **visible lineage**: `Report Info` (manifest: cutoff, trust
  criteria, import-run ids, workbook hash, coverage) and `Transaction Detail`
  (every cost line with a stable `sourceTxn:lineId` drill-down key).
- Real **versioned-budget test** (v0 preserved, approved v1 current, pending v2
  excluded, cutoff falls back to v0) since Summit Ridge only carries v0.

## Phase 3 — classification loop (done)

Codes the uncoded cost lines a QBO import produces, so they feed the same job-cost
report the canonical data already does.

```bash
SR=integration-data/summit-ridge
python scripts/cfo.py init --db outputs/c.db --tenant "Summit Ridge Construction Pty Ltd" --force
python scripts/cfo.py ingest --db outputs/c.db --source $SR/qbo-exports/qbo-transaction-detail-current.csv \
  --adapter qbo --projects $SR/project-input.xlsx --control $SR/qbo-exports/control-totals.csv
python scripts/cfo.py load-workbook --db outputs/c.db --workbook $SR/project-input.xlsx

python scripts/cfo.py classify       --db outputs/c.db                          # rule/learned/commitment auto; rest proposed
python scripts/cfo.py review-export   --db outputs/c.db --out outputs/review.xlsx
# contractor edits review.xlsx: decision = accept / replace (+ replacement_code)
python scripts/cfo.py review-apply    --db outputs/c.db --review outputs/review.xlsx

python tests/test_phase3.py
```

| File | Purpose |
|---|---|
| `scripts/classify.py` | The method stack + review round-trip + rule promotion + an evaluation harness that scores precision against known truth. |

**The stack (brief §6), tried in order per uncoded line:**

1. `rule` — approved mapping rule authored directly · deterministic, auto-apply
2. `learned` — mapping rule promoted from a past human review · deterministic, auto-apply
3. `commitment` — unique (project, vendor) subcontract commitment · deterministic, auto-apply
4. `llm` (injected proposer / Claude) or `heuristic` (built-in token-overlap fallback) · **proposed → human review**
5. `unclassified` — explicit abstain, never force-fit

Deterministic tiers apply directly; proposals wait for sign-off in a round-trip
review workbook; approving one **promotes a mapping rule**, so the next run codes
it deterministically. The proposer's output is always validated against the
allowed cost-code list — **Claude can never introduce a code outside the chart**.

**Safety hardening (independent review):**
- **Only approved rules auto-apply.** Commitment matches are now *proposals*, not
  silent auto-codes (a sub can invoice dayworks/materials outside its subcontract).
- **Project-scoped validation** — a proposal/replacement must be a code valid for
  that line's project (from its activities), not just anywhere in the tenant chart.
- **Deterministic, conflict-safe rules** — two equally-specific rules with
  different codes → abstain rather than guess.
- **Conservative promotion** — a single acceptance never creates an auto-applying
  rule. Rules are `candidate` until ≥2 consistent, conflict-free confirmations, or
  an explicit "approve rule" decision in the Rule Groups sheet.
- **Batch-bound workbooks** — each export is stamped with this database's `db_uid`
  + tenant + a `batch_token`; a workbook from another db is rejected.
- **Idempotent + stale-safe apply** — `UNIQUE(batch, classification)` plus a
  per-line content hash (optimistic concurrency): re-applying, or applying a stale
  workbook after the line moved, changes nothing (`already` / `stale` buckets).
- **Committed precision regression** — `expected/classification-truth.csv` (keyed
  by business identity) drives a test asserting commitment precision = 100%,
  heuristic ≥ 90%, and **false-auto-apply = 0**.

**Review workbook (4 sheets):** `Batch` (binding), `Rule Groups` (approve a whole
vendor→code mapping in one row — 595 lines collapse to ~72 groups), `Line Review`
(locked id/hash columns, accept/replace dropdowns, score as %), `Allowed Codes`
(project-scoped reference). Coverage is reported both **by line and by value**.

## Phase 4 — retention tracker (done)

```bash
python scripts/cfo.py retention --db outputs/demo.db --out outputs/retention.xlsx
python tests/test_phase45.py
```

Per project × contract: retention **held by the client** from us (an asset we get
back), retention **we withhold from subs** (a liability we pay out), the **release
date** per trigger (practical completion / defects-liability end), the net
position, and a "claimable soon" flag for releases within 90 days. On Summit
Ridge: **$460,569 owed to us**, **$131,824 we hold**, net **$328,745**.

## Phase 5 — cash flow forecast (done)

```bash
python scripts/cfo.py cashflow --db outputs/demo.db --out outputs/cashflow.xlsx
```

13 weekly buckets, 3 scenarios (contractual / expected = +median days late /
conservative = +2× median). Projects remaining contract-to-bill and cost-to-
complete over their schedule, timed by payment terms, plus retention releases,
payroll, overhead and the next GST payment. **Blocking-input check** → PARTIAL
mode if any are missing. An **Assumptions sheet discloses what is and isn't
modelled** (notably: inflows/outflows are modelled from remaining contract work +
schedule, not an open AR/AP ledger, which a QBO transaction-detail export doesn't
carry). On Summit Ridge it catches a real shortfall: **$284k opening → underwater
by Aug 29 → −$830k by late October**, driven by full overheads against only 2 of
6 jobs still billing.

| File | Purpose |
|---|---|
| `scripts/retention.py` | Retention tracker. |
| `scripts/cashflow.py` | Weekly 3-scenario cash flow forecast with disclosed assumptions. |

## Phase 6 — expense leak audit (done)

```bash
python scripts/cfo.py audit --db outputs/demo.db --out outputs/audit.xlsx
python tests/test_phase6.py
```

Three review sheets (all labeled REVIEW ALERTS, never "confirmed"): possible
**duplicate bills** (same vendor+amount+memo on separate bills), **unusually large
costs** (robust median/MAD test, ≥5-sample minimum so small vendors don't
false-fire), and **vendor concentration** (supply-agreement / bulk-buy candidates).

## A/R + A/P wiring — cash flow becomes real

```bash
python scripts/cfo.py ingest-aging --db outputs/demo.db --ar $SR/qbo-exports/ar-aging.csv --ap $SR/qbo-exports/ap-aging.csv
```

Loading open receivables/payables switches the cash flow from **schedule-based
estimate** to **invoice-based forecast**. On Summit Ridge the estimate warned of a
−$586k shortfall; the invoice-based forecast (real open A/R $848k in, A/P $370k out)
shows **no shortfall, low +$349k** — the difference between a scary estimate and a
bankable forecast.

## Status

All build phases done: ingest + reconcile · job cost · classification · retention ·
cash flow (invoice-based) · expense audit. **44 acceptance tests green**
(phase0–3, phase45, phase6). See `REVIEW-PACKET.md` for the independent-review
handoff.

## Realistic integration dataset

`integration-data/summit-ridge/` is a reproducible 18-month synthetic contractor
dataset built against the completed Phase 1 contracts. It supplements rather
than replaces the small deterministic fixture.

```bash
python scripts/make_realistic_data.py
node scripts/build_realistic_workbook.mjs
python tests/test_realistic_data.py
```

The fixed seed (`20260814`) generates six projects, 563 transaction headers and
1,244 transaction lines. Both the canonical and grouped-QBO representations
reconcile to AUD 6,986,858.47 ex-GST and AUD 552,535.51 GST. Three cumulative
canonical exports exercise overlapping re-imports and idempotency. Known cases
include a void, credit note, split-project bill, company-level costs, ambiguous
cost coding, and untrusted spreadsheet/prompt-like text.
