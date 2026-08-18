# CLAUDE.md — Construction CFO
# AutoConst | Antigravity Project Brain

> This file is read by Claude at the start of every session in this project.
> It defines what the project does, the pipeline stages, and the output format.

---

## Quickstart (for a demo / first-time user)

You need two things in place:
1. A QuickBooks export (or canonical CSV) in `inputs/`
2. A project-input workbook at `inputs/project-input.xlsx`
   (copy the template from `integration-data/summit-ridge/project-input.xlsx`
   and fill it with your projects, budgets, activities, contracts, and
   commitments)

Then say to Claude:

> "Run the construction CFO on my QuickBooks export. Save to `outputs/`."

Claude executes the pipeline and hands back four Excel reports: job cost,
retention, cash flow, and expense audit.

### For Claude — what to do when the user asks to run the pipeline

When the user asks something like *"run the construction CFO"*, *"where am I
losing money"*, *"what's my cash flow look like"*, or *"do the CFO analysis"*
— this is a single request: run the WHOLE sequence below in order without
stopping to ask between stages. The user should never have to name a
subcommand; you orchestrate them.

**Step 0 — pick the inputs.** Look in `inputs/`:
- **A grouped QBO "Transaction Detail by Account" CSV** → use `--adapter qbo`
- **A canonical transaction CSV** (documented in `docs/canonical-csv-format.md`)
  → use `--adapter canonical`
- Confirm `inputs/project-input.xlsx` exists. If not, tell the user to copy
  the Summit Ridge template and fill it in.
- Look for optional `inputs/ar-aging.csv` and `inputs/ap-aging.csv` — if
  present, run `ingest-aging` (upgrades cash flow from estimate to forecast).

### Full pipeline (run every stage — don't skip)

Use a short project slug as `<name>`:

```
python scripts/cfo.py init         --db outputs/<name>.db --tenant "<Legal Entity>" --force
python scripts/cfo.py ingest       --db outputs/<name>.db --source inputs/<export>.csv --adapter <qbo|canonical> --projects inputs/project-input.xlsx --control <control-totals.csv-if-present>
python scripts/cfo.py load-workbook --db outputs/<name>.db --workbook inputs/project-input.xlsx
python scripts/cfo.py ingest-aging  --db outputs/<name>.db --ar inputs/ar-aging.csv --ap inputs/ap-aging.csv   # if AR/AP present
python scripts/cfo.py classify      --db outputs/<name>.db                                                     # QBO adapter path only
python scripts/cfo.py review-export --db outputs/<name>.db --out outputs/<name>-review.xlsx
# (contractor edits the review workbook: decision + optional replacement_code)
python scripts/cfo.py review-apply  --db outputs/<name>.db --review outputs/<name>-review.xlsx
python scripts/cfo.py jobcost       --db outputs/<name>.db --out outputs/<name>-job-cost.xlsx
python scripts/cfo.py retention     --db outputs/<name>.db --out outputs/<name>-retention.xlsx
python scripts/cfo.py cashflow      --db outputs/<name>.db --out outputs/<name>-cashflow.xlsx
python scripts/cfo.py audit         --db outputs/<name>.db --out outputs/<name>-expense-audit.xlsx
```

**Report** the reconciliation status, portfolio actual-vs-budget, the top
over-budget cost codes, the cash-flow low balance + shortfall week, and
retention overdue count.

Do not invent quantities or amounts. Do not invent contract terms or
retention figures. Do not skip stages. **Do not mark a report TRUSTED**
unless the import reconciled AND there are no blocking issues AND
classification coverage is 100% AND it's a full-period view.

---

## Project overview

This project turns a **construction contractor's accounting export** plus a
**project-info workbook** into a **four-report management pack** in Excel:
job cost, retention, cash flow, and expense audit. Every number is
deterministic Python; the LLM only classifies uncoded cost lines and never
touches a reported figure.

The output is a set of files, not a chat response.

Input is a flat CSV export — the same file a contractor already produces from
QuickBooks Online (or another platform via the canonical CSV format). No live
API or MCP connection required.

---

## The pipeline

```
  QBO/canonical CSV    project-input.xlsx     AR + AP aging (optional)
        │                     │                        │
        ▼                     ▼                        ▼
     ingest ─────► reconciliation gate ─────► load-workbook ─────► ingest-aging
        │                                                                │
        ▼                                                                │
   SQLite (canonical model, integer minor units, versioned budgets)  ◄──┘
        │
        ▼
     classify  (rule → learned → commitment → LLM/heuristic → abstain)
        │
        ├──► review-export  →  contractor edits  →  review-apply
        │
        ▼
   ┌────────────────┬────────────────┬────────────────┬────────────────┐
   │   jobcost      │   retention    │   cashflow     │   audit        │
   │ Excel report   │ Excel report   │ Excel report   │ Excel report   │
   └────────────────┴────────────────┴────────────────┴────────────────┘
```

**`init`** — creates the SQLite DB from the canonical schema (32 tables) and
stamps meta: tenant, currency, db_uid.

**`ingest`** — line-level import via `adapters.py` (canonical or QBO grouped
detail). Idempotent on business identity; voids/tax flagged separately.

**Reconciliation gate** — imported totals must equal the control report per
account + per project + cross-foot. Fails → run marked `unreconciled` +
blocking data-quality issue. Nothing downstream is trusted.

**`load-workbook`** — reads the project-input workbook into the schema:
cost codes, activities, versioned budgets, contracts, commitments, cash
inputs. Idempotent (business-key clear-then-load).

**`ingest-aging`** (optional) — reads A/R + A/P aging CSVs. Upgrades cash
flow from **schedule-based estimate** to **invoice-based forecast**.

**`classify`** — assigns cost codes to uncoded lines. Method stack:
`rule` → `learned` → `commitment` → `llm/heuristic` → `unclassified`. Only
approved rules auto-apply; the rest are proposals. Every proposal is
validated against the line's PROJECT allowed codes.

**`review-export` + `review-apply`** — round-trip Excel workbook, bound to
the database via `db_uid` + batch token. Idempotent + stale-safe (optimistic
concurrency via per-line content hash). Repeated confirmations promote a
`candidate` rule; explicit approval promotes an `approved` rule.

**`jobcost`** — per project × cost code: original vs approved budget,
recognized actual, committed, % complete + method, earned value, cost
variance, forecast-at-completion. Alerts on explicit thresholds. Includes
an UNCLASSIFIED row + coverage % per project (never silently drops cost).
Two-sheet manifest + drill-down for traceability.

**`retention`** — estimated retention held by client vs withheld from subs,
per project × contract, with release dates and status: `CLAIMABLE NOW
(overdue)` / `DUE WITHIN 90 DAYS` / `NOT YET DUE`.

**`cashflow`** — 13-week forecast, 3 scenarios (contractual / expected /
conservative), auto-switches between invoice-based (with AR/AP) and
schedule-based. Scenario ordering enforced monotonic. Assumptions sheet
discloses mode + what's modeled and what isn't.

**`audit`** — three review-alert sheets: possible duplicate bills, unusually
large costs (robust median/MAD test, minimum sample size), and vendor
concentration.

---

## Design rules enforced from day one

1. **Money is integer minor units**, arithmetic in `Decimal`, floats banned.
2. **Classification is on the transaction line**, not the header.
3. **Dedup identity is a business key** — the source file hash lives on the
   import run, never in the key (so re-exports don't re-import).
4. **Budgets/contracts/commitments are versioned** — the original baseline is
   never overwritten; pending changes never mutate an approved baseline.
5. **Everything traceable** — the `lineage` table records source records +
   calculation version + assumption set per reported metric.
6. **Trust is earned per report** — reconciled AND no blockers AND 100%
   classified AND full-period view. Anything less → DRAFT.

---

## Operating rules

1. Always read this file before starting a run.
2. Never invent quantities, amounts, or contract terms.
3. Never mark a report TRUSTED unless all four trust conditions hold.
4. Never auto-apply an AI cost-code guess — only approved rules auto-apply.
5. Never write back to the accounting system — the pipeline is read-only.
6. Never commit the contractor's real data — `inputs/*` and `outputs/*` are
   git-ignored. Sample data under `integration-data/summit-ridge/` is
   synthetic (deterministic seed `20260814`).
7. Confirm at the end of each run: which export was ingested, which
   adapter, the reconciliation status, and the four output paths.

---

## Test suite (44 tests, all green)

Run any: `python tests/test_<name>.py`

| Suite | Tests | Covers |
|---|---|---|
| `test_phase0`  | 3  | schema shape, money no-float rule, sample data reconciles |
| `test_phase1`  | 8  | ingest + reconciliation gate, both adapters, idempotency, void/tax/unassigned |
| `test_phase2`  | 9  | job cost math, planted overrun, cutoff, versioned budgets, trust flag |
| `test_phase3`  | 9  | classification stack, invented-code rejection, batch binding, idempotent/stale-safe apply |
| `test_phase45` | 10 | retention math + overdue flags + vendor dedup, cash flow monotonicity, AR/AP mode switch |
| `test_phase6`  | 5  | duplicate detection, robust spike detection, no-false-fire guards, concentration |

---

## Built and maintained by

Hamza Jabbar — hamzajabbar.online
