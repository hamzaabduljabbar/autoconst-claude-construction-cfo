# Construction CFO

**A local financial-intelligence pipeline for a small-to-mid construction contractor.**

Reads a QuickBooks-shaped accounting export (or a canonical CSV) plus a
structured project workbook, checks the totals reconcile, classifies every
cost line to a budget code, and produces four Excel management reports:

- **Job Cost** — which jobs are over budget, per project × cost code
- **Retention** — estimated retention held by clients vs withheld from subs, with overdue-release flags
- **Cash Flow** — 13-week forecast, 3 scenarios (contractual / expected / conservative)
- **Expense Audit** — possible duplicate bills, unusual costs (robust MAD test), vendor concentration

Every number is deterministic Python; the LLM only classifies uncoded cost
lines and is never used as a calculator. Nothing is written back to the
accounting system — the pipeline is read-only by design.

## Requirements

- Python 3.10+
- `pip install -r requirements.txt` (openpyxl)

## Quickstart

```bash
# 1. clone
git clone https://github.com/hamzaabduljabbar/autoconst-claude-construction-cfo
cd autoconst-claude-construction-cfo

# 2. install
pip install -r requirements.txt

# 3. run the full pipeline against the committed synthetic dataset
SR=integration-data/summit-ridge
python scripts/cfo.py init          --db outputs/demo.db --tenant "Summit Ridge Construction Pty Ltd" --force
python scripts/cfo.py ingest        --db outputs/demo.db --source $SR/canonical/canonical-transactions.csv \
                                    --adapter canonical --projects $SR/project-input.xlsx \
                                    --control $SR/qbo-exports/control-totals.csv
python scripts/cfo.py load-workbook --db outputs/demo.db --workbook $SR/project-input.xlsx
python scripts/cfo.py ingest-aging  --db outputs/demo.db --ar $SR/qbo-exports/ar-aging.csv --ap $SR/qbo-exports/ap-aging.csv
python scripts/cfo.py jobcost       --db outputs/demo.db --out outputs/demo-job-cost.xlsx
python scripts/cfo.py retention     --db outputs/demo.db --out outputs/demo-retention.xlsx
python scripts/cfo.py cashflow      --db outputs/demo.db --out outputs/demo-cashflow.xlsx
python scripts/cfo.py audit         --db outputs/demo.db --out outputs/demo-expense-audit.xlsx
```

Full run-through and design notes are in `SETUP.md` and `PROJECT-BRIEF.md`.

## Bring your own data

Two required inputs go in `inputs/`:

1. **Accounting export** — a QuickBooks Online *Transaction Detail by Account*
   CSV (use `--adapter qbo`) or a canonical CSV (`docs/canonical-csv-format.md`).
2. **Project-input workbook** at `inputs/project-input.xlsx` — copy the template
   from `integration-data/summit-ridge/project-input.xlsx` and fill in your
   projects, activities, budgets, contracts, and commitments.

Optional (turns the cash forecast from an estimate into a real invoice-based
forecast):

- `inputs/ar-aging.csv` — open receivables (`Customer Balance Detail` export)
- `inputs/ap-aging.csv` — open payables (`Vendor Balance Detail` export)

`inputs/*` and `outputs/*` are git-ignored — real contractor data never gets
committed.

## What it does and doesn't claim

**Real, from posted transactions:** job cost, portfolio variance, reconciled
grand totals, invoice-based cash flow when A/R + A/P are loaded.

**Deliberately labeled as estimates:** retention (derived from % complete, not
certified claims), cash flow when only the transaction detail is loaded
(schedule-based estimate mode).

**Labeled as review alerts, not accusations:** expense-audit output.

**Deliberately not supported:** writing back to the accounting system;
universal accounting-platform support (v1 is QuickBooks Online CSV + canonical);
statistical confidence intervals; unattended email delivery.

## Design rules

- Money is integer minor units; arithmetic in `Decimal`; floats banned.
- Classification is at the transaction-line level, not the header.
- Dedup uses business identity (`tenant + platform + txn_type + txn_id`) —
  the source file hash lives on the import run, not in the key.
- Budgets / contracts / commitments are versioned. Pending never mutates
  approved.
- A report is stamped `TRUSTED` only if reconciled AND no blocking issues
  AND 100% classification coverage AND full-period. Otherwise it says `DRAFT`.
- The classifier's LLM path is bounded by the project's allowed cost codes —
  it can never introduce a code outside the chart.

## Tests

44 acceptance tests across 6 phases, all green:

```bash
python tests/test_phase0.py     # 3  — schema + money rules + sample reconcile
python tests/test_phase1.py     # 8  — ingest + reconciliation gate
python tests/test_phase2.py     # 9  — job cost
python tests/test_phase3.py     # 9  — classification loop + review round-trip
python tests/test_phase45.py    # 10 — retention + cash flow + AR/AP wiring
python tests/test_phase6.py     # 5  — expense audit detectors
```

## Repository layout

```
scripts/           the deterministic Python engine + adapters + report builders
integration-data/  reproducible synthetic contractor (Summit Ridge, seed 20260814)
sample-data/       small Phase-0 fixture
tests/             44 acceptance tests
docs/              canonical CSV format + design notes
inputs/, outputs/  runtime — git-ignored
```

## Status

All six build phases done. See `REVIEW-PACKET.md` for the independent-review
handoff.

## Author

Hamza Jabbar — hamzajabbar.online
