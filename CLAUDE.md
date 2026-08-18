# CLAUDE.md — Construction CFO

> This file is read by Claude at the start of every session in this project.

---

## What this project does

Turns a builder's accounting data + a workbook of their project info into
four Excel reports: **Job Cost**, **Retention**, **Cash Flow**, and
**Expense Audit**. Every number traces back to the transaction it came
from. Nothing is written back to the accounting system.

---

## First time here? Read this to the user

If the user has never used this project before, say to them:

> *"Drop two files into `inputs/`:*
>
> *1. Your accounting export — from QuickBooks: Reports → Transaction
> Detail by Account → Export → CSV*
> *2. Your project workbook — copy the example from*
> *`integration-data/summit-ridge/project-input.xlsx` to*
> *`inputs/project-input.xlsx` and fill in your own projects, budgets,*
> *and contracts.*
>
> *Then just tell me to run the CFO on your data. I'll produce four*
> *Excel reports in `outputs/`."*
>
> *Want to see it work first? Say "run the CFO on the example dataset"*
> *— I'll do a full demo run on the included Summit Ridge builder in*
> *about ten seconds, no setup needed.*

---

## How to run it

When the user asks something like *"run the construction CFO"*, *"where
am I losing money"*, *"do the CFO analysis"*, *"cash flow forecast"*, or
*"which jobs are over budget"* — treat it as a single request. Run every
stage below in order without stopping to ask between steps.

### Step 0 — Find the inputs

Look in `inputs/`:

- **A QuickBooks "Transaction Detail by Account" CSV** → use `--adapter qbo`
- **A canonical CSV** (see `docs/canonical-csv-format.md`) → use `--adapter canonical`
- Confirm `inputs/project-input.xlsx` exists. If not, tell them to copy the
  example from `integration-data/summit-ridge/project-input.xlsx`.
- Optional: `inputs/ar-aging.csv` + `inputs/ap-aging.csv` (open invoices).
  If present, run `ingest-aging` — this upgrades cash flow from an estimate
  to a real forecast.

If the user said *"run on the example dataset"*, use the files under
`integration-data/summit-ridge/` instead — that's the demo builder.

### Step 1 — Run the pipeline

Use a short slug for the builder as `<name>` (e.g. their initials):

```
python scripts/cfo.py init          --db outputs/<name>.db --tenant "<Company Name>" --force
python scripts/cfo.py ingest        --db outputs/<name>.db --source inputs/<export>.csv --adapter <qbo|canonical> --projects inputs/project-input.xlsx --control <control-totals.csv-if-any>
python scripts/cfo.py load-workbook --db outputs/<name>.db --workbook inputs/project-input.xlsx
python scripts/cfo.py ingest-aging  --db outputs/<name>.db --ar inputs/ar-aging.csv --ap inputs/ap-aging.csv    # only if AR/AP present
python scripts/cfo.py classify      --db outputs/<name>.db                                                       # only for QBO exports
python scripts/cfo.py review-export --db outputs/<name>.db --out outputs/<name>-review.xlsx
# — contractor edits the review workbook (decision + optional replacement_code) —
python scripts/cfo.py review-apply  --db outputs/<name>.db --review outputs/<name>-review.xlsx
python scripts/cfo.py jobcost       --db outputs/<name>.db --out outputs/<name>-job-cost.xlsx
python scripts/cfo.py retention     --db outputs/<name>.db --out outputs/<name>-retention.xlsx
python scripts/cfo.py cashflow      --db outputs/<name>.db --out outputs/<name>-cashflow.xlsx
python scripts/cfo.py audit         --db outputs/<name>.db --out outputs/<name>-expense-audit.xlsx
```

### Step 2 — Report back honestly

Give the user, in plain English:

- **Reconciliation status.** Did the imported totals match to the cent?
  If not, name the mismatch. Every report is stamped `DRAFT — NOT TRUSTED`
  when this fails.
- **Job cost.** Total actual vs approved budget across the portfolio. Then
  list the projects and cost codes flagged over-budget (orange rows) with
  the dollar amount over. Point them at the story — often it's the same
  cost category blowing out across several projects.
- **Retention.** Estimated retention owed to them vs held from subs. If any
  release is `CLAIMABLE NOW (overdue)`, list those first — that's money to
  chase today. Always call retention an *estimate* (it's calculated from
  % complete, not from certified claims).
- **Cash flow.** Whether it's an **invoice-based forecast** (real open A/R
  and A/P loaded) or a **schedule-based estimate** (only transactions). The
  low balance and week per scenario. If a shortfall is predicted, name the
  week. If it's still an estimate, tell them: *"drop your Customer Balance
  Detail and Vendor Balance Detail exports in `inputs/` as `ar-aging.csv`
  and `ap-aging.csv` and re-run for a real forecast."*
- **Expense audit.** Counts of duplicate alerts, unusual costs, and vendor
  concentration flags. Always call these **review alerts, not confirmed
  errors** — a human eyeballs before acting.

Point to the four Excel files as clickable links.

---

## Rules you must not break

1. **Never invent a number.** If a total doesn't reconcile, say so. If a
   line has no cost code, leave it in the UNCLASSIFIED row — don't guess.
2. **Never mark a report TRUSTED** unless: the import reconciled AND there
   are no blocking data-quality issues AND classification coverage is 100%
   AND it's a full-period view. Anything less → `DRAFT`.
3. **Never auto-apply an AI cost-code guess.** Only approved deterministic
   rules auto-apply. Everything else is a proposal that waits for the
   contractor's sign-off in the review workbook.
4. **Retention is always an estimate.** Never call it "certified" or "owed"
   as if it were a bill. It's derived from % complete.
5. **Audit output is a review alert, never a confirmed error.** No accusation
   language.
6. **Never write back to the accounting system.** The pipeline is read-only.
7. **Never put the contractor's real data anywhere it might get shared.**
   `inputs/` and `outputs/` are git-ignored. Don't upload them.

---

## What each report actually shows

**Job Cost** (`<name>-job-cost.xlsx`) — every project × cost code, side by
side: original budget, approved budget (v0 + approved revisions), actual
spent, committed, % complete, earned value, cost variance, forecast at
completion, and an OVER-BUDGET flag. Orange rows are the problems.
Includes a Report Info sheet (the manifest — what data went in, what
reconciled) and a Transaction Detail sheet (every cost line, so any figure
in the report drills back to its source).

**Retention** (`<name>-retention.xlsx`) — one row per contract:
estimated retention held by client, estimated retention held from subs,
release date, and a status column: `CLAIMABLE NOW (overdue)` (red),
`DUE WITHIN 90 DAYS` (yellow), or `NOT YET DUE`.

**Cash Flow** (`<name>-cashflow.xlsx`) — a Summary sheet (lowest balance
+ shortfall week per scenario), a Weekly sheet (each week's inflows,
outflows, and closing balance), and an Assumptions sheet that discloses
the mode (invoice-based or schedule-based) and exactly what's modeled.

**Expense Audit** (`<name>-expense-audit.xlsx`) — three sheets:
Duplicate Alerts, Unusual Costs, and Vendor Concentration. All are
review flags for a human to check.

---

## Author

Hamza Jabbar — hamzajabbar.online
