---
name: construction-cfo
description: Turn a construction contractor's accounting export (QuickBooks Online or a canonical CSV) plus a structured project workbook into four reconciled Excel management reports — job cost, retention, cash flow, expense audit — with every number traceable to a source transaction. Use when the user asks for job costing, retention tracking, cash-flow forecasting, expense audits, "where am I losing money", "when will I run out of cash", or CFO-style analysis of accounting data. Reads a real QuickBooks 'Transaction Detail by Account' export or a canonical CSV; optionally reads open A/R and A/P aging exports to upgrade the cash forecast from schedule-based estimate to invoice-based forecast. All arithmetic is deterministic Python; the LLM only classifies uncoded transaction lines and never touches a reported number.
---

# Construction CFO

Deterministic financial pipeline for a small-to-mid construction contractor.
Ingests a QuickBooks-shaped accounting export + a project-info workbook, checks
that the totals reconcile, classifies each cost line to a budget code, then
produces four traceable Excel reports. The Python engine is the product; this
skill is the guided interface + orchestration layer.

The scripts live in a cloned copy of the repo:
https://github.com/hamzaabduljabbar/autoconst-claude-construction-cfo

---

## When to invoke this skill

Trigger phrases like:
- "run the construction CFO on my QuickBooks export"
- "which jobs am I losing money on?" / "job cost report"
- "when will I run out of cash?" / "cash flow forecast"
- "what retention am I owed?" / "retention tracker"
- "audit my expenses" / "find duplicate bills" / "vendor concentration"
- Any request to turn accounting data + project data into a management report pack.

Do NOT use this skill for: quantity takeoffs from drawings (use `drawing-takeoff`)
or from BIM models (use `boq-from-ifc`); estimating; anything that requires
writing back to the accounting system (the pipeline is read-only by design).

---

## Step 1 — Find the project directory

The Python scripts live in a cloned copy of the repo. In order of preference:

1. The **current working directory** — if it contains `scripts/cfo.py` and
   `scripts/schema.sql`, use it.
2. `~/autoconst-claude-construction-cfo/`
3. `~/Downloads/autoconst-claude-construction-cfo/`

If none exist, tell the user:

> *"I need the construction-cfo scripts. Clone the repo first:*
> `git clone https://github.com/hamzaabduljabbar/autoconst-claude-construction-cfo ~/autoconst-claude-construction-cfo`
> *then ask me again."*

`cd` into whichever directory has the scripts before running any commands.

---

## Step 2 — Locate the inputs

The pipeline needs **two required** inputs and takes **two optional** ones:

**Required:**
- **Accounting export** in `inputs/` — either the QuickBooks *"Transaction
  Detail by Account"* CSV (real grouped export) or a canonical CSV.
- **Project-input workbook** at `inputs/project-input.xlsx` — the contractor's
  budgets, activities, % complete, contracts, and commitments. A template lives
  at `integration-data/summit-ridge/project-input.xlsx`.

**Optional (strongly recommended for cash flow):**
- **A/R aging CSV** — open receivables (`Customer Balance Detail` export)
- **A/P aging CSV** — open payables (`Vendor Balance Detail` export)

Without A/R and A/P, the cash forecast runs in **schedule-based estimate** mode
and says so on the Assumptions sheet. With them it becomes an **invoice-based
forecast**.

If the required inputs are missing, tell the user exactly what to drop where.
Never fabricate transactions, budgets, or contract terms.

---

## Step 3 — Run the pipeline in order

Use a short project slug (e.g. the tenant's short name) as `<name>`:

```bash
python scripts/cfo.py init         --db outputs/<name>.db --tenant "<Legal Entity Name>" --force
python scripts/cfo.py ingest       --db outputs/<name>.db \
                                   --source inputs/<accounting-export>.csv \
                                   --adapter <qbo|canonical> \
                                   --projects inputs/project-input.xlsx \
                                   --control <optional-control-totals.csv>
python scripts/cfo.py load-workbook --db outputs/<name>.db --workbook inputs/project-input.xlsx
python scripts/cfo.py ingest-aging  --db outputs/<name>.db --ar inputs/ar-aging.csv --ap inputs/ap-aging.csv   # optional
python scripts/cfo.py classify      --db outputs/<name>.db          # only needed for QBO (uncoded) imports
python scripts/cfo.py review-export --db outputs/<name>.db --out outputs/<name>-review.xlsx   # if classify surfaced proposals
# (contractor fills 'decision' + optional replacement_code in the review workbook)
python scripts/cfo.py review-apply  --db outputs/<name>.db --review outputs/<name>-review.xlsx
python scripts/cfo.py jobcost       --db outputs/<name>.db --out outputs/<name>-job-cost.xlsx
python scripts/cfo.py retention     --db outputs/<name>.db --out outputs/<name>-retention.xlsx
python scripts/cfo.py cashflow      --db outputs/<name>.db --out outputs/<name>-cashflow.xlsx
python scripts/cfo.py audit         --db outputs/<name>.db --out outputs/<name>-expense-audit.xlsx
```

**Adapter choice:** `--adapter qbo` for a real QuickBooks Online grouped
"Transaction Detail by Account" export. `--adapter canonical` for the clean
canonical CSV format (documented in `docs/canonical-csv-format.md`) — any
platform that can export transactions can be mapped to this.

**Reconciliation gate:** if `--control` is provided (a control-totals CSV per
account + project), the ingest checks imported totals reconcile to the cent.
An unreconciled run is stamped as `DRAFT — NOT TRUSTED` on every report.

**Timing:** small dataset (a few hundred lines) finishes in seconds. Full
Summit Ridge fixture (563 headers / 1,244 lines) takes ~5 s end-to-end.

---

## Step 4 — Report the result honestly

After each report finishes, tell the user:

- **Job cost:** grand total actual vs approved budget, portfolio variance, and
  the projects/cost codes flagged over-budget (orange rows). If the report is
  marked DRAFT, name the reason (unreconciled or under-classified) and the
  gap in dollars.
- **Retention:** est. client retention vs sub retention, net position, and the
  count of releases flagged **CLAIMABLE NOW (overdue)** — those are the most
  urgent claims. Always call retention an *estimate* (it's derived from %
  complete, not from certified claims).
- **Cash flow:** the mode (**invoice-based forecast** or **schedule-based
  estimate**), the low balance and week per scenario, and whether a shortfall
  is projected within the horizon. If mode is schedule-based, tell the user
  they can upgrade it by dropping their A/R and A/P aging exports in `inputs/`
  and re-running `ingest-aging` + `cashflow`.
- **Expense audit:** counts of duplicate alerts, unusual-cost alerts, and
  vendor concentration flags. Always call these **review alerts, not
  confirmed errors**.

Every reported figure resolves back to source records via the report's Report
Info + Transaction Detail sheets. Do not editorialize the numbers.

---

## Absolute rules

- **Money is exact.** All amounts are integer minor units; arithmetic is
  Decimal-only. Floats are banned for money (the helper raises `TypeError`).
- **Never invent a code, an amount, or a retention figure.** The classifier
  validates every proposal against the project's allowed cost codes; missing
  data is left unclassified, not force-fit.
- **Never mark a report `TRUSTED`** unless the import reconciled AND there
  are no blocking data-quality issues AND classification coverage is 100% AND
  it's a full-period view. A date-cutoff report is analytical/DRAFT by
  definition.
- **Never auto-apply an AI cost-code guess.** Only approved deterministic rules
  auto-apply. LLM/heuristic/commitment matches are proposals that wait for
  sign-off in the review workbook.
- **Never overstate what the pipeline knows.** Retention is estimated from %
  complete (labeled "Estimated Retention Position"). Cash flow is a 13-week
  schedule-based plan or invoice-based forecast (labeled honestly). Audit
  outputs are review alerts, not confirmed errors.
- **Never write back to the accounting system.** The pipeline is read-only.
- **Never commit the contractor's real data.** `inputs/*` and `outputs/*` are
  gitignored. Sample/integration data under `integration-data/` is synthetic.
