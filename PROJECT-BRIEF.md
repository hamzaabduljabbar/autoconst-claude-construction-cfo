# Construction CFO — Project Brief & Architecture

> Status: discovery complete, ready for schema design + narrow end-to-end prototype.
> This document is the technical specification. Competitive positioning lives in `PRODUCT-STRATEGY.md`.
> Two independent architecture reviews (2026-08-13) have been incorporated; see Appendix A for the change log.

---

## 1. One-line description

> A Claude-guided local financial reporting pipeline that turns a contractor's accounting export plus a structured project input workbook into a reconciled, traceable management-report pack (job cost, cash flow, retention, portfolio) in Excel.

Claude is **not** the accounting database, the calculation engine, or the source of financial truth. Claude orchestrates and explains. Python validates and calculates. SQLite preserves the record. Excel presents results and captures structured review decisions.

## 2. Responsibility boundary (the governing rule of the whole system)

| Component | Owns |
|---|---|
| **Claude (the skill)** | Discovering input files, interpreting them, asking targeted clarification questions, classifying *only* the transaction lines that deterministic rules could not resolve, and narrating results from already-approved metrics. |
| **Python (deterministic engine)** | Normalization, deduplication, reconciliation, all accounting math, cash-flow math, lineage generation, Excel generation, and enforcement of accounting policy. |
| **SQLite (source of truth)** | Persistent normalized records, import history, versioned budgets/contracts/commitments, classifications and reviews, assumptions, report runs, lineage, data-quality findings. |
| **Excel (interface out + in)** | Presenting the management pack; collecting structured classification-review decisions for round-trip import. |

**Claude may never independently decide:** whether two records are duplicates, whether an import reconciles, how actual cost is calculated, which budget revision is approved, how remaining commitment is calculated, whether a report is trusted, which cash flows are included, whether missing data is immaterial, or how money is rounded. Those are deterministic policies in Python. Claude explains them and, where a genuine business judgment is required, asks the contractor to make it.

## 3. Architecture

```
Contractor
    │
    ▼
Claude skill — workflow orchestrator
    ├── discovers & inspects input files
    ├── identifies missing information, asks targeted questions
    ├── invokes deterministic Python tools
    ├── classifies only unresolved transaction lines
    ├── explains validation & reconciliation issues
    └── narrates approved report metrics
    │
    ▼
Python application — deterministic engine
    ├── QuickBooks Online adapter + canonical-CSV adapter
    ├── normalization → transaction lines
    ├── idempotent import (dedup, updates, voids, reversals, credits)
    ├── reconciliation gate
    ├── accounting + cash-flow calculations
    ├── lineage generation
    └── Excel workbook generation (sanitized)
    │
    ▼
SQLite — persistent source of truth
    │
    ▼
Excel outputs
    ├── {date}-management-pack.xlsx
    └── {date}-classification-review.xlsx
```

Deterministic stages do the math. LLM stages (classification of leftovers, narrative) do judgment only. LLMs are unreliable calculators and good classifiers — the split is deliberate.

## 4. Data model

### 4.1 Transactions are lines, not headers

A single supplier bill can carry materials for Project A, plant hire for Project B, a company-level delivery fee, and tax — each line needs its own account / project / cost code. Classification is always at the **line** level.

```
transaction_header               transaction_line
├── source_tenant_id             ├── header_id (FK)
├── source_platform              ├── native_line_id
├── native_transaction_type      ├── account
├── native_transaction_id        ├── project           (from accounting system)
├── vendor / party               ├── cost_code         (WE ADD THIS)
├── date                         ├── tax_code
├── currency                     ├── quantity (fixed scale)
└── header_total_minor           └── amount_minor (integer minor units)
```

### 4.2 Money is integer minor units

SQLite has no arbitrary-precision decimal type, so money is stored as **integer minor units** with an explicit currency and scale:

```
amount_minor = 123456, currency = AUD, currency_scale = 2   →   AUD 1,234.56
```

Python does all arithmetic in `Decimal`, reads/writes minor units at the boundary. Binary floating-point is banned for money, quantities, percentages, unit prices. **MVP is single functional currency per contractor;** multi-currency is deferred.

### 4.3 Deduplication uses business identity, not the file hash

Every re-export has a new file hash, so the hash must not be part of the dedup key or the same transaction re-imports forever. Identity is:

```
(source_tenant_id, source_platform, native_transaction_type,
 native_transaction_id, native_line_id)
```

The file hash lives on the **import run**, not the key:

```
import_runs: source_file_hash, imported_at, source_period, adapter_version
```

The importer is idempotent and must correctly handle: changed transactions, deletions, voids, reversals, credit notes, missing native line IDs, and the same platform used by different companies. When native IDs are absent, compute a documented fingerprint and **quarantine** ambiguous matches to a review sheet — never silently merge or discard. Rejected/quarantined/duplicate rows are always counted and surfaced.

### 4.4 Project-side data (the input workbook — see §9)

```
Project
├── activities
│   ├── cost_code
│   ├── quantity + unit
│   ├── planned start / finish
│   ├── progress_method   (physical qty / weighted milestones / units /
│   │                       0-100 / 50-50 / level-of-effort / manual %)
│   └── progress_updates  (history, each with date + method + value)
├── budget_versions           (v0 original + approved revisions; never overwritten)
│   └── budget_lines per cost_code
├── contracts (head)
│   ├── contract value, payment terms, retention % + release triggers
│   ├── contract_changes       (approved variations, versioned)
│   └── pending variations      (visible, excluded from approved baseline)
└── commitments (subcontracts)
    ├── original commitment
    ├── commitment_changes      (approved, versioned)
    ├── pending changes          (visible, excluded from approved baseline)
    ├── payment terms, retention withheld
    └── claim states: certified / invoiced / paid / retained
```

**Why versioning is non-negotiable:** without versioned budgets and change orders, a cost overrun is indistinguishable from approved scope growth. Every variance is meaningless until we can say "actual vs *which* approved budget version."

## 5. The reconciliation gate

Between ingest and reporting there is a hard control:

> Imported totals must equal the totals on the **named source control report**, along the dimensions that report actually supplies. Where the source report supplies account totals, account totals must reconcile; where it supplies project totals, project totals must reconcile. Header total must equal the sum of its lines.

Each adapter declares its own reconciliation contract:

```
control_report_name, accounting_basis, date_range, tax_basis,
currency_basis, included_transaction_types,
available_control_dimensions, rounding_policy
```

We do **not** demand a project-level reconciliation when the chosen source report has no valid project control. Unassigned-project costs stay visible rather than being forced somewhere.

### 5.1 Behavior when a run does not reconcile

Not a silent pass, not a total hard stop. Graded:

- Always produce an **Import Summary** and **Data Quality** sheet.
- Diagnostic **draft** reports may still generate where calculation is possible, stamped `UNRECONCILED — DO NOT RELY ON THIS REPORT`.
- **No** executive LLM narrative from unreconciled data.
- **No** email / auto-distribution of unreconciled reports.
- The Python engine returns a warning/failure status to the Claude workflow.
- Exporting a *full* management pack from unreconciled data requires an explicit contractor override.

Claude's job here is to explain the discrepancy and guide the contractor toward fixing the source data.

## 6. Classification stack (rules first, LLM last, abstain allowed)

Cost-code classification is a stack tried in order on each unresolved **line**:

1. **Approved rule match** — contractor-approved vendor/account/project → cost-code rules. Deterministic.
2. **Learned mapping** — prior *manually approved* classifications for the same vendor+account+project. Deterministic.
3. **LLM classification** — only for lines that fell through 1–2. Returns a code (validated against the allowed list) + reason + numeric score.
4. **UNCLASSIFIED** — explicit abstain. Nothing is force-fit; these go to the review sheet as-is.

"Confidence" is the **method that produced the answer** (rule > learned > LLM > unclassified) plus per-code precision measured against the running correction history — not a bare HIGH/LOW label on LLM output. Claude can never introduce a cost code outside the allowed list; the Python validator rejects any code that isn't in `cost_codes`.

### 6.1 Corrections become rules, not just prompt examples

When a contractor overrides a classification in the review workbook, the correction is validated against SQLite and stored in `classification_reviews`. A correction that repeats for the same vendor/account/project can be **promoted to an approved `mapping_rule`** (contractor-confirmed) so it becomes deterministic next run — not merely a few-shot hint. (Fully automated rule promotion is deferred; v1 promotion is contractor-confirmed.)

## 7. The reports (management pack sheets)

All numbers are deterministic Python. Every reported value carries lineage to source records + calculation version + assumption set (§8), surfaced via detail sheets with stable drill-down IDs — **not** the claim "traces to a transaction ID," which is only true for actual cost.

### 7.1 Job Cost sheet (per project × cost code)

- **Original budget** (v0) and **Approved budget** (v0 + approved revisions + approved variations) — shown side by side.
- **Recognized actual cost** = posted supplier bills + payroll journals + approved accruals + other posted cost − credits/reversals. This is the cost figure. It is **not** "paid + invoiced-not-paid" (that double counts under accrual).
- **Cash paid** and **AP outstanding** shown as *separate* metrics, never added into actual cost.
- **Commitment waterfall** (columns that must *not* be summed together, labelled as such):
  current approved commitment · pending changes · certified to date · invoiced to date · paid to date · retention withheld · **remaining approved commitment** (= current approved − recognized actual against commitment − approved reductions).
- **% complete** from the latest progress update, with the update date and the **progress method** shown.
- **Earned value** = approved budget × % complete — valid only where progress and budget share a compatible WBS basis (enforced per activity).
- **Cost variance** = earned − recognized actual.
- **Forecast at completion** — the **method is explicit and configurable per line**: `actual + remaining budget`, `budget / CPI`, `actual / progress`, or manual ETC. The method used prints next to the number.
- **Variance alert** conditions are explicit: minimum absolute variance AND minimum budget amount AND minimum completion threshold AND percentage-variance threshold AND progress-data freshness. No vague "sample size sufficient."

### 7.2 Cash Flow sheet (weekly buckets, 90 days)

**Blocking inputs** (no trusted business-level forecast without these, though Claude may collect explicit "not applicable" confirmations conversationally):
opening cash position · forecast start date · functional currency · open AR · open AP · payroll (or explicit exclusion) · overhead (or explicit exclusion).

**Degradable inputs** (forecast proceeds, coverage noted): client-specific payment history · PO timing detail · loans (if contractor confirms none) · tax payments (if contractor confirms none in horizon).

**Scenarios:**
- **Contractual** (v1 headline) — every payment on its contractual due date.
- **Expected** — receipt date = contractual due + **median** historical days-late per debtor.
- **Conservative** — receipt date = contractual due + **80th-percentile** historical days-late; supplier/payroll payments held on-time or earlier; retention releases may slip; pending inflows probability-weighted.

Fallback hierarchy for the timing estimate: client-specific history (if sample sufficient) → contractor-wide history → configured default → contractual due date, **visibly flagged** when no estimate exists. Expected/conservative are only shown once adequate history exists; otherwise v1 shows contractual only.

The chart's uncertainty band is labelled **input coverage**, not statistical "confidence" (we have no calibrated model yet). Each week distinguishes: confirmed / probable / assumption-derived / known-but-unmodeled cash flows.

### 7.3 Retention sheet

Every project × contract: retention held by client, retention withheld from subs, the date each tranche becomes claimable, and running outstanding retention.

### 7.4 Portfolio (Dashboard) sheet

One row per project: contract value (approved), recognized actual, forecast final cost, forecast margin (loss-makers in red, sorted worst-first). Pending variations shown in a separate risk column — never folded into the approved baseline.

### 7.5 Supporting sheets

Assumptions · Data Quality · Import Summary · Transaction Detail · Lineage. All imported strings are sanitized against Excel formula injection (any cell value starting `= + - @` is prefixed with `'`).

## 8. Lineage

Requirement, stated correctly:

> Every reported value has lineage to one or more source records, a calculation version, and an assumption set.

Not everything traces to a transaction ID: budget → budget version; contract value → contract/variation; % complete → progress update; forecast → formula + assumptions; expected receipt → claim + payment assumption. Lineage is recorded per reported metric and surfaced through detail sheets with stable IDs.

## 9. The input workbook — *structured project-and-forecast input workbook*

Renamed from "small project workbook" because it is not small. It carries required + optional sheets, a workbook schema version, column definitions, controlled lists, validation rules, example rows, import diagnostics, and named ownership per input. Claude makes it manageable by walking the contractor through missing/invalid fields — but the workbook is a real data-entry surface and onboarding it is a genuine (disclosed) time cost.

## 10. Claude-guided run protocol (mandatory, ordered, never casually reordered)

1. Discover input files.
2. Identify adapter + supported export type.
3. Validate project-workbook schema version.
4. Create an import manifest.
5. Run deterministic ingestion.
6. Run deterministic reconciliation.
7. Stop for blocking data-quality issues / ask targeted questions.
8. Apply approved deterministic mapping rules.
9. Apply learned mappings.
10. Send only unresolved eligible lines to LLM classification.
11. Validate LLM response against allowed schema + cost codes.
12. Preserve unresolved lines as `UNCLASSIFIED`.
13. Calculate job-cost, retention, cash-flow metrics in Python.
14. Determine draft vs trusted report status.
15. Generate + verify both Excel workbooks.
16. Generate narrative only from reconciled, approved metrics.
17. Present a concise completion / warning / required-action summary.

## 11. MVP boundary

**Supported customer:** single legal entity · one functional currency · QuickBooks Online · accrual accounting · 3–10 active projects · project-tagged supplier bills · one standardized input workbook.

**Supported inputs:** one named QBO transaction-detail report · one named QBO reconciliation/control report · canonical-CSV alternative · project & approved-budget workbook · opening bank balance · open AR & AP · explicit recurring cash-flow assumptions.

**Supported outputs:** management pack · classification-review workbook.

**Capabilities:** Claude-guided onboarding · idempotent transaction-line import · source-total reconciliation · actual-vs-approved-budget job costing · budget & change-order history · basic approved commitments · retention tracking · contractual cash-flow scenario · classification rules + LLM fallback + abstain · visible data-quality warnings · source/calculation/assumption lineage.

**Deferred until after MVP validation:** QuickBooks MCP · Xero/MYOB/Zoho adapters · multi-currency · statistical payment-behavior forecasting · statistical confidence intervals · unattended email delivery · broad anomaly detection · complex EVM · fully automated rule promotion · expected/conservative scenarios where history is inadequate.

## 12. Scheduling

MVP is **user-initiated**: the contractor opens Claude Code and invokes `construction-cfo`. Because the full experience is an interactive Claude session, cron/Task Scheduler is *not* equivalent. Later, a deterministic `python -m construction_cfo run --no-llm` can import recognized files, apply existing approved mappings, keep leftovers `UNCLASSIFIED`, and generate reports with warnings — and only *that* gets scheduled, once idempotency, auth, file locking, failure notification, and overlapping-run prevention exist.

## 13. Security controls retained for v1

Local skill on the contractor's own machine, so this is proportionate, not SaaS-grade — but the LLM interface *raises* some of these, it doesn't remove them:

- Tell the contractor which fields are sent to Claude; minimize data in classification prompts.
- Treat imported descriptions / supplier text as **untrusted**; a supplier name must never alter the skill's workflow (prompt-injection resistance).
- Require structured LLM output, validate against allowed cost codes; never execute model-produced SQL/formulas/shell derived from source text.
- Redact sensitive data from logs; version output files rather than overwrite; email delivery off by default.
- Provide local file-permission, backup, and data-deletion guidance; keep any connector credentials out of workbooks and the DB.

## 14. Phase plan

**Phase 0 — Schema + scaffold.** Repo layout, dependencies, the MVP SQLite schema (§15), sample dataset, `init` command. Outcome: empty DB with the right shape; environment validated.

**Phase 1 — Ingest & reconcile (non-LLM vertical slice).** QBO adapter + canonical-CSV adapter, idempotent line-level import, dedup/update/void/reversal handling, the reconciliation gate. Outcome: import sample QBO export → totals reconcile to the named control report → prove idempotency on re-import.

**Phase 2 — Job cost report.** Actual-vs-approved-budget, versioned budgets, change orders, commitment waterfall, progress methods, FAC methods. Outcome: a job-cost sheet whose every number resolves to source + calc version.

**Phase 3 — Classification loop.** Rule → learned → LLM → abstain; the round-trip review workbook (§16 protocol); correction → rule promotion. Outcome: leftovers classified, corrections validated against SQLite and fed back.

**Phase 4 — Retention tracker.** Per §7.3.

**Phase 5 — Cash-flow forecast.** Blocking/degradable input gating, contractual scenario first, input-coverage band. Outcome: trusted contractual forecast or an explicit refusal listing missing blocking inputs.

**Phase 6 — Narrative.** LLM summary from reconciled, approved metrics only.

**Phase 7 — Additional adapters + scheduling + scenarios.** Only after the deterministic CLI is reliable.

The **Python application is the real product**; the Claude skill is the guided interface + orchestration layer. That keeps testing, reproducibility, and any future scheduling safe.

## 15. MVP schema (target relations, not a fixed table count)

```
source_systems      import_runs        source_records
transaction_headers transaction_lines  projects
cost_codes          activities         progress_updates
budget_versions     budget_lines       contracts
contract_changes    commitments        commitment_changes
claims              claim_lines        payments
classifications     classification_reviews  mapping_rules
forecast_runs       forecast_assumptions    report_runs
lineage             data_quality_issues
```

The objective is the minimum relations that satisfy the versioning, lineage, reconciliation, and reporting promises above — some (e.g. `claim_lines`, `payments`) may be thin stubs in the first slice but exist so the schema doesn't need breaking migrations later.

## 16. Classification-review workbook protocol

Sheet carries per row: immutable classification ID · transaction-line ID · report-run ID · original proposed code · classification method · proposed score · reviewer decision · replacement code · reviewer comment · review timestamp · workbook schema version. The importer **rejects**: unknown IDs · edited immutable IDs · duplicate decisions · invalid cost codes · incompatible workbook versions · formula-based review values · rows for archived/unknown projects. Hidden IDs improve presentation but establish nothing — all review data is validated against SQLite.

## 17. Acceptance tests (MVP is not "usable" until all pass)

1. Re-importing an overlapping QBO export does not duplicate records.
2. A modified transaction updates/versions correctly.
3. Voided/reversed/credited transactions represented correctly.
4. Header totals equal line totals.
5. Imported totals reconcile to the named QBO control report.
6. Unassigned-project costs stay visible.
7. Tax-inclusive and tax-exclusive values never mix.
8. Approved budget revisions preserve the original baseline.
9. Pending changes do not alter approved budget.
10. Payments do not double-count actual cost.
11. Remaining commitments do not overlap recognized actuals.
12. Missing critical inputs prevent a trusted cash forecast.
13. Invalid classification-review rows are rejected.
14. Claude cannot introduce an arbitrary cost code.
15. Imported text cannot become an Excel formula.
16. Unreconciled data cannot produce an authoritative narrative.
17. Every dashboard metric resolves to source records + calc version.
18. Workbook totals agree with database totals.
19. Repeating a deterministic run is reproducible.
20. Monetary values agree across CSV, SQLite, Python, Excel.
21. Injected text in a description cannot alter the skill's instructions/tool flow.
22. An interrupted Claude session resumes/restarts without corrupting the DB.

---

## Appendix A — Review change log

**Review 1 (structural):** transaction lines vs headers · versioned budgets/change orders · commitment states · reconciliation gate · Decimal/minor-units · classification stack · cash-flow input completeness · scenarios · lineage restated · dedup key · anomaly detector · formula-injection · five→two workbooks. All incorporated.

**Review 2 (consistency + accounting):** removed all old/new contradictions by rewriting the brief clean · dedup identity moved to business key with file-hash on import_runs · money as integer minor units in SQLite · actual cost = recognized (not paid+invoiced) with cash/AP separate · commitment waterfall with non-summable columns · pending≠approved baseline · expected/conservative timing defined via median / 80th-pct days-late with fallback hierarchy · chart band relabelled "input coverage" · blocking vs degradable cash inputs · progress methods per activity · explicit variance-alert conditions · classification-review protocol validated against SQLite · reconciliation contract per adapter · graded unreconciled behavior · ~24-relation MVP schema · removed 60-second / CFO-equal / QB-MCP-read-only claims · workbook renamed · Claude-guided run protocol · 22 acceptance tests · **competitive commentary moved to `PRODUCT-STRATEGY.md`**.

**Consciously deferred (documented, not silently dropped):** multi-currency · statistical confidence intervals · fully automated rule promotion · Xero/MYOB/Zoho · unattended scheduling · broad anomaly detection · full security posture beyond §13.
