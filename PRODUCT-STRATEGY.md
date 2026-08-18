# Construction CFO — Product Strategy & Positioning

> Separated from `PROJECT-BRIEF.md` on review recommendation: the technical brief stays focused on requirements, constraints, accounting policy, and architecture. Competitive and go-to-market thinking lives here.

## 1. The opportunity

Small-to-mid construction contractors already keep detailed accounting records because tax law forces them to. That data sits underused. The opportunity is to layer project-level intelligence on top of it and give a contractor who can't afford a finance hire the visibility a larger firm's controller provides — job costing, cash-flow forecasting, retention tracking, portfolio margin.

The concept is validated by the market: this is textbook construction finance (earned value, commitment control, cash-flow forecasting, retention are decades-old disciplines), and there's active demand — competitors are teaching it and charging for it.

## 2. Competitive landscape

**Concept popularizer (the "Tim" video):** demonstrates the architecture — accounting system + project controls data + Claude workflows — as a lead magnet for a paid community of ~95 construction AI workflows.

- **What it proves:** the demand is real, the architecture is sound, and contractors are willing to pay for this.
- **Where it stops:** the demo runs on mock data with the LLM reasoning freely over raw JSON. No reconciliation, no transaction lines, no versioned budgets or change orders, no lineage, no confidence discipline, no correction loop. The cash "forecast" has no opening balance or AR/AP. It's a concept film, not a tool — the actual build is left to the buyer.

**Incumbent software:** accounting platforms (QuickBooks, Xero) and construction-management/ERP products already provide *parts* of job costing, budgeting, reporting, and scheduled exports. They exist and they work.

**Honest differentiation** — not "nobody else does this," but:

| Axis | Our edge |
|---|---|
| Affordability | No ERP license, no finance hire; runs locally as a Claude skill. |
| Cross-platform normalization | Canonical model over exports, not lock-in to one vendor's reporting. |
| Workflow simplicity | Claude guides intake, missing-data questions, classification, and explanation conversationally. |
| Transparency | Every number has visible lineage to source + calculation + assumptions; unreconciled data is flagged, never dressed up. |

## 3. Why we beat the concept-video approach

The gap between *demo* and *product* is exactly the punch-list the independent reviews produced. Closing it is the entire moat:

- **Deterministic spine** — normalized SQLite + Python math, replayable and reconciled, vs live LLM reasoning over JSON each run.
- **Accounting correctness** — transaction lines, versioned budgets, change orders, commitment waterfalls, recognized-vs-cash actual cost. The things that make the difference between "looks authoritative" and "is correct."
- **Trust discipline** — reconciliation gate, HIGH/LOW replaced with a rules→learned→LLM→abstain method stack, lineage on every metric.
- **Learning loop** — contractor corrections validated against the DB and promoted to deterministic rules, so it improves per-contractor instead of re-guessing.
- **No lock-in** — standalone skill, works offline for the deterministic pipeline, not gated behind a paid community.

The danger to avoid is *also* shipping a demo — a polished workbook that looks authoritative while resting on incomplete commitments, inconsistent progress, missing change orders, and unrealistic payment timing. That is worse than an obvious failure. The brief's reconciliation gate and disclosed-assumptions posture exist specifically so we never ship that.

## 4. Positioning statement

> A locally generated, reconciled, traceable construction management reporting pack — job cost, cash flow, retention, portfolio — produced through a Claude-guided workflow, with clearly disclosed assumptions.

We do **not** position as: a virtual CFO replacement, universal accounting compatibility, or guaranteed multi-week cash forewarning. Those are unearned until demonstrated on multiple real contractors and real exports.

## 5. Target customer (initial)

Single-entity contractors, one functional currency, QuickBooks Online, accrual accounting, 3–10 active projects, project-tagged bills. Narrow on purpose — a $500k builder and a $20M multi-entity firm have radically different processes, and the larger one may already run an ERP. Win the narrow profile first, expand adapters and scenarios after validation.

## 6. Fit with the existing skill portfolio

This sits alongside the repo's other construction skills (drawing-takeoff, boq-from-ifc, cad-data-extract, 4d-scheduler) and reuses their proven patterns: SQLite spine, Excel output, confidence flagging, non-vision deterministic stacks. Same house style, same trust discipline — a coherent product line for construction contractors, not a one-off.
