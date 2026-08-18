-- Construction CFO — canonical SQLite schema (Phase 0)
-- Source of truth for one contractor (one functional currency, single legal entity in MVP).
--
-- Design rules enforced here (from PROJECT-BRIEF.md):
--   * Money is INTEGER minor units (cents). No REAL for money, ever. Scale + currency travel with it.
--   * Classification is at the transaction_line level, never the header.
--   * Dedup identity is a business key; the file hash lives on import_runs, NOT in the key.
--   * Budgets, contracts and commitments are VERSIONED — the original baseline is never overwritten.
--   * Pending changes are stored but flagged; they never mutate an approved baseline.
--   * Every reported number can reach a source record + calculation version + assumption set via `lineage`.
--
-- Apply with: sqlite3 cfo.db < schema.sql   (or `python scripts/cfo.py init`)

PRAGMA foreign_keys = ON;

-- ---------------------------------------------------------------------------
-- 0. Meta / provenance
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT
);
-- expected keys: schema_version, tenant_id, functional_currency, currency_scale, created_at

CREATE TABLE IF NOT EXISTS source_systems (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,          -- the contractor / company file this data belongs to
    platform      TEXT NOT NULL,          -- 'qbo' | 'canonical_csv' | (future) 'xero' ...
    company_name  TEXT,
    accounting_basis TEXT,                -- 'accrual' | 'cash'
    UNIQUE(tenant_id, platform)
);

CREATE TABLE IF NOT EXISTS import_runs (
    id               INTEGER PRIMARY KEY,
    source_system_id INTEGER NOT NULL,
    source_file_name TEXT,
    source_file_hash TEXT,                -- SHA-256 of the raw file. Lives HERE, not in the dedup key.
    adapter          TEXT NOT NULL,       -- adapter that parsed it
    adapter_version  TEXT NOT NULL,
    source_period    TEXT,                -- e.g. '2026-07-01..2026-07-31'
    imported_at      TEXT NOT NULL,       -- ISO-8601
    row_count_raw    INTEGER,
    row_count_kept   INTEGER,
    row_count_rejected INTEGER,
    row_count_quarantined INTEGER,
    status           TEXT,                -- 'ingested' | 'reconciled' | 'unreconciled' | 'failed'
    FOREIGN KEY(source_system_id) REFERENCES source_systems(id)
);

-- Immutable raw layer: every row as it arrived, before normalization. Never mutated.
CREATE TABLE IF NOT EXISTS source_records (
    id            INTEGER PRIMARY KEY,
    import_run_id INTEGER NOT NULL,
    row_index     INTEGER,               -- position in the source file
    raw_json      TEXT NOT NULL,         -- the original row, verbatim
    disposition   TEXT,                  -- 'kept' | 'rejected' | 'quarantined' | 'duplicate'
    disposition_reason TEXT,
    FOREIGN KEY(import_run_id) REFERENCES import_runs(id)
);

-- ---------------------------------------------------------------------------
-- 1. Reference / dimensions
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS accounts (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    native_id     TEXT,                  -- account id in the source system
    name          TEXT NOT NULL,         -- 'Materials', 'Wages', 'Plant Hire', 'Subcontract' ...
    account_type  TEXT,                  -- 'expense' | 'cogs' | 'income' | 'asset' | 'liability' ...
    UNIQUE(tenant_id, native_id)
);

CREATE TABLE IF NOT EXISTS parties (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    native_id     TEXT,
    name          TEXT NOT NULL,         -- vendor / customer / subcontractor
    role          TEXT,                  -- 'vendor' | 'customer' | 'subcontractor'
    UNIQUE(tenant_id, native_id, role)
);

CREATE TABLE IF NOT EXISTS projects (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    native_id     TEXT,                  -- project/customer:job id in the source system
    code          TEXT,                  -- short code the contractor uses
    name          TEXT NOT NULL,
    status        TEXT,                  -- 'active' | 'complete' | 'archived'
    UNIQUE(tenant_id, native_id)
);

CREATE TABLE IF NOT EXISTS cost_codes (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    code          TEXT NOT NULL,         -- e.g. '03-30-00' or contractor's own scheme
    name          TEXT NOT NULL,         -- 'Cast-in-place concrete' / 'Bulk earthworks'
    UNIQUE(tenant_id, code)
);

-- Which cost codes are valid on which project (the allowed list the classifier is bound to).
CREATE TABLE IF NOT EXISTS activities (
    id             INTEGER PRIMARY KEY,
    project_id     INTEGER NOT NULL,
    cost_code_id   INTEGER NOT NULL,
    name           TEXT,
    quantity       INTEGER,              -- fixed-scale integer (see quantity_scale)
    quantity_scale INTEGER DEFAULT 0,
    unit           TEXT,
    planned_start  TEXT,
    planned_finish TEXT,
    progress_method TEXT,                -- 'physical_qty'|'weighted_milestones'|'units'|'0_100'|'50_50'|'level_of_effort'|'manual_pct'
    FOREIGN KEY(project_id)   REFERENCES projects(id),
    FOREIGN KEY(cost_code_id) REFERENCES cost_codes(id),
    UNIQUE(project_id, cost_code_id)
);

CREATE TABLE IF NOT EXISTS progress_updates (
    id            INTEGER PRIMARY KEY,
    activity_id   INTEGER NOT NULL,
    as_of_date    TEXT NOT NULL,
    method        TEXT,                  -- method actually used for this reading
    pct_complete  INTEGER,               -- 0..10000 (basis points, i.e. scale 2 of a percent) OR
    qty_installed INTEGER,               -- fixed-scale, when method is quantity-based
    qty_scale     INTEGER DEFAULT 0,
    recorded_by   TEXT,
    FOREIGN KEY(activity_id) REFERENCES activities(id)
);

-- ---------------------------------------------------------------------------
-- 2. Transactions (header + line; classification lives on the line)
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS transaction_headers (
    id                     INTEGER PRIMARY KEY,
    import_run_id          INTEGER NOT NULL,
    tenant_id              TEXT NOT NULL,
    source_platform        TEXT NOT NULL,
    native_transaction_type TEXT NOT NULL,   -- 'Bill' | 'Expense' | 'JournalEntry' | 'CreditNote' | 'Payment' ...
    native_transaction_id  TEXT NOT NULL,
    party_id               INTEGER,
    txn_date               TEXT,
    currency               TEXT NOT NULL,
    header_total_minor     INTEGER,          -- signed, minor units
    is_void                INTEGER DEFAULT 0,
    is_reversal            INTEGER DEFAULT 0,
    content_sig            TEXT,             -- signature of header+lines; drives idempotent re-import
    supersedes_header_id   INTEGER,          -- set when this is an updated version of an earlier import
    -- Business-identity dedup key. NOTE: source_file_hash is deliberately NOT here.
    UNIQUE(tenant_id, source_platform, native_transaction_type, native_transaction_id),
    FOREIGN KEY(import_run_id) REFERENCES import_runs(id),
    FOREIGN KEY(party_id)      REFERENCES parties(id)
);

CREATE TABLE IF NOT EXISTS transaction_lines (
    id              INTEGER PRIMARY KEY,
    header_id       INTEGER NOT NULL,
    native_line_id  TEXT,                 -- may be NULL -> fingerprint path used at import
    line_fingerprint TEXT,                -- deterministic hash when native_line_id absent
    account_id      INTEGER,
    project_id      INTEGER,              -- may be NULL: unassigned-project cost (stays visible)
    cost_code_id    INTEGER,              -- NULL until classified
    tax_code        TEXT,
    quantity        INTEGER,
    quantity_scale  INTEGER DEFAULT 0,
    amount_minor    INTEGER NOT NULL,     -- signed, minor units, tax-exclusive unless tax_code marks otherwise
    is_tax          INTEGER DEFAULT 0,    -- 1 = this line is a tax amount (kept separate; never mixed into cost)
    memo            TEXT,                 -- vendor/description text — treated as UNTRUSTED downstream
    -- line identity within a header
    UNIQUE(header_id, native_line_id),
    FOREIGN KEY(header_id)    REFERENCES transaction_headers(id),
    FOREIGN KEY(account_id)   REFERENCES accounts(id),
    FOREIGN KEY(project_id)   REFERENCES projects(id),
    FOREIGN KEY(cost_code_id) REFERENCES cost_codes(id)
);

-- ---------------------------------------------------------------------------
-- 3. Versioned budgets
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS budget_versions (
    id            INTEGER PRIMARY KEY,
    project_id    INTEGER NOT NULL,
    version_no    INTEGER NOT NULL,      -- 0 = original baseline
    label         TEXT,                  -- 'Original tender' | 'Approved revision 1'
    status        TEXT,                  -- 'approved' | 'pending' | 'superseded'
    approved_by   TEXT,
    approved_date TEXT,
    created_at    TEXT,
    FOREIGN KEY(project_id) REFERENCES projects(id),
    UNIQUE(project_id, version_no)
);

CREATE TABLE IF NOT EXISTS budget_lines (
    id                INTEGER PRIMARY KEY,
    budget_version_id INTEGER NOT NULL,
    cost_code_id      INTEGER NOT NULL,
    amount_minor      INTEGER NOT NULL,
    note              TEXT,
    FOREIGN KEY(budget_version_id) REFERENCES budget_versions(id),
    FOREIGN KEY(cost_code_id)      REFERENCES cost_codes(id),
    UNIQUE(budget_version_id, cost_code_id)
);

-- ---------------------------------------------------------------------------
-- 4. Head contracts (revenue side) — versioned
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS contracts (
    id                 INTEGER PRIMARY KEY,
    project_id         INTEGER NOT NULL,
    party_id           INTEGER,           -- the client
    contract_value_minor INTEGER,         -- original value
    payment_terms_days INTEGER,           -- e.g. 30 for Net-30
    claim_frequency    TEXT,              -- 'monthly' | 'milestone'
    retention_pct_bp   INTEGER,           -- retention %, basis points (e.g. 500 = 5.00%)
    retention_release_trigger TEXT,       -- 'practical_completion' | 'defects_liability_end'
    median_days_late   INTEGER DEFAULT 0, -- client's historical payment lateness (cash-flow scenarios)
    FOREIGN KEY(project_id) REFERENCES projects(id),
    FOREIGN KEY(party_id)   REFERENCES parties(id)
);

CREATE TABLE IF NOT EXISTS contract_changes (
    id            INTEGER PRIMARY KEY,
    contract_id   INTEGER NOT NULL,
    seq_no        INTEGER NOT NULL,
    description   TEXT,
    value_minor   INTEGER NOT NULL,       -- +/- change to contract value
    status        TEXT NOT NULL,          -- 'approved' | 'pending'  (pending never alters the baseline)
    approved_by   TEXT,
    approved_date TEXT,
    FOREIGN KEY(contract_id) REFERENCES contracts(id),
    UNIQUE(contract_id, seq_no)
);

-- ---------------------------------------------------------------------------
-- 5. Commitments (subcontracts / POs, cost side) — versioned, with claim states
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS commitments (
    id                 INTEGER PRIMARY KEY,
    project_id         INTEGER NOT NULL,
    party_id           INTEGER,           -- the subcontractor / supplier
    cost_code_id       INTEGER,
    original_value_minor INTEGER NOT NULL,
    payment_terms_days INTEGER,
    retention_pct_bp   INTEGER,           -- retention we withhold FROM the sub
    FOREIGN KEY(project_id)   REFERENCES projects(id),
    FOREIGN KEY(party_id)     REFERENCES parties(id),
    FOREIGN KEY(cost_code_id) REFERENCES cost_codes(id)
);

CREATE TABLE IF NOT EXISTS commitment_changes (
    id            INTEGER PRIMARY KEY,
    commitment_id INTEGER NOT NULL,
    seq_no        INTEGER NOT NULL,
    description   TEXT,
    value_minor   INTEGER NOT NULL,
    status        TEXT NOT NULL,          -- 'approved' | 'pending'
    approved_by   TEXT,
    approved_date TEXT,
    FOREIGN KEY(commitment_id) REFERENCES commitments(id),
    UNIQUE(commitment_id, seq_no)
);

-- Claims received from subs (and, by role, claims we issue up the chain).
CREATE TABLE IF NOT EXISTS claims (
    id            INTEGER PRIMARY KEY,
    commitment_id INTEGER,               -- NULL for head-contract claims we issue
    contract_id   INTEGER,               -- set for head-contract claims (revenue)
    claim_no      INTEGER,
    claim_date    TEXT,
    state         TEXT,                  -- 'certified' | 'invoiced' | 'paid'
    gross_minor   INTEGER,
    retention_minor INTEGER,
    net_minor     INTEGER,
    due_date      TEXT,
    FOREIGN KEY(commitment_id) REFERENCES commitments(id),
    FOREIGN KEY(contract_id)   REFERENCES contracts(id)
);

CREATE TABLE IF NOT EXISTS claim_lines (
    id            INTEGER PRIMARY KEY,
    claim_id      INTEGER NOT NULL,
    cost_code_id  INTEGER,
    amount_minor  INTEGER NOT NULL,
    FOREIGN KEY(claim_id)     REFERENCES claims(id),
    FOREIGN KEY(cost_code_id) REFERENCES cost_codes(id)
);

CREATE TABLE IF NOT EXISTS payments (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    direction     TEXT NOT NULL,         -- 'in' (receipt) | 'out' (payment)
    party_id      INTEGER,
    claim_id      INTEGER,               -- links a payment to the claim it settles, when known
    paid_date     TEXT,
    amount_minor  INTEGER NOT NULL,
    FOREIGN KEY(party_id) REFERENCES parties(id),
    FOREIGN KEY(claim_id) REFERENCES claims(id)
);

-- Open accounts-receivable items (money clients owe us) from an A/R aging export.
CREATE TABLE IF NOT EXISTS ar_items (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    party_id      INTEGER,               -- customer
    project_id    INTEGER,
    doc_type      TEXT,                  -- 'Invoice' | 'Claim' ...
    doc_no        TEXT,
    issue_date    TEXT,
    due_date      TEXT,
    amount_minor  INTEGER,               -- original amount
    open_minor    INTEGER,               -- still outstanding
    import_run_id INTEGER,
    FOREIGN KEY(party_id) REFERENCES parties(id),
    FOREIGN KEY(project_id) REFERENCES projects(id),
    UNIQUE(tenant_id, doc_type, doc_no)
);

-- Open accounts-payable items (money we owe suppliers/subs) from an A/P aging export.
CREATE TABLE IF NOT EXISTS ap_items (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    party_id      INTEGER,               -- vendor / subcontractor
    project_id    INTEGER,
    doc_type      TEXT,                  -- 'Bill' ...
    doc_no        TEXT,
    issue_date    TEXT,
    due_date      TEXT,
    amount_minor  INTEGER,
    open_minor    INTEGER,
    import_run_id INTEGER,
    FOREIGN KEY(party_id) REFERENCES parties(id),
    FOREIGN KEY(project_id) REFERENCES projects(id),
    UNIQUE(tenant_id, doc_type, doc_no)
);

-- ---------------------------------------------------------------------------
-- 6. Classification + review + learned rules
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS classifications (
    id                 INTEGER PRIMARY KEY,
    transaction_line_id INTEGER NOT NULL,
    report_run_id      INTEGER,
    proposed_cost_code_id INTEGER,       -- NULL = UNCLASSIFIED (explicit abstain)
    method             TEXT NOT NULL,    -- 'rule' | 'learned' | 'llm' | 'unclassified'
    score              INTEGER,          -- 0..10000 basis points; NULL for deterministic methods
    reason             TEXT,
    created_at         TEXT,
    FOREIGN KEY(transaction_line_id)   REFERENCES transaction_lines(id),
    FOREIGN KEY(proposed_cost_code_id) REFERENCES cost_codes(id)
);

-- A single export of proposals for sign-off. Binds a workbook to THIS database
-- and tenant so a workbook from another db (with overlapping integer ids) cannot
-- modify the wrong records.
CREATE TABLE IF NOT EXISTS review_batches (
    id           INTEGER PRIMARY KEY,
    tenant_id    TEXT NOT NULL,
    db_uid       TEXT NOT NULL,          -- must equal meta.db_uid at apply time
    batch_token  TEXT NOT NULL UNIQUE,   -- opaque token stamped into the workbook
    created_at   TEXT,
    row_count    INTEGER,
    status       TEXT                    -- 'open' | 'applied'
);

CREATE TABLE IF NOT EXISTS classification_reviews (
    id                 INTEGER PRIMARY KEY,
    classification_id  INTEGER NOT NULL,
    transaction_line_id INTEGER NOT NULL,
    review_batch_id    INTEGER,
    report_run_id      INTEGER,
    original_code_id   INTEGER,
    reviewer_decision  TEXT,             -- 'accept' | 'replace'
    replacement_code_id INTEGER,
    line_hash          TEXT,             -- content hash at export (optimistic concurrency)
    reviewer_comment   TEXT,
    reviewer           TEXT,
    reviewed_at        TEXT,
    workbook_version   TEXT,
    FOREIGN KEY(classification_id)   REFERENCES classifications(id),
    FOREIGN KEY(transaction_line_id) REFERENCES transaction_lines(id),
    FOREIGN KEY(review_batch_id)     REFERENCES review_batches(id),
    -- one review per classification per batch -> re-applying a workbook is a no-op
    UNIQUE(review_batch_id, classification_id)
);

-- Deterministic mappings promoted from repeated, contractor-confirmed corrections.
CREATE TABLE IF NOT EXISTS mapping_rules (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    party_id      INTEGER,               -- vendor scope (optional)
    account_id    INTEGER,               -- account scope (optional)
    project_id    INTEGER,               -- project scope (optional)
    cost_code_id  INTEGER NOT NULL,      -- the code this rule assigns
    status        TEXT NOT NULL,         -- 'approved' (auto-applies) | 'candidate' (needs sign-off)
    support_count INTEGER DEFAULT 0,     -- consistent confirmed decisions backing this rule
    created_from_review_id INTEGER,
    created_at    TEXT,
    FOREIGN KEY(party_id)     REFERENCES parties(id),
    FOREIGN KEY(account_id)   REFERENCES accounts(id),
    FOREIGN KEY(project_id)   REFERENCES projects(id),
    FOREIGN KEY(cost_code_id) REFERENCES cost_codes(id)
);

-- ---------------------------------------------------------------------------
-- 7. Forecasts + assumptions
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS forecast_runs (
    id            INTEGER PRIMARY KEY,
    project_id    INTEGER,               -- NULL = business-level roll-up
    scenario      TEXT NOT NULL,         -- 'contractual' | 'expected' | 'conservative'
    horizon_start TEXT,
    horizon_weeks INTEGER,
    created_at    TEXT,
    FOREIGN KEY(project_id) REFERENCES projects(id)
);

CREATE TABLE IF NOT EXISTS forecast_assumptions (
    id             INTEGER PRIMARY KEY,
    forecast_run_id INTEGER NOT NULL,
    key            TEXT NOT NULL,        -- 'opening_cash_minor', 'default_days_late', 'payroll_weekly_minor' ...
    value          TEXT NOT NULL,
    coverage       TEXT,                 -- 'confirmed' | 'probable' | 'assumption' | 'unmodeled'
    FOREIGN KEY(forecast_run_id) REFERENCES forecast_runs(id)
);

-- ---------------------------------------------------------------------------
-- 8. Reporting, lineage, data quality
-- ---------------------------------------------------------------------------

CREATE TABLE IF NOT EXISTS report_runs (
    id            INTEGER PRIMARY KEY,
    tenant_id     TEXT NOT NULL,
    generated_at  TEXT NOT NULL,
    trusted       INTEGER DEFAULT 0,     -- 1 only if the run reconciled
    engine_version TEXT,
    notes         TEXT
);

-- Every reported value's provenance: source record(s) + calc version + assumption set.
CREATE TABLE IF NOT EXISTS lineage (
    id             INTEGER PRIMARY KEY,
    report_run_id  INTEGER NOT NULL,
    metric         TEXT NOT NULL,        -- 'job_cost.recognized_actual' etc.
    ref_sheet      TEXT,                 -- workbook sheet
    ref_cell       TEXT,                 -- stable drill-down id / cell
    source_type    TEXT,                 -- 'transaction_line' | 'budget_version' | 'progress_update' | 'formula' ...
    source_ids     TEXT,                 -- JSON array of source row ids
    calculation_version TEXT,
    assumption_set TEXT,                 -- JSON or forecast_run reference
    FOREIGN KEY(report_run_id) REFERENCES report_runs(id)
);

CREATE TABLE IF NOT EXISTS data_quality_issues (
    id             INTEGER PRIMARY KEY,
    import_run_id  INTEGER,
    report_run_id  INTEGER,
    severity       TEXT NOT NULL,        -- 'blocking' | 'warning' | 'info'
    category       TEXT,                 -- 'reconciliation' | 'duplicate' | 'missing_input' | 'quarantine' ...
    message        TEXT NOT NULL,
    detail_json    TEXT,
    created_at     TEXT,
    FOREIGN KEY(import_run_id) REFERENCES import_runs(id),
    FOREIGN KEY(report_run_id) REFERENCES report_runs(id)
);

-- ---------------------------------------------------------------------------
-- Indexes
-- ---------------------------------------------------------------------------

CREATE INDEX IF NOT EXISTS idx_lines_header    ON transaction_lines(header_id);
CREATE INDEX IF NOT EXISTS idx_lines_project   ON transaction_lines(project_id);
CREATE INDEX IF NOT EXISTS idx_lines_costcode  ON transaction_lines(cost_code_id);
CREATE INDEX IF NOT EXISTS idx_lines_account   ON transaction_lines(account_id);
CREATE INDEX IF NOT EXISTS idx_headers_date    ON transaction_headers(txn_date);
CREATE INDEX IF NOT EXISTS idx_class_line      ON classifications(transaction_line_id);
CREATE INDEX IF NOT EXISTS idx_progress_act    ON progress_updates(activity_id);
CREATE INDEX IF NOT EXISTS idx_dq_severity     ON data_quality_issues(severity);
CREATE INDEX IF NOT EXISTS idx_lineage_run     ON lineage(report_run_id);
