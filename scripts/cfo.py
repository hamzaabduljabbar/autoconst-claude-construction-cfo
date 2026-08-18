"""Construction CFO — command-line entry point.

Phase 0 provides:
    init    create the SQLite database from schema.sql and stamp meta
    check   verify the database has the expected shape (all tables + meta)

Later phases add: ingest, reconcile, jobcost, classify, retention, cashflow, report.

Usage:
    python scripts/cfo.py init  --db outputs/demo.db --tenant "Demo Builders Pty Ltd"
    python scripts/cfo.py check --db outputs/demo.db
"""

from __future__ import annotations

import argparse
import datetime as _dt
import sqlite3
import uuid
from pathlib import Path

SCHEMA_VERSION = "0.1.0"
SCHEMA_PATH = Path(__file__).with_name("schema.sql")

# The full set of relations the MVP schema promises (PROJECT-BRIEF.md §15).
EXPECTED_TABLES = {
    "meta", "source_systems", "import_runs", "source_records",
    "accounts", "parties", "projects", "cost_codes", "activities",
    "progress_updates", "transaction_headers", "transaction_lines",
    "budget_versions", "budget_lines", "contracts", "contract_changes",
    "commitments", "commitment_changes", "claims", "claim_lines", "payments",
    "classifications", "classification_reviews", "mapping_rules", "review_batches",
    "forecast_runs", "forecast_assumptions", "report_runs",
    "lineage", "data_quality_issues", "ar_items", "ap_items",
}


def _now() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")


def _connect(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def cmd_init(args: argparse.Namespace) -> int:
    db_path = Path(args.db)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    if db_path.exists() and not args.force:
        print(f"[init] {db_path} already exists — pass --force to recreate.")
        return 1
    if db_path.exists():
        db_path.unlink()

    schema = SCHEMA_PATH.read_text(encoding="utf-8")
    conn = _connect(str(db_path))
    conn.executescript(schema)

    meta = {
        "schema_version": SCHEMA_VERSION,
        "tenant_id": args.tenant,
        "functional_currency": args.currency,
        "currency_scale": str(args.scale),
        "db_uid": uuid.uuid4().hex,          # stable identity binding review workbooks to this db
        "created_at": _now(),
    }
    conn.executemany(
        "INSERT OR REPLACE INTO meta(key, value) VALUES (?, ?)",
        list(meta.items()),
    )
    conn.commit()
    conn.close()

    print(f"[init] created {db_path}")
    print(f"[init] schema_version={SCHEMA_VERSION} tenant='{args.tenant}' "
          f"currency={args.currency} scale={args.scale}")
    return 0


def cmd_check(args: argparse.Namespace) -> int:
    db_path = Path(args.db)
    if not db_path.exists():
        print(f"[check] no database at {db_path}")
        return 1

    conn = _connect(str(db_path))
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    ).fetchall()
    present = {r[0] for r in rows}
    missing = EXPECTED_TABLES - present
    extra = present - EXPECTED_TABLES

    meta = dict(conn.execute("SELECT key, value FROM meta").fetchall())
    conn.close()

    print(f"[check] {db_path}")
    print(f"[check] tables present: {len(present)}  expected: {len(EXPECTED_TABLES)}")
    if missing:
        print(f"[check] MISSING: {sorted(missing)}")
    if extra:
        print(f"[check] extra (non-schema): {sorted(extra)}")
    print(f"[check] meta: schema_version={meta.get('schema_version')} "
          f"tenant='{meta.get('tenant_id')}' "
          f"currency={meta.get('functional_currency')} "
          f"scale={meta.get('currency_scale')}")

    ok = not missing and meta.get("schema_version") == SCHEMA_VERSION
    print("[check] OK" if ok else "[check] FAILED")
    return 0 if ok else 1


def main() -> int:
    p = argparse.ArgumentParser(prog="cfo", description="Construction CFO CLI")
    sub = p.add_subparsers(dest="cmd", required=True)

    pi = sub.add_parser("init", help="create the database from schema.sql")
    pi.add_argument("--db", required=True)
    pi.add_argument("--tenant", default="Demo Builders Pty Ltd")
    pi.add_argument("--currency", default="AUD")
    pi.add_argument("--scale", type=int, default=2)
    pi.add_argument("--force", action="store_true")
    pi.set_defaults(func=cmd_init)

    pc = sub.add_parser("check", help="verify the database shape")
    pc.add_argument("--db", required=True)
    pc.set_defaults(func=cmd_check)

    pg = sub.add_parser("ingest", help="import a source export into the database")
    pg.add_argument("--db", required=True)
    pg.add_argument("--source", required=True, help="path to the export file")
    pg.add_argument("--adapter", required=True, choices=["canonical", "qbo"])
    pg.add_argument("--projects", help="path to the project-input workbook (for project resolution)")
    pg.add_argument("--platform", default="qbo", help="source platform tag")
    pg.add_argument("--control", help="control-totals CSV; if given, reconcile after import")
    pg.set_defaults(func=cmd_ingest)

    pr = sub.add_parser("reconcile", help="reconcile imported totals against control totals")
    pr.add_argument("--db", required=True)
    pr.add_argument("--control", required=True)
    pr.set_defaults(func=cmd_reconcile)

    pw = sub.add_parser("load-workbook", help="load project-side data from the input workbook")
    pw.add_argument("--db", required=True)
    pw.add_argument("--workbook", required=True)
    pw.set_defaults(func=cmd_load_workbook)

    pj = sub.add_parser("jobcost", help="build the job cost report (Excel)")
    pj.add_argument("--db", required=True)
    pj.add_argument("--out", required=True)
    pj.add_argument("--as-of", dest="as_of", help="reporting cutoff date (YYYY-MM-DD)")
    pj.set_defaults(func=cmd_jobcost)

    pcl = sub.add_parser("classify", help="classify uncoded cost lines (rule/learned/commitment/heuristic)")
    pcl.add_argument("--db", required=True)
    pcl.set_defaults(func=cmd_classify)

    pre = sub.add_parser("review-export", help="write the classification review workbook")
    pre.add_argument("--db", required=True)
    pre.add_argument("--out", required=True)
    pre.set_defaults(func=cmd_review_export)

    pra = sub.add_parser("review-apply", help="apply an edited classification review workbook")
    pra.add_argument("--db", required=True)
    pra.add_argument("--review", required=True)
    pra.set_defaults(func=cmd_review_apply)

    prt = sub.add_parser("retention", help="build the retention tracker (Excel)")
    prt.add_argument("--db", required=True)
    prt.add_argument("--out", required=True)
    prt.set_defaults(func=cmd_retention)

    pag = sub.add_parser("ingest-aging", help="load A/R and A/P aging exports (open invoices)")
    pag.add_argument("--db", required=True)
    pag.add_argument("--ar", help="A/R aging CSV (open receivables)")
    pag.add_argument("--ap", help="A/P aging CSV (open payables)")
    pag.set_defaults(func=cmd_ingest_aging)

    pcf = sub.add_parser("cashflow", help="build the cash flow forecast (Excel)")
    pcf.add_argument("--db", required=True)
    pcf.add_argument("--out", required=True)
    pcf.set_defaults(func=cmd_cashflow)

    pau = sub.add_parser("audit", help="build the expense leak audit (Excel)")
    pau.add_argument("--db", required=True)
    pau.add_argument("--out", required=True)
    pau.set_defaults(func=cmd_audit)

    args = p.parse_args()
    return args.func(args)


def cmd_load_workbook(args: argparse.Namespace) -> int:
    from load_workbook import load_workbook_into_db
    load_workbook_into_db(args.db, args.workbook)
    return 0


def cmd_jobcost(args: argparse.Namespace) -> int:
    from jobcost import build_job_cost_report
    build_job_cost_report(args.db, args.out, as_of=args.as_of)
    return 0


def cmd_classify(args: argparse.Namespace) -> int:
    from classify import classify_db
    classify_db(args.db)
    return 0


def cmd_review_export(args: argparse.Namespace) -> int:
    from classify import write_review_workbook
    write_review_workbook(args.db, args.out)
    return 0


def cmd_review_apply(args: argparse.Namespace) -> int:
    from classify import apply_review
    apply_review(args.db, args.review)
    return 0


def cmd_retention(args: argparse.Namespace) -> int:
    from retention import build_retention_report
    build_retention_report(args.db, args.out)
    return 0


def cmd_ingest_aging(args: argparse.Namespace) -> int:
    from aging import ingest_aging
    ingest_aging(args.db, ar_path=args.ar, ap_path=args.ap)
    return 0


def cmd_cashflow(args: argparse.Namespace) -> int:
    from cashflow import build_cashflow_forecast
    build_cashflow_forecast(args.db, args.out)
    return 0


def cmd_audit(args: argparse.Namespace) -> int:
    from audit import build_expense_audit
    build_expense_audit(args.db, args.out)
    return 0


def cmd_ingest(args: argparse.Namespace) -> int:
    from ingest import run_ingest, run_reconcile
    run_id = run_ingest(args.db, args.source, args.adapter,
                        projects_workbook=args.projects, source_platform=args.platform)
    if args.control:
        ok, _ = run_reconcile(args.db, args.control)
        return 0 if ok else 2
    return 0


def cmd_reconcile(args: argparse.Namespace) -> int:
    from ingest import run_reconcile
    ok, _ = run_reconcile(args.db, args.control)
    return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
