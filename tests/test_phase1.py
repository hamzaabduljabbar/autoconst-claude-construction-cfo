"""Phase 1 acceptance tests — ingest + reconciliation gate.

Covers acceptance tests #1-7 from PROJECT-BRIEF.md §17, on both the clean
canonical fixture and the messy QBO fixture:

  #1 re-importing an overlapping export does not duplicate records
  #2 a modified transaction updates (not appends)
  #3 voided transactions represented + excluded from totals
  #4 header totals equal line totals
  #5 imported totals reconcile to the control report (both adapters)
  #6 unassigned-project costs remain visible
  #7 tax lines never mix into cost totals

Run:  python tests/test_phase1.py
"""

from __future__ import annotations

import csv
import shutil
import sqlite3
import subprocess
import sys
import tempfile
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
SAMPLE = ROOT / "sample-data"
sys.path.insert(0, str(SCRIPTS))

import cfo  # noqa: E402
from ingest import run_ingest, run_reconcile, UNASSIGNED  # noqa: E402

CANON = SAMPLE / "canonical-transactions.csv"
MESSY = SAMPLE / "qbo-transaction-detail-messy.csv"
CONTROL = SAMPLE / "control-totals.csv"
WORKBOOK = SAMPLE / "project-input.xlsx"


def _fresh_db(d: Path) -> str:
    db = d / "t.db"
    subprocess.run([sys.executable, str(SCRIPTS / "cfo.py"), "init",
                    "--db", str(db), "--tenant", "Northwind Construction Pty Ltd",
                    "--force"], check=True, capture_output=True)
    return str(db)


def _counts(db):
    conn = sqlite3.connect(db)
    h = conn.execute("SELECT COUNT(*) FROM transaction_headers").fetchone()[0]
    l = conn.execute("SELECT COUNT(*) FROM transaction_lines").fetchone()[0]
    conn.close()
    return h, l


def test_reconcile_both_adapters():          # #5
    for src, adapter in [(CANON, "canonical"), (MESSY, "qbo")]:
        with tempfile.TemporaryDirectory() as d:
            db = _fresh_db(Path(d))
            run_ingest(db, str(src), adapter, projects_workbook=str(WORKBOOK))
            ok, issues = run_reconcile(db, str(CONTROL), report=False)
            assert ok, f"{adapter} did not reconcile: {issues}"


def test_idempotent_reimport():              # #1
    for src, adapter in [(CANON, "canonical"), (MESSY, "qbo")]:
        with tempfile.TemporaryDirectory() as d:
            db = _fresh_db(Path(d))
            run_ingest(db, str(src), adapter, projects_workbook=str(WORKBOOK))
            h1, l1 = _counts(db)
            run_ingest(db, str(src), adapter, projects_workbook=str(WORKBOOK))
            h2, l2 = _counts(db)
            assert (h1, l1) == (h2, l2), f"{adapter}: re-import changed counts {(h1,l1)}->{(h2,l2)}"


def test_modified_transaction_updates():     # #2
    with tempfile.TemporaryDirectory() as d:
        db = _fresh_db(Path(d))
        run_ingest(db, str(CANON), "canonical", projects_workbook=str(WORKBOOK))
        h1, l1 = _counts(db)
        # change one amount on B-1001 (12400 -> 12900) and re-import
        modified = Path(d) / "modified.csv"
        with CANON.open(encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
            fields = rows[0].keys()
        for r in rows:
            if r["txn_id"] == "B-1001" and r["amount"] == "12400.00":
                r["amount"] = "12900.00"
        with modified.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(fields))
            w.writeheader(); w.writerows(rows)
        run_ingest(db, str(modified), "canonical", projects_workbook=str(WORKBOOK))
        h2, l2 = _counts(db)
        assert (h2, l2) == (h1, l1), "modified re-import should update in place, not append"
        conn = sqlite3.connect(db)
        amt = conn.execute("""SELECT l.amount_minor FROM transaction_lines l
            JOIN transaction_headers h ON h.id=l.header_id
            WHERE h.native_transaction_id='B-1001' AND l.amount_minor=1290000""").fetchone()
        conn.close()
        assert amt is not None, "updated amount not found after re-import"


def test_void_excluded_but_present():        # #3
    with tempfile.TemporaryDirectory() as d:
        db = _fresh_db(Path(d))
        run_ingest(db, str(CANON), "canonical", projects_workbook=str(WORKBOOK))
        conn = sqlite3.connect(db)
        voided = conn.execute(
            "SELECT COUNT(*) FROM transaction_headers WHERE is_void=1").fetchone()[0]
        conn.close()
        assert voided >= 1, "voided transaction should be present and flagged"
        ok, _ = run_reconcile(db, str(CONTROL), report=False)
        assert ok, "totals should reconcile with void excluded"


def test_header_equals_lines():              # #4
    with tempfile.TemporaryDirectory() as d:
        db = _fresh_db(Path(d))
        run_ingest(db, str(MESSY), "qbo", projects_workbook=str(WORKBOOK))
        conn = sqlite3.connect(db)
        bad = conn.execute("""
            SELECT h.native_transaction_id, h.header_total_minor, SUM(l.amount_minor)
            FROM transaction_headers h JOIN transaction_lines l ON l.header_id=h.id
            GROUP BY h.id
            HAVING h.header_total_minor != SUM(l.amount_minor)""").fetchall()
        conn.close()
        assert not bad, f"header total != sum of lines for: {bad}"


def test_unassigned_visible():               # #6
    with tempfile.TemporaryDirectory() as d:
        db = _fresh_db(Path(d))
        run_ingest(db, str(MESSY), "qbo", projects_workbook=str(WORKBOOK))
        conn = sqlite3.connect(db)
        unassigned = conn.execute("""
            SELECT COALESCE(SUM(l.amount_minor),0) FROM transaction_lines l
            JOIN transaction_headers h ON h.id=l.header_id
            WHERE l.project_id IS NULL AND l.is_tax=0 AND h.is_void=0""").fetchone()[0]
        conn.close()
        assert unassigned > 0, "company-level (unassigned-project) costs should be visible"


def test_tax_never_in_cost():                # #7
    with tempfile.TemporaryDirectory() as d:
        db = _fresh_db(Path(d))
        run_ingest(db, str(CANON), "canonical", projects_workbook=str(WORKBOOK))
        conn = sqlite3.connect(db)
        # any account total computed WITH tax would exceed the control figure
        got = conn.execute("""
            SELECT COALESCE(SUM(l.amount_minor),0) FROM transaction_lines l
            JOIN transaction_headers h ON h.id=l.header_id
            WHERE h.is_void=0 AND l.is_tax=0""").fetchone()[0]
        with_tax = conn.execute("""
            SELECT COALESCE(SUM(l.amount_minor),0) FROM transaction_lines l
            JOIN transaction_headers h ON h.id=l.header_id
            WHERE h.is_void=0""").fetchone()[0]
        conn.close()
        assert with_tax > got, "sanity: tax lines exist"
        # the reconcile (which uses is_tax=0) must pass -> tax stayed out of cost
        ok, _ = run_reconcile(db, str(CONTROL), report=False)
        assert ok


def test_unreconciled_is_flagged():          # negative control for the gate
    with tempfile.TemporaryDirectory() as d:
        db = _fresh_db(Path(d))
        run_ingest(db, str(CANON), "canonical", projects_workbook=str(WORKBOOK))
        # corrupt one control figure
        bad_control = Path(d) / "bad-control.csv"
        with CONTROL.open(encoding="utf-8") as f:
            rows = list(csv.reader(f))
        for r in rows:
            if len(r) >= 3 and r[0] == "account" and r[1] == "Materials":
                r[2] = "99999.00"
        with bad_control.open("w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerows(rows)
        ok, issues = run_reconcile(db, str(bad_control), report=False)
        assert not ok and issues, "a wrong control total must fail the gate"
        conn = sqlite3.connect(db)
        status = conn.execute("SELECT status FROM import_runs ORDER BY id DESC LIMIT 1").fetchone()[0]
        conn.close()
        assert status == "unreconciled"


def _run_all():
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"PASS {name}")


if __name__ == "__main__":
    _run_all()
    print("\nAll Phase 1 acceptance checks passed.")
