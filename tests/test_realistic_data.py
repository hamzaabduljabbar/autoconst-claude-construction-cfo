"""Integration checks for the Summit Ridge realistic synthetic dataset."""
from __future__ import annotations
import json, sqlite3, subprocess, sys, tempfile
from pathlib import Path

ROOT=Path(__file__).resolve().parent.parent; SCRIPTS=ROOT/"scripts"; DATA=ROOT/"integration-data"/"summit-ridge"
sys.path.insert(0,str(SCRIPTS))
from ingest import run_ingest, run_reconcile

def fresh(d):
    db=Path(d)/"realistic.db"
    subprocess.run([sys.executable,str(SCRIPTS/"cfo.py"),"init","--db",str(db),"--tenant","Summit Ridge Construction Pty Ltd","--force"],check=True,capture_output=True)
    return str(db)

def ensure_data():
    if not (DATA/"expected"/"expected-results.json").exists(): subprocess.run([sys.executable,str(SCRIPTS/"make_realistic_data.py")],check=True)

def test_realistic_canonical_and_qbo_reconcile():
    ensure_data(); wb=DATA/"project-input.xlsx"; control=DATA/"qbo-exports"/"control-totals.csv"
    for src,adapter in [(DATA/"canonical"/"canonical-transactions.csv","canonical"),(DATA/"qbo-exports"/"qbo-transaction-detail-current.csv","qbo")]:
        with tempfile.TemporaryDirectory() as d:
            db=fresh(d); run_ingest(db,str(src),adapter,projects_workbook=str(wb)); ok,issues=run_reconcile(db,str(control),report=False); assert ok,issues

def test_overlapping_exports_are_idempotent():
    ensure_data(); wb=DATA/"project-input.xlsx"; files=sorted((DATA/"canonical").glob("canonical-through-*.csv"))
    with tempfile.TemporaryDirectory() as d:
        db=fresh(d)
        for f in files: run_ingest(db,str(f),"canonical",projects_workbook=str(wb))
        conn=sqlite3.connect(db); got=conn.execute("select count(*) from transaction_headers").fetchone()[0]; conn.close()
        expected=json.loads((DATA/"expected"/"expected-results.json").read_text())["transaction_headers"]
        assert got==expected,(got,expected)

def test_known_edge_cases_present():
    ensure_data(); rows=(DATA/"canonical"/"canonical-transactions.csv").read_text(encoding="utf-8")
    for value in ("B-VOID-01","CN-9001","B-SPLIT-01","B-SEC-01"): assert value in rows

if __name__=="__main__":
    for n,f in sorted(globals().items()):
        if n.startswith("test_") and callable(f): f(); print("PASS",n)
