"""Phase 2 acceptance tests — job cost report against Summit Ridge.

Covers brief acceptance items #8 (approved revisions preserve baseline),
#17/#18 (metrics resolve + workbook agrees with DB), plus the earned-value /
variance / FAC math and that planted overrun scenarios actually surface.

Run:  python tests/test_phase2.py
"""

from __future__ import annotations

import subprocess
import sqlite3
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
SR = ROOT / "integration-data" / "summit-ridge"
sys.path.insert(0, str(SCRIPTS))

from ingest import run_ingest, run_reconcile          # noqa: E402
from load_workbook import load_workbook_into_db        # noqa: E402
from jobcost import compute_rows, build_job_cost_report  # noqa: E402


def _prepared_db(d: Path, reconcile=True) -> str:
    db = str(d / "sr.db")
    subprocess.run([sys.executable, str(SCRIPTS / "cfo.py"), "init", "--db", db,
                    "--tenant", "Summit Ridge Construction Pty Ltd", "--force"],
                   check=True, capture_output=True)
    run_ingest(db, str(SR / "canonical" / "canonical-transactions.csv"), "canonical",
               projects_workbook=str(SR / "project-input.xlsx"))
    if reconcile:
        run_reconcile(db, str(SR / "qbo-exports" / "control-totals.csv"), report=False)
    load_workbook_into_db(db, str(SR / "project-input.xlsx"))
    return db


def test_earned_value_and_variance_math():
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d))
        conn = sqlite3.connect(db)
        for proj in compute_rows(conn):
            for r in proj["rows"]:
                pct = r["pct_bp"] or 0
                assert r["earned"] == r["approved"] * pct // 10000, "earned = approved * pct"
                assert r["variance"] == r["earned"] - r["actual"], "variance = earned - actual"
                assert r["fac"] == r["actual"] + max(r["approved"] - r["earned"], 0), "FAC formula"
        conn.close()


def test_labour_overrun_surfaces():
    """MCF (PRJ-002) has a planted labour overrun on Site Services -> must alert."""
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d))
        conn = sqlite3.connect(db)
        mcf = [p for p in compute_rows(conn) if p["code"] == "MCF"][0]
        site = [r for r in mcf["rows"] if r["cost_code"] == "01-50-00"][0]
        conn.close()
        assert site["alert"], "Site Services overrun should raise an alert"
        assert site["actual"] > site["approved"], "actual must exceed approved"
        assert site["variance"] < -100_000_00, "the overrun should be large and negative"


def test_all_planted_overrun_projects_alert():
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d))
        conn = sqlite3.connect(db)
        by_code = {p["code"]: p for p in compute_rows(conn)}
        conn.close()
        # PRJ-002 (MCF) loss pressure and PRJ-005 (ISL) materials overrun
        for code in ("MCF", "ISL"):
            assert any(r["alert"] for r in by_code[code]["rows"]), f"{code} should have an alert"


def test_workbook_totals_agree_with_db():
    """Sum of job-cost actuals == sum of project+cost_code-tagged cost lines in the DB."""
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d))
        conn = sqlite3.connect(db)
        report_actual = sum(r["actual"] for p in compute_rows(conn) for r in p["rows"])
        db_actual = conn.execute("""
            SELECT COALESCE(SUM(l.amount_minor),0) FROM transaction_lines l
            JOIN transaction_headers h ON h.id=l.header_id
            WHERE l.project_id IS NOT NULL AND l.cost_code_id IS NOT NULL
              AND l.is_tax=0 AND h.is_void=0""").fetchone()[0]
        conn.close()
        assert report_actual == db_actual, f"{report_actual} != {db_actual}"


def test_trust_requires_reconciled_and_full_coverage():
    # Summit Ridge full period IS reconciled but has ~$2k unclassified cost,
    # so it must NOT be trusted (coverage < 100%).
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d), reconcile=True)
        s = build_job_cost_report(db, str(Path(d) / "sr.xlsx"))
        assert s["reconciled"] is True
        assert s["unclassified"] > 0
        assert s["trusted"] is False
    # unreconciled -> untrusted
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d), reconcile=False)
        s = build_job_cost_report(db, str(Path(d) / "bad.xlsx"))
        assert s["trusted"] is False


def test_trust_true_when_fully_classified():
    """Once the uncoded lines are coded, a reconciled full-period report is trusted."""
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d), reconcile=True)
        conn = sqlite3.connect(db)
        any_code = conn.execute("SELECT id FROM cost_codes LIMIT 1").fetchone()[0]
        conn.execute("UPDATE transaction_lines SET cost_code_id=? "
                     "WHERE project_id IS NOT NULL AND cost_code_id IS NULL AND is_tax=0",
                     (any_code,))
        conn.commit(); conn.close()
        s = build_job_cost_report(db, str(Path(d) / "ok.xlsx"))
        assert s["unclassified"] == 0
        assert s["trusted"] is True
        # a date-cutoff view of the same data is analytical -> never "trusted"
        s2 = build_job_cost_report(db, str(Path(d) / "cut.xlsx"), as_of="2025-12-31")
        assert s2["trusted"] is False


def test_reporting_cutoff_filters_transactions():
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d))
        full = build_job_cost_report(db, str(Path(d) / "full.xlsx"))
        early = build_job_cost_report(db, str(Path(d) / "early.xlsx"), as_of="2025-12-31")
        assert early["grand"]["actual"] < full["grand"]["actual"], "cutoff must exclude later cost"


def test_workbook_load_is_idempotent():
    with tempfile.TemporaryDirectory() as d:
        db = _prepared_db(Path(d))
        load_workbook_into_db(db, str(SR / "project-input.xlsx"))  # load a 2nd time
        conn = sqlite3.connect(db)
        counts = {t: conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                  for t in ("progress_updates", "contracts", "contract_changes",
                            "commitments", "budget_versions", "budget_lines")}
        conn.close()
        assert counts == {"progress_updates": 66, "contracts": 6, "contract_changes": 3,
                          "commitments": 36, "budget_versions": 6, "budget_lines": 66}, counts


def test_versioned_budget_selection():
    """Real versioned-budget behavior: approved v1 becomes current, v0 preserved,
    pending v2 excluded. Built as an explicit fixture (Summit Ridge has only v0)."""
    with tempfile.TemporaryDirectory() as d:
        db = str(Path(d) / "v.db")
        subprocess.run([sys.executable, str(SCRIPTS / "cfo.py"), "init", "--db", db,
                        "--tenant", "T", "--force"], check=True, capture_output=True)
        conn = sqlite3.connect(db)
        pid = conn.execute("INSERT INTO projects(tenant_id,native_id,code,name,status) "
                           "VALUES ('T','P1','P1','P1','active')").lastrowid
        ccid = conn.execute("INSERT INTO cost_codes(tenant_id,code,name) "
                            "VALUES ('T','03-30-00','Concrete')").lastrowid
        for ver, status, amt, adate in [(0, "approved", 100_000_00, None),
                                        (1, "approved", 120_000_00, "2026-02-01"),
                                        (2, "pending", 150_000_00, "2026-03-01")]:
            bvid = conn.execute("INSERT INTO budget_versions(project_id,version_no,label,status,approved_date) "
                                "VALUES (?,?,?,?,?)", (pid, ver, f"v{ver}", status, adate)).lastrowid
            conn.execute("INSERT INTO budget_lines(budget_version_id,cost_code_id,amount_minor) "
                         "VALUES (?,?,?)", (bvid, ccid, amt))
        conn.commit()
        from jobcost import _approved_budget, _original_budget
        assert _original_budget(conn, pid, ccid) == 100_000_00, "v0 baseline preserved"
        assert _approved_budget(conn, pid, ccid, None) == 120_000_00, "approved v1 is current"
        # pending v2 must NOT be selected even though its amount is higher
        assert _approved_budget(conn, pid, ccid, None) != 150_000_00, "pending v2 excluded"
        # cutoff BEFORE v1 approval -> falls back to v0 baseline
        assert _approved_budget(conn, pid, ccid, "2026-01-01") == 100_000_00, "cutoff excludes later revision"
        conn.close()


def _run_all():
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"PASS {name}")


if __name__ == "__main__":
    _run_all()
    print("\nAll Phase 2 acceptance checks passed.")
