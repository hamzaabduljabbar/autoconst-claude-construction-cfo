"""Phase 4 + 5 acceptance tests — retention tracker and cash flow forecast.

Run:  python tests/test_phase45.py
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

from ingest import run_ingest, run_reconcile        # noqa: E402
from load_workbook import load_workbook_into_db      # noqa: E402
from retention import compute_retention              # noqa: E402
from cashflow import compute_cashflow, HORIZON_WEEKS  # noqa: E402
from aging import ingest_aging                        # noqa: E402


def _prepared(d: Path) -> str:
    db = str(d / "sr.db")
    subprocess.run([sys.executable, str(SCRIPTS / "cfo.py"), "init", "--db", db,
                    "--tenant", "Summit Ridge Construction Pty Ltd", "--force"],
                   check=True, capture_output=True)
    run_ingest(db, str(SR / "canonical" / "canonical-transactions.csv"), "canonical",
               projects_workbook=str(SR / "project-input.xlsx"))
    run_reconcile(db, str(SR / "qbo-exports" / "control-totals.csv"), report=False)
    load_workbook_into_db(db, str(SR / "project-input.xlsx"))
    return db


def test_retention_math_and_net():
    with tempfile.TemporaryDirectory() as d:
        conn = sqlite3.connect(_prepared(Path(d)))
        rows = compute_retention(conn)
        conn.close()
        assert len(rows) == 6
        for r in rows:
            assert r["client_retention"] >= 0 and r["sub_retention"] >= 0
            assert r["net"] == r["client_retention"] - r["sub_retention"]
            # client retention = retention% of revenue certified; never exceeds it
            assert r["client_retention"] <= r["revenue_certified"]
        # a completed project at ~100% should have retention close to pct x contract
        whx = next(r for r in rows if r["code"] == "WHX")
        assert whx["pct_bp"] >= 9000, "WHX is a completed project"
        assert whx["client_retention"] > 0


def test_retention_release_dates_present():
    with tempfile.TemporaryDirectory() as d:
        conn = sqlite3.connect(_prepared(Path(d)))
        rows = compute_retention(conn)
        conn.close()
        for r in rows:
            assert r["release"], f"{r['code']} should have a release date"
            assert r["trigger"] in ("practical_completion", "defects_liability_end")


def test_cashflow_running_balance_is_consistent():
    with tempfile.TemporaryDirectory() as d:
        conn = sqlite3.connect(_prepared(Path(d)))
        cf = compute_cashflow(conn)
        conn.close()
        assert not cf["missing"], f"Summit Ridge has all blocking inputs: {cf['missing']}"
        for s, sc in cf["scenarios"].items():
            assert len(sc["rows"]) == HORIZON_WEEKS
            bal = cf["opening"]
            for row in sc["rows"]:
                bal += row["net"]
                assert row["balance"] == bal, f"{s}: running balance must equal opening + Σ net"


def test_scenarios_differ_when_client_is_late():
    with tempfile.TemporaryDirectory() as d:
        conn = sqlite3.connect(_prepared(Path(d)))
        cf = compute_cashflow(conn)
        conn.close()
        # with positive median_days_late, expected receipts land later than contractual,
        # so at some week the expected balance is below contractual
        contractual = [r["balance"] for r in cf["scenarios"]["contractual"]["rows"]]
        expected = [r["balance"] for r in cf["scenarios"]["expected"]["rows"]]
        assert expected != contractual, "late-paying clients should shift the expected curve"


def test_overdue_retention_is_flagged():
    """Release dates in the past must surface as CLAIMABLE NOW, not be hidden."""
    with tempfile.TemporaryDirectory() as d:
        conn = sqlite3.connect(_prepared(Path(d)))
        rows = compute_retention(conn, as_of="2026-08-14")
        conn.close()
        overdue = [r["code"] for r in rows if r["claim_status"].startswith("CLAIMABLE NOW")]
        # APR, ISL, MCF, WHX all have release dates before 2026-08-14
        for code in ("APR", "ISL", "MCF", "WHX"):
            assert code in overdue, f"{code} is overdue and must be flagged; got {overdue}"


def test_sub_retention_not_double_counted():
    """Sub retention is deduped by vendor: it can never exceed the sub's retention
    rate applied to that vendor's total cost more than once."""
    with tempfile.TemporaryDirectory() as d:
        db = _prepared(Path(d))
        conn = sqlite3.connect(db)
        rows = compute_retention(conn)
        # independent recompute: retention% (max per vendor) × vendor cost, once each
        indep_total = 0
        for pid, in conn.execute("SELECT id FROM projects"):
            by_party = {}
            for party_id, rbp in conn.execute(
                    "SELECT party_id, retention_pct_bp FROM commitments WHERE project_id=?", (pid,)):
                if party_id and rbp:
                    by_party[party_id] = max(by_party.get(party_id, 0), rbp)
            for party_id, rbp in by_party.items():
                cost = conn.execute("""SELECT COALESCE(SUM(l.amount_minor),0)
                    FROM transaction_lines l JOIN transaction_headers h ON h.id=l.header_id
                    WHERE l.project_id=? AND h.party_id=? AND l.is_tax=0 AND h.is_void=0""",
                    (pid, party_id)).fetchone()[0]
                indep_total += cost * rbp // 10000
        conn.close()
        assert sum(r["sub_retention"] for r in rows) == indep_total


def test_scenarios_are_monotonic_every_week():
    """conservative <= expected <= contractual at EVERY week (delayed receipts can
    never improve the cash position)."""
    with tempfile.TemporaryDirectory() as d:
        conn = sqlite3.connect(_prepared(Path(d)))
        cf = compute_cashflow(conn)
        conn.close()
        C = cf["scenarios"]["contractual"]["rows"]
        E = cf["scenarios"]["expected"]["rows"]
        V = cf["scenarios"]["conservative"]["rows"]
        for i in range(HORIZON_WEEKS):
            assert C[i]["balance"] >= E[i]["balance"] >= V[i]["balance"], \
                f"week {i}: scenario ordering violated"


def test_cashflow_flags_the_shortfall():
    with tempfile.TemporaryDirectory() as d:
        conn = sqlite3.connect(_prepared(Path(d)))
        cf = compute_cashflow(conn)
        conn.close()
        exp = cf["scenarios"]["expected"]
        # Summit Ridge is winding down (4 of 6 jobs complete) with full overheads ->
        # a genuine cash shortfall must be detected within the horizon
        assert exp["first_shortfall"] is not None, "the forecast should catch the shortfall"
        assert exp["min_balance"] < cf["opening"], "balance should fall below opening"


def test_aging_switches_to_invoice_based_forecast():
    """Loading open A/R + A/P turns the schedule-based ESTIMATE into an invoice-
    based forecast, and the two give materially different answers."""
    with tempfile.TemporaryDirectory() as d:
        db = _prepared(Path(d))
        conn = sqlite3.connect(db)
        est = compute_cashflow(conn)
        assert est["mode"] == "schedule-based"
        conn.close()
        # ingest real open invoices
        ing = ingest_aging(db, ar_path=str(SR / "qbo-exports" / "ar-aging.csv"),
                           ap_path=str(SR / "qbo-exports" / "ap-aging.csv"))
        assert ing["ar"] == 8 and ing["ap"] == 8
        conn = sqlite3.connect(db)
        real = compute_cashflow(conn)
        conn.close()
        assert real["mode"] == "invoice-based"
        # the real forecast must still be internally consistent + monotonic
        for s in ("contractual", "expected", "conservative"):
            bal = real["opening"]
            for row in real["scenarios"][s]["rows"]:
                bal += row["net"]
                assert row["balance"] == bal
        C = real["scenarios"]["contractual"]["rows"]
        V = real["scenarios"]["conservative"]["rows"]
        for i in range(HORIZON_WEEKS):
            assert C[i]["balance"] >= V[i]["balance"]
        # and it should differ from the estimate (real open AR changes the picture)
        assert real["scenarios"]["expected"]["min_balance"] != est["scenarios"]["expected"]["min_balance"]


def test_aging_ingest_is_idempotent():
    with tempfile.TemporaryDirectory() as d:
        db = _prepared(Path(d))
        ingest_aging(db, ar_path=str(SR / "qbo-exports" / "ar-aging.csv"),
                     ap_path=str(SR / "qbo-exports" / "ap-aging.csv"))
        ingest_aging(db, ar_path=str(SR / "qbo-exports" / "ar-aging.csv"),
                     ap_path=str(SR / "qbo-exports" / "ap-aging.csv"))
        conn = sqlite3.connect(db)
        ar = conn.execute("SELECT COUNT(*) FROM ar_items").fetchone()[0]
        ap = conn.execute("SELECT COUNT(*) FROM ap_items").fetchone()[0]
        conn.close()
        assert (ar, ap) == (8, 8), f"re-import must not duplicate: {(ar, ap)}"


def _run_all():
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"PASS {name}")


if __name__ == "__main__":
    _run_all()
    print("\nAll Phase 4+5 acceptance checks passed.")
