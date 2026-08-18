"""Phase 6 acceptance tests — expense leak audit detectors.

Summit Ridge is clean data, so these tests PLANT known leaks in a small database
and assert the detectors catch them (and don't false-fire on normal spread).

Run:  python tests/test_phase6.py
"""

from __future__ import annotations

import subprocess
import sqlite3
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from audit import compute_audit, SPIKE_MIN_SAMPLE, CONCENTRATION_MIN  # noqa: E402


def _db(d: Path) -> sqlite3.Connection:
    db = str(d / "a.db")
    subprocess.run([sys.executable, str(SCRIPTS / "cfo.py"), "init", "--db", db,
                    "--tenant", "T", "--force"], check=True, capture_output=True)
    conn = sqlite3.connect(db)
    conn.execute("PRAGMA foreign_keys = ON")
    ss = conn.execute("INSERT INTO source_systems(tenant_id, platform) VALUES ('T','qbo')").lastrowid
    conn.execute(
        "INSERT INTO import_runs(source_system_id, adapter, adapter_version, imported_at, status) "
        "VALUES (?, 'test','t', '2026-07-01', 'ingested')", (ss,))
    return conn


def _vendor(conn, name):
    return conn.execute("INSERT INTO parties(tenant_id, native_id, name, role) "
                        "VALUES ('T',?,?, 'vendor')", (name, name)).lastrowid


def _acc(conn, name="Materials"):
    row = conn.execute("SELECT id FROM accounts WHERE name=?", (name,)).fetchone()
    if row:
        return row[0]
    return conn.execute("INSERT INTO accounts(tenant_id, native_id, name, account_type) "
                        "VALUES ('T',?,?, 'cogs')", (name, name)).lastrowid


def _bill(conn, vendor_id, acc_id, date, amount_minor, memo, txn):
    run_id = conn.execute("SELECT id FROM import_runs LIMIT 1").fetchone()[0]
    hid = conn.execute("""INSERT INTO transaction_headers(import_run_id, tenant_id, source_platform,
        native_transaction_type, native_transaction_id, party_id, txn_date, currency, is_void)
        VALUES (?,'T','qbo','Bill',?,?,?, 'AUD', 0)""", (run_id, txn, vendor_id, date)).lastrowid
    conn.execute("""INSERT INTO transaction_lines(header_id, account_id, amount_minor, is_tax, memo)
        VALUES (?,?,?,0,?)""", (hid, acc_id, amount_minor, memo))


def test_detects_duplicate_bill():
    with tempfile.TemporaryDirectory() as d:
        conn = _db(Path(d))
        v = _vendor(conn, "Acme Supplies"); a = _acc(conn)
        # same vendor + amount + memo on two separate bills
        _bill(conn, v, a, "2026-07-01", 480000, "Concrete delivery", "B-1")
        _bill(conn, v, a, "2026-07-08", 480000, "Concrete delivery", "B-2")
        conn.commit()
        res = compute_audit(conn); conn.close()
        assert any(x["severity"] == "HIGH" for x in res["duplicates"]), res["duplicates"]


def test_detects_cost_spike_with_robust_test():
    with tempfile.TemporaryDirectory() as d:
        conn = _db(Path(d))
        v = _vendor(conn, "Normal Vendor"); a = _acc(conn)
        # a stable baseline of small lines, then one big outlier
        for i in range(SPIKE_MIN_SAMPLE + 1):
            _bill(conn, v, a, "2026-07-01", 100000 + i * 500, f"routine {i}", f"N-{i}")
        _bill(conn, v, a, "2026-07-15", 5_000_000, "one-off huge charge", "SPIKE")
        conn.commit()
        res = compute_audit(conn); conn.close()
        assert any("huge" in (s["example"][7] or "") for s in res["spikes"]), \
            [s["amount"] for s in res["spikes"]]


def test_no_false_spike_on_uniform_vendor():
    with tempfile.TemporaryDirectory() as d:
        conn = _db(Path(d))
        v = _vendor(conn, "Steady Vendor"); a = _acc(conn)
        for i in range(10):
            _bill(conn, v, a, "2026-07-01", 200000 + i * 1000, f"line {i}", f"S-{i}")
        conn.commit()
        res = compute_audit(conn); conn.close()
        assert res["spikes"] == [], "uniform spend must not raise spike alerts"


def test_small_vendor_not_judged():
    """Below the minimum sample size, a vendor is never spike-flagged."""
    with tempfile.TemporaryDirectory() as d:
        conn = _db(Path(d))
        v = _vendor(conn, "Tiny Vendor"); a = _acc(conn)
        _bill(conn, v, a, "2026-07-01", 100000, "a", "T-1")
        _bill(conn, v, a, "2026-07-02", 9_000_000, "huge but only 2 samples", "T-2")
        conn.commit()
        res = compute_audit(conn); conn.close()
        assert res["spikes"] == [], "too few samples to judge — must not flag"


def test_vendor_concentration_flagged():
    with tempfile.TemporaryDirectory() as d:
        conn = _db(Path(d))
        big = _vendor(conn, "Dominant Supplier"); a = _acc(conn)
        _bill(conn, big, a, "2026-07-01", CONCENTRATION_MIN + 50_000_00, "bulk", "D-1")
        small = _vendor(conn, "Minor Supplier")
        _bill(conn, small, a, "2026-07-01", 10_000_00, "small", "M-1")
        conn.commit()
        res = compute_audit(conn); conn.close()
        names = [c["vendor"] for c in res["concentration"]]
        assert "Dominant Supplier" in names and "Minor Supplier" not in names


def _run_all():
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"PASS {name}")


if __name__ == "__main__":
    _run_all()
    print("\nAll Phase 6 acceptance checks passed.")
