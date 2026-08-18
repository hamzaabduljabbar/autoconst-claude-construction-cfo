"""Phase 0 acceptance checks: schema shape, money rules, sample-data consistency.

Run:  python -m pytest tests/ -q      (or)   python tests/test_phase0.py
No pytest dependency required — the file runs standalone too.
"""

from __future__ import annotations

import csv
import sqlite3
import subprocess
import sys
import tempfile
from collections import defaultdict
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
SAMPLE = ROOT / "sample-data"
sys.path.insert(0, str(SCRIPTS))

from money import to_minor, from_minor, fmt_minor  # noqa: E402
import cfo  # noqa: E402


def test_money_no_float():
    assert to_minor("1234.56") == 123456
    assert from_minor(123456) == Decimal("1234.56")
    assert fmt_minor(123456) == "AUD 1,234.56"
    assert fmt_minor(-50000) == "AUD -500.00"
    try:
        to_minor(12.5)  # a float must be refused
    except TypeError:
        pass
    else:
        raise AssertionError("float was not rejected")


def test_init_and_check():
    with tempfile.TemporaryDirectory() as d:
        db = Path(d) / "t.db"
        rc = subprocess.run(
            [sys.executable, str(SCRIPTS / "cfo.py"), "init", "--db", str(db),
             "--tenant", "T", "--force"], capture_output=True, text=True)
        assert rc.returncode == 0, rc.stderr
        conn = sqlite3.connect(db)
        tables = {r[0] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'")}
        conn.close()
        assert cfo.EXPECTED_TABLES <= tables, cfo.EXPECTED_TABLES - tables


def test_sample_data_reconciles():
    """The heart of Phase 0: control totals must equal the sum of the lines,
    account totals must equal project totals, and tax must be excluded from cost."""
    if not (SAMPLE / "canonical-transactions.csv").exists():
        subprocess.run([sys.executable, str(SCRIPTS / "make_sample_data.py")],
                       check=True, capture_output=True)

    by_account = defaultdict(Decimal)
    by_project = defaultdict(Decimal)
    tax_total = Decimal(0)
    with (SAMPLE / "canonical-transactions.csv").open(encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["is_void"] == "1":
                continue
            amt = Decimal(row["amount"])
            if row["is_tax"] == "1":
                tax_total += amt
                continue
            by_account[row["account"]] += amt
            by_project[row["project"] or "(unassigned)"] += amt

    controls = defaultdict(dict)
    ctrl_tax = None
    with (SAMPLE / "control-totals.csv").open(encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["dimension"] == "tax":
                ctrl_tax = Decimal(row["amount"])
            else:
                controls[row["dimension"]][row["key"]] = Decimal(row["amount"])

    # 1. every account total matches the control report
    for acc, v in by_account.items():
        assert v == controls["account"][acc], f"account {acc}: {v} != {controls['account'][acc]}"
    # 2. every project total matches
    for prj, v in by_project.items():
        assert v == controls["project"][prj], f"project {prj}: {v} != {controls['project'][prj]}"
    # 3. account grand total == project grand total (cross-foot)
    assert sum(by_account.values()) == sum(by_project.values())
    # 4. tax reconciles and is separate from cost
    assert tax_total == ctrl_tax
    # 5. the documented grand total matches
    assert sum(by_account.values()) == controls["total"]["cost_ex_tax"]


def _run_all():
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"PASS {name}")


if __name__ == "__main__":
    _run_all()
    print("\nAll Phase 0 checks passed.")
