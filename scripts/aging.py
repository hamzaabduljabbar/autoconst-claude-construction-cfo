"""Ingest A/R and A/P aging exports into ar_items / ap_items.

Turns the cash-flow report from a schedule-based ESTIMATE into an invoice-based
forecast: real open receivables (money in, with due dates) and payables (money
out). Idempotent on the document number.

    ingest_aging(db, ar_path=None, ap_path=None) -> counts
"""

from __future__ import annotations

import csv
import datetime as _dt
import sqlite3
from pathlib import Path

from money import to_minor


def _now():
    return _dt.datetime.now().isoformat(timespec="seconds")


def _resolve_party(conn, tenant, name, role):
    if not name:
        return None
    row = conn.execute("SELECT id FROM parties WHERE tenant_id=? AND name=? AND role=?",
                       (tenant, name, role)).fetchone()
    if row:
        return row[0]
    # reuse a vendor row if one exists under a different role, else create
    return conn.execute("INSERT INTO parties(tenant_id, native_id, name, role) VALUES (?,?,?,?)",
                        (tenant, name, name, role)).lastrowid


def _resolve_project(conn, tenant, code):
    if not code:
        return None
    row = conn.execute("SELECT id FROM projects WHERE tenant_id=? AND (code=? OR native_id=?)",
                       (tenant, code, code)).fetchone()
    return row[0] if row else None


def ingest_aging(db: str, ar_path=None, ap_path=None) -> dict:
    conn = sqlite3.connect(db)
    conn.execute("PRAGMA foreign_keys = ON")
    tenant = dict(conn.execute("SELECT key,value FROM meta").fetchall())["tenant_id"]

    ss = conn.execute("SELECT id FROM source_systems WHERE tenant_id=? LIMIT 1", (tenant,)).fetchone()
    ss_id = ss[0] if ss else conn.execute(
        "INSERT INTO source_systems(tenant_id, platform) VALUES (?, 'qbo')", (tenant,)).lastrowid
    run_id = conn.execute(
        "INSERT INTO import_runs(source_system_id, adapter, adapter_version, imported_at, status) "
        "VALUES (?, 'aging', 'aging-0.1.0', ?, 'ingested')", (ss_id, _now())).lastrowid

    counts = {"ar": 0, "ap": 0}

    def load(path, table, party_role, name_col):
        n = 0
        with open(path, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                party = _resolve_party(conn, tenant, (r.get(name_col) or "").strip(), party_role)
                proj = _resolve_project(conn, tenant, (r.get("project") or "").strip())
                conn.execute(f"""INSERT OR REPLACE INTO {table}
                    (tenant_id, party_id, project_id, doc_type, doc_no, issue_date, due_date,
                     amount_minor, open_minor, import_run_id)
                    VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (tenant, party, proj, r.get("doc_type"), r.get("doc_no"),
                     r.get("issue_date"), r.get("due_date"),
                     to_minor(r["amount"]), to_minor(r["open_balance"]), run_id))
                n += 1
        return n

    if ar_path:
        counts["ar"] = load(ar_path, "ar_items", "customer", "customer")
    if ap_path:
        counts["ap"] = load(ap_path, "ap_items", "vendor", "vendor")

    conn.commit()
    conn.close()
    print(f"[aging] loaded {counts['ar']} receivables, {counts['ap']} payables")
    return counts
