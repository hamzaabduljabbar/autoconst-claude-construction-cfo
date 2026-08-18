"""Platform adapters — turn a source export into normalized transaction lines.

Every adapter yields the SAME record shape (a NormLine), so the importer never
needs to know which platform the data came from:

    NormLine = {
        txn_type, txn_id, native_line_id (str|None), date (ISO str),
        vendor, account (name), project_raw (str, '' = unassigned),
        cost_code (str, '' = none), tax_code (str), amount (Decimal),
        currency, is_tax (bool), is_void (bool), memo (str)
    }

project_raw is resolved to a canonical project by the importer (a display name
from QBO or a native id from canonical both map to the same project).

Adapters are the ONLY place platform-specific mess lives. Downstream code is
platform-agnostic.
"""

from __future__ import annotations

import csv
import datetime as _dt
from decimal import Decimal, InvalidOperation
from pathlib import Path


CANONICAL_ADAPTER_VERSION = "canonical-0.1.0"
QBO_ADAPTER_VERSION = "qbo-txn-detail-0.1.0"


def _norm(txn_type, txn_id, native_line_id, date, vendor, account, project_raw,
          cost_code, tax_code, amount, currency, is_tax, is_void, memo):
    return dict(
        txn_type=txn_type, txn_id=txn_id, native_line_id=native_line_id,
        date=date, vendor=vendor, account=account, project_raw=project_raw or "",
        cost_code=cost_code or "", tax_code=tax_code or "",
        amount=amount, currency=currency, is_tax=bool(is_tax),
        is_void=bool(is_void), memo=memo or "",
    )


# --------------------------------------------------------------------------- #
# Canonical CSV
# --------------------------------------------------------------------------- #

def parse_canonical(path: str | Path):
    """The clean, documented canonical transaction-line format."""
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            yield _norm(
                txn_type=row["txn_type"].strip(),
                txn_id=row["txn_id"].strip(),
                native_line_id=(row.get("line_id") or "").strip() or None,
                date=row["date"].strip(),
                vendor=row.get("vendor", "").strip(),
                account=row["account"].strip(),
                project_raw=(row.get("project") or "").strip(),
                cost_code=(row.get("cost_code") or "").strip(),
                tax_code=(row.get("tax_code") or "").strip(),
                amount=Decimal(row["amount"]),
                currency=row.get("currency", "AUD").strip() or "AUD",
                is_tax=(row.get("is_tax") == "1"),
                is_void=(row.get("is_void") == "1"),
                memo=(row.get("memo") or "").strip(),
            )


# --------------------------------------------------------------------------- #
# QuickBooks Online — "Transaction Detail by Account"
# --------------------------------------------------------------------------- #

def _parse_money(s: str) -> Decimal | None:
    """'$12,400.00' / '-$1,800.00' / '($1,800.00)' -> Decimal. '' -> None."""
    s = (s or "").strip()
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace("$", "").replace(",", "").strip()
    if not s:
        return None
    try:
        d = Decimal(s)
    except InvalidOperation:
        return None
    return -d if neg else d


def _parse_date(s: str) -> str | None:
    """MM/DD/YYYY -> ISO. Returns None if it isn't a date (e.g. a group header)."""
    s = (s or "").strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return _dt.datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return None


# rows we skip outright in the grouped report
_SKIP_PREFIXES = ("total for", "total ", "beginning balance")


# header-cell aliases -> our canonical column role. Real QBO uses "Transaction
# date"/"Description"; our synthetic fixture used "Date"/"Memo/Description".
_COL_ALIASES = {
    "date": "date", "transaction date": "date",
    "transaction type": "type", "type": "type",
    "num": "num", "no.": "num", "number": "num",
    "name": "name",
    "customer/project": "project", "customer": "project", "project": "project",
    "memo/description": "memo", "description": "memo", "memo": "memo",
    "amount": "amount",
    "split": "split", "distribution account": "dist_account",
    "balance": "balance",
}


def _find_header(rows):
    """Return (index, {role: col_index}) for the column-header row. Detected by the
    presence of a 'Transaction type' cell, so a leading blank column and a
    'Date' vs 'Transaction date' difference are both tolerated."""
    for i, r in enumerate(rows):
        norm = [c.strip().lower() for c in r]
        if "transaction type" in norm or "type" in norm:
            roles = {}
            for ci, cell in enumerate(norm):
                role = _COL_ALIASES.get(cell)
                if role and role not in roles:
                    roles[role] = ci
            if "type" in roles and "amount" in roles:
                return i, roles
    raise ValueError("could not find the QBO column header row "
                     "(expected a 'Transaction type' and 'Amount' column)")


def parse_qbo(path: str | Path):
    """Reconstruct transaction lines from a grouped 'Transaction Detail by
    Account' export (also handles the synthetic messy fixture).

    The account is a GROUP HEADER row, not a per-line column; nested sub-accounts
    simply set a deeper leaf account. Subtotal ('Total for ...'), grand-total
    ('TOTAL'), title, footer and blank rows are discarded. A blank Amount (e.g.
    inventory-adjust stubs) is skipped. Voided rows are flagged via the memo.

    NOTE: this report is double-entry — a transaction appears under every account
    it posts to. Isolating *cost* therefore requires account-type filtering
    downstream; this adapter faithfully yields every posting line as-is.
    """
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))

    header_idx, col = _find_header(rows)
    ci_date = col.get("date")
    ci_type = col.get("type")
    ci_num = col.get("num")
    ci_name = col.get("name")
    ci_proj = col.get("project")
    ci_memo = col.get("memo")
    ci_amt = col["amount"]

    def cell(r, ci):
        return r[ci].strip() if ci is not None and ci < len(r) else ""

    current_account = None
    for r in rows[header_idx + 1:]:
        if not r or not any(c.strip() for c in r):
            continue  # blank separator
        first_nonempty = next((c.strip() for c in r if c.strip()), "")
        low = first_nonempty.lower()
        if low == "total" or any(low.startswith(p) for p in _SKIP_PREFIXES):
            continue  # subtotal / grand total
        if "basis" in low and ("am " in low or "pm " in low or "gmt" in low):
            continue  # footer timestamp
        ttype = cell(r, ci_type)
        date_cell = cell(r, ci_date)
        iso = _parse_date(date_cell)
        # group-header row: no parseable date and no transaction type ->
        # the first non-empty cell names the (leaf) account for what follows.
        if iso is None and not ttype:
            current_account = first_nonempty
            continue
        if iso is None or not ttype:
            continue  # not a data row we understand
        if current_account is None:
            continue
        amount = _parse_money(cell(r, ci_amt))
        if amount is None:
            continue  # blank amount -> not a postable line (e.g. inventory stub)
        memo = cell(r, ci_memo)
        is_void = memo.lower().startswith("voided")
        account = current_account
        is_tax = account.lower().startswith("gst")
        yield _norm(
            txn_type=ttype,
            txn_id=cell(r, ci_num),
            native_line_id=None,   # no stable line id in this report -> fingerprint path
            date=iso,
            vendor=cell(r, ci_name),
            account=account,
            project_raw=cell(r, ci_proj),
            cost_code="",          # cost code lives OUTSIDE QBO -> classified later
            tax_code="GST" if is_tax else "",
            amount=amount,
            currency="AUD",
            is_tax=is_tax,
            is_void=is_void,
            memo=memo[len("Voided:"):].strip() if is_void else memo,
        )


def validate_qbo_subtotals(path: str | Path):
    """Grade the parser using the report's OWN 'Total for X' subtotal rows.
    Returns (checked, mismatches). A leaf account's parsed line-sum must equal
    its declared 'Total for <account>' figure (the non 'with sub-accounts' one).
    """
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    header_idx, col = _find_header(rows)
    ci_date, ci_type, ci_amt = col.get("date"), col.get("type"), col["amount"]

    def cell(r, ci):
        return r[ci].strip() if ci is not None and ci < len(r) else ""

    current, running = None, {}
    declared = {}
    for r in rows[header_idx + 1:]:
        if not r or not any(c.strip() for c in r):
            continue
        first = next((c.strip() for c in r if c.strip()), "")
        low = first.lower()
        if low.startswith("total for ") and "with sub-accounts" not in low:
            name = first[len("Total for "):].strip()
            amt = _parse_money(cell(r, ci_amt))
            if amt is not None:
                declared[name] = amt
            continue
        if low == "total" or low.startswith("total for"):
            continue
        ttype = cell(r, ci_type)
        iso = _parse_date(cell(r, ci_date))
        if iso is None and not ttype:
            current = first
            running.setdefault(current, Decimal("0"))
            continue
        if iso is None or not ttype or current is None:
            continue
        amt = _parse_money(cell(r, ci_amt))
        if amt is not None:
            running[current] = running.get(current, Decimal("0")) + amt

    checked, mismatches = 0, []
    for name, want in declared.items():
        got = running.get(name)
        if got is None:
            continue
        checked += 1
        if got != want:
            mismatches.append((name, got, want))
    return checked, mismatches


ADAPTERS = {
    "canonical": (parse_canonical, CANONICAL_ADAPTER_VERSION),
    "qbo": (parse_qbo, QBO_ADAPTER_VERSION),
}
