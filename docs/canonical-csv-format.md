# Canonical CSV format

The universal transaction-line format the `canonical` adapter reads. Any
accounting platform (QuickBooks, Xero, MYOB, Zoho, a bespoke spreadsheet) can
export to this shape — that's how you use the pipeline without being locked to
one vendor.

One row per **transaction line** (not per transaction — a single bill split
across two projects is two rows).

## Columns (in order)

| Column | Required | Example | Notes |
|---|---|---|---|
| `tenant` | yes | `Northwind Construction Pty Ltd` | Legal entity name |
| `platform` | yes | `qbo` \| `xero` \| `myob` \| `csv` | Source system tag |
| `txn_type` | yes | `Bill`, `Expense`, `JournalEntry`, `CreditNote` | Native transaction type |
| `txn_id` | yes | `B-1002` | Native transaction id (business identity for dedup) |
| `line_id` | recommended | `B-1002-L2` | Stable line id within the transaction. If blank, the importer fingerprints the line. |
| `date` | yes | `2026-07-05` | ISO 8601 |
| `vendor` | yes | `SteelFab Australia` | Party name |
| `account` | yes | `Materials` | Chart-of-accounts name |
| `project` | optional | `PRJ-002` \| blank | Project id/code. Blank = unassigned (stays visible) |
| `cost_code` | optional | `05-12-00` \| blank | Cost-code. Blank = classify later |
| `tax_code` | optional | `GST` \| blank | Tax jurisdiction code |
| `amount` | yes | `5400.00` | Signed decimal string. Credits negative. **Never a float — always a string.** |
| `currency` | yes | `AUD` | ISO 4217 |
| `is_tax` | yes | `0` \| `1` | Set 1 for GST/VAT lines. Tax NEVER mixes into cost totals. |
| `is_void` | yes | `0` \| `1` | Voided transactions kept but excluded from totals |
| `memo` | optional | `SHS columns — canopy` | Free text (sanitized against Excel formula injection on export) |

## Example (2 lines of a 4-line bill)

```
tenant,platform,txn_type,txn_id,line_id,date,vendor,account,project,cost_code,tax_code,amount,currency,is_tax,is_void,memo
Northwind Construction Pty Ltd,qbo,Bill,B-1002,B-1002-L1,2026-07-05,SteelFab Australia,Materials,PRJ-01,05-12-00,,8600.00,AUD,0,0,310UB structural steel — portal frames
Northwind Construction Pty Ltd,qbo,Bill,B-1002,B-1002-L2,2026-07-05,SteelFab Australia,Materials,PRJ-02,05-12-00,,5400.00,AUD,0,0,SHS columns — canopy
```

## Rules

- **Amounts are strings**, not numbers. The importer converts via `Decimal` →
  integer minor units. A binary float will be rejected at the money boundary.
- **`txn_id` must be stable across exports** — it's the dedup key.
  Re-importing the same file must not create duplicates.
- **Tax lines must have `is_tax=1`** and their own row. The reconciliation
  gate checks tax separately.
- **Voids stay** — write the row with `is_void=1` rather than deleting it, so
  the audit trail is preserved.
- **Unassigned project cost** — leave `project` blank rather than fabricating
  one. The pipeline shows it under `(unassigned)` and never hides it.

## Optional companion: control totals

For the reconciliation gate, supply a `control-totals.csv` alongside:

```
dimension,key,amount,currency
account,Materials,30360.00,AUD
account,Subcontractors,53450.00,AUD
project,PRJ-001,52850.00,AUD
project,(unassigned),4620.00,AUD
tax,GST Payable,9071.00,AUD
total,cost_ex_tax,105210.00,AUD
```

The ingest checks account totals, project totals, tax total, and grand total
match to the cent, and that account grand total cross-foots to project grand
total. Any mismatch → run marked `unreconciled` and every downstream report is
stamped **DRAFT — NOT TRUSTED**.
