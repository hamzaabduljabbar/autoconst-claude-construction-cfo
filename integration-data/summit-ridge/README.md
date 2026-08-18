# Summit Ridge realistic integration dataset

This package is **realistic synthetic data**, not the records of a real contractor.
It is generated deterministically with seed `20260814` and targets the existing
Phase 1 canonical and grouped-QBO adapters.

## Scenario

- Australian accrual-basis contractor; AUD functional currency; 10% GST.
- Six construction projects across 18 months.
- Healthy, labour-overrun, late-client, retention-heavy, materials-overrun, and
  pending-variation project scenarios.
- Bills, GST lines, payroll journals, overhead, a credit, a void, an ambiguous
  cost code, a split-project bill, and hostile spreadsheet/prompt-like text.

## Regenerate

```bash
python scripts/make_realistic_data.py
node scripts/build_realistic_workbook.mjs
python tests/test_realistic_data.py
```

Do not replace the small Phase 0 fixture. Use this package for integration,
scale, idempotency, reconciliation, and later reporting tests.
