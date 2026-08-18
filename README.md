# Construction CFO

**Your accounting data goes in. Four reports a CFO would make come out.**

Job costing, retention tracker, cash flow forecast, and an expense audit —
all in Excel, every number traceable back to the transaction it came from.
No accountant needed. No monthly SaaS fee. Runs on your own machine.

---

## What you get

Point it at your QuickBooks export and a workbook of your project info, and
it gives you four Excel reports:

**Job Cost** — every project, every budget line. What you spent vs what you
budgeted. Anything over budget is highlighted in orange so you can see
where the money's leaking at a glance.

**Retention Tracker** — how much retention your clients are holding from
you (money you get back at project completion), and how much you're
holding from your subbies. Overdue releases flagged in red so you know
which claims to chase first.

**Cash Flow Forecast** — 13-week outlook. Three scenarios (on-time /
expected / worst case). Tells you the week your bank balance drops
lowest and whether you're going to run out. Load your open invoices too
and it becomes a proper invoice-based forecast, not just an estimate.

**Expense Audit** — flags possible duplicate bills, unusually large
charges, and vendors where you spend enough that a supply agreement
might save you real money.

---

## Who this is for

Builders doing $500k–$20M a year who can't justify hiring a CFO or paying
for construction-specific accounting software, but who are big enough to
lose real money on a job before they notice.

If you're using QuickBooks Online (or any accounting tool you can export
CSVs from), you can use this.

---

## What you need

1. **Your accounting data** — export a "Transaction Detail by Account"
   report from QuickBooks as CSV. Takes 30 seconds.

2. **Your project info** — one Excel workbook listing your projects,
   budgets, how far along each is, your contracts and subcontractors.
   A filled-in example is included so you know exactly what goes where.

That's it. Nothing to log into. No API keys. Your data stays on your
computer.

---

## How to use it

You need Claude Code installed. Then:

```
git clone https://github.com/hamzaabduljabbar/autoconst-claude-construction-cfo
cd autoconst-claude-construction-cfo
pip install -r requirements.txt
```

Drop your accounting export and your project workbook into the `inputs/`
folder, then ask Claude:

> "Run the construction CFO on my QuickBooks export."

Claude does the rest. Four Excel files land in `outputs/`. Open them.

## Try it before you use your own data

An example dataset is included so you can see exactly what the reports
look like before you set anything up. Ask Claude:

> "Run the construction CFO on the example dataset."

You'll get the four reports back in about 10 seconds. The Job Cost report
will show you every project blowing its Site Services budget — a total
of ~$847k in overruns the tool catches automatically.

---

## What it will and won't tell you

**Real, straight from your books:** which jobs are over budget and by how
much. Where you're spending. Duplicate bills. Vendor concentration.

**Honest estimates (labeled as such):** retention amounts (calculated
from % complete — a real retention ledger lives outside accounting), and
cash flow when you only load transactions (upgrades to a real forecast
when you also load your open invoices).

**Not designed to do:** write anything back to QuickBooks, replace your
bookkeeper, or do tax. It only reads.

---

## Frequently asked

**Will my data leave my computer?** No. Everything runs locally.

**Do I need to be technical?** You need to be able to install Claude
Code and run one command to set it up. After that, you just talk to
Claude in plain English.

**What if I don't use QuickBooks?** If your platform can export CSVs
(Xero, MYOB, Zoho, or a spreadsheet you keep), you can convert to the
included format. See `docs/canonical-csv-format.md`.

**What if my accounting data doesn't match up?** The tool tells you.
It won't produce a report on numbers it can't verify. Any mismatch is
flagged and the report is stamped "DRAFT — NOT TRUSTED" so you never
mistake bad data for good.

---

## Built by

Hamza Jabbar — hamzajabbar.online
