# IDBI AWS Billing Tool  v3.0

Three-tab local web tool for generating BoM-applied AWS billing sheets.

> **The complete, authoritative rule set is in [`RULES.md`](RULES.md).**
> This README is just setup + a summary.

## Setup (once)

```bash
cd BillingOm
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Opens at http://localhost:8501

## Three tabs

| Tab | Purpose | Files needed |
|---|---|---|
| DBD Monthly Billing | DBD account billing sheet | CUR CSV + Invoice PDF + BoM xlsx |
| DR / Estimate Sheet | DR or cost estimate from calculator | Calculator CSV + BoM xlsx |
| ITD Monthly Billing | ITD account billing sheet | CUR CSV + Invoice PDF + BoM xlsx |

## What changed in v3

v3 folds in every fix from the two review cases plus four new rules. All are
documented as enforceable rules in `RULES.md`; in brief:

1. **Windows EC2 uses m5a in place of m6a** (both scale through the m5a BoM
   lines). Windows **r6** stays non-BoM @8%, unchanged.
2. **Direct Connect** is priced from the BoM leased-line rate
   ("Dedicated DC-Cloud connect leased line charges", ₹19,912.5 / 1 Gbps):
   `(((19912.5/1000)*mbps)*qty)/rate`, discount 75%. Applied on every tab.
3. **Zero-cost free-tier lines are removed** — never billed, never counted
   toward a BoM quantity.
4. **Presentation** matches the reviewed manual/Claude sheet, and the **Notes
   column appears only when there are notes**.

Bug fixes captured as rules (see `RULES.md` R1–R6):
- Compute rows are tagged to their **own** account (no more mis-attribution).
- **CPU-credit** charges are always billed when present.
- **No instance resolves to 0 vCPU/mem** (unknown types get a parametric
  fallback), so nothing wrongly drops to the 8% path.
- The footer month is always the sheet's actual month.

## To reopen next time

```bash
cd BillingOm
source .venv/bin/activate        # Windows: .venv\Scripts\activate
streamlit run app.py
```
