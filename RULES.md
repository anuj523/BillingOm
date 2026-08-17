# IDBI AWS Billing Tool — Encoded Rules (v3)

This file is the single source of truth for the pricing rules the tool applies.
Every rule below is enforced in code (`idbi_billing/config.py`, `pricing.py`,
`dr_estimate.py`, `render.py`). Rules apply to **all three tabs** — DBD billing,
ITD billing, and DR / Estimate — unless stated otherwise.

## Data sourcing
- **CUR CSV** → quantities, configuration and SKU text only. Never pricing.
  Only `RecordType = LinkedLineItem` rows for the tab's allowed accounts.
- **Invoice PDF** → USD→INR conversion rate, billing period, and the gross USD
  charges that price non-BoM rows.
- **Section B BoM xlsx** → the pricing authority for orange (BoM) rows; read
  live from the sheet named `Section B AWS BoM` (col A serial, col F unit price,
  col G discount %). Falls back to the values in `config.py` if unreadable.

## Row types
- **Orange (BoM) rows** — unit price and discount % come from the BoM, scaled to
  actual consumption. Formula chain per row: `J = I/H`, `L = J·(1−K)`,
  `M = L·H`, `N = M·rate`.
- **Non-BoM rows** — 8% discount off public pricing (Section D).

## EC2 substitution
- **Exact match first:** `(OS, vCPU, mem)` → BoM line.
- **RHEL r6** → anchor `rhel_4v32g` (r6a.xlarge, 4 vCPU), scale by vCPU ratio.
- **RHEL m6** → anchor `rhel_16v64g` (m6g.4xlarge, 16 vCPU), scale by vCPU ratio.
- **Windows m-family (m5a *and* m6a)** → priced from the **m5a BoM lines**
  (`win_4v16g` / `win_8v32g` / `win_16v64g`), scale by vCPU ratio. *(point 1 —
  "take m5a in place of m6a".)*
- **Windows r6** → **no BoM equivalent → 8%**, left as-is. *(point 1.)*
- Discount % is always the BoM discount, unchanged by any scaling multiplier.

## EBS Block SSD (gp3) slabs
`≤128→128`, `≤256→256`, `≤512→512`, `≤1024→1024`, `>1024→2048` GB. Disc 41.078%.
EBS **Snapshot** is a separate, non-BoM row @8% (never folded into Block SSD).

## Direct Connect  *(point 2 — new)*
BoM line 40, **"Dedicated DC-Cloud connect leased line charges"**: ₹19,912.5
covers **1 Gbps (= 1000 Mbps)**. Costing:

```
Indicative (USD) = ((( 19912.5 / 1000 ) * actual_mbps ) * quantity ) / rate
Discount % = 75%
```

Link speed (Mbps) is parsed from the CUR ItemDescription/UsageType (e.g.
"200Mbps", "1G", "10Gbps"); quantity = number of connections
(port-hours / 730, min 1). Applied on every tab.
*Worked example:* 200 Mbps × 1 → ₹995.63 after the 75% discount.

## Other BoM rows (unchanged)
Object Storage Hot (per 25 TB, 39.18%), NAT Data (per 100 GB, 70%; NAT hours
stay 8%), Static Public IP (per IP, 99%, folds all VPC IPv4), Application LB
(per 730 ALB-hrs, 75%), Network Firewall (line 37, 60%, 2 endpoints = 1 unit),
WAF (per 1M requests, 40.29%), KMS (line 39, 99%), Network Data Transfer
(80%, egress hard-set to 0 GB).

---

# Bug-fix rules captured from Case 1 & Case 2

These were defects in the previous tool; each is now a hard rule.

## R1 — Compute rows belong to their own account  *(Case 1)*
Each account's compute rows are tagged to **that** account at the moment they
are built. The previous code re-tagged the entire accumulated list with the
*last* account, mis-attributing other accounts' instances. **Rule:** never
re-stamp already-built rows with a later account.

## R2 — CPU credits are always billed when present  *(Case 1)*
T3/T3A CPU-credit charges are emitted as their own sub-row whenever
`CostBeforeTax > 0`. The previous guard (`len(rows) == 0`) dropped them
whenever the account also had box usage.

## R3 — Every instance resolves to real vCPU/mem  *(Case 2)*
Instance vCPU/memory are resolved from an expanded table, and any instance not
in the table falls back to a parametric estimate (size → vCPU, family letter →
memory ratio). **An instance must never resolve to (0, 0)** — that was what
silently pushed `m5a.xlarge` and `t3.small` onto the 8% path instead of the
Windows BoM line.

## R4 — Zero-cost free-tier lines are excluded  *(point 3)*
Any CUR line with `CostBeforeTax = 0` (e.g. "under monthly free tier") is
dropped before pricing — it never becomes its own row **and** never contributes
consumption to a BoM quantity (this removes the 850 GB / 1044 GB free-tier
inflation seen in Case 2).

## R5 — Correct, dynamic footer month  *(Case 1)*
The Section-D footer note always interpolates the sheet's own month
(`Apr-2026`, etc.). No hard-coded month is ever emitted.

## R6 — Notes column only when there are notes  *(presentation)*
The "Notes" column (and the Working Notes block) is rendered **only** when at
least one row carries a note or a working note exists. With no notes, the sheet
ends at "Converted Monthly Cost" (14 columns).

## R7 — No billable service is silently dropped  *(v3.1 catch-all)*
Every CUR `ProductName` with `CostBeforeTax > 0` that has **no dedicated builder**
is billed by a catch-all at the non-BoM **8%** rate (one row per account/service),
so unmodelled services (Route 53, CloudFront, Lambda, SNS, SQS, Secrets Manager,
EFS, DynamoDB, …) can no longer vanish from the sheet. Zero-cost/free-tier lines
stay excluded (R4). The catch-all does **not** touch usage types inside
already-handled products — VPN (under "Amazon Virtual Private Cloud"), Transit
Gateway priced as BoM, and RDS BoM family substitution still require their
Section B BoM rates and are pending those inputs.

## R8 — Network Firewall is priced per endpoint  *(Dec-2025 approved)*
Each running firewall endpoint = 1 BoM unit: `qty = max(1, round(endpoint_hours/730))`.
The earlier "2 endpoints = 1 firewall" rule under-billed by half (e.g. 1488 endpoint-hrs
= 2 endpoints → qty 2, ₹46,314.72, not qty 1 / ₹23,157.36).

## R9 — Site-to-Site VPN (BoM line 35)
VPN usage (APS3-VPN-Usage-Hours under "Amazon Virtual Private Cloud") is a BoM
row: ₹10,035 per connection per month, discount 78%. Connections =
max(1, round(VPN-usage-hours / 730)). e.g. 1488 hrs → 2 connections → ₹4,415.40.

## R10 — Non-BoM rows price from invoice GROSS, not CUR net
The CUR's CostBeforeTax is already net of the AWS Distribution Program Discount
(~12%). Non-BoM rows are "8% off public (gross) pricing", so pricing them from
CUR net double-discounts and under-bills. Each non-BoM row's indicative USD is
scaled up by the account's gross/net ratio (invoice gross ÷ CUR net) before the
8% is applied. Fixes Transit Gateway, Route 53, CloudTrail, CloudWatch, Config,
etc. to match the approved sheet exactly.
