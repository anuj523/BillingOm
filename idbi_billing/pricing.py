"""
Pricing engine — converts CUR + invoice + BoM into ordered Row objects.
Encodes every rule from the rule book (see RULES.md).

v3 fixes & rules:
  * Case-1: per-account compute rows are tagged to their OWN account
    (the old code overwrote every row with the last account).
  * Case-1: T3A CPU-credit lines are always billed when present
    (old guard `len(rows)==0` dropped them whenever box usage existed).
  * Case-2: instance vCPU/mem always resolve to a non-zero value — table
    first, parametric fallback second — so a real instance never silently
    drops to the 8% path because it read (0,0).
  * Point 1: Windows m-family (m5a/m6a) priced from the m5a BoM lines;
    Windows r6 stays non-BoM @8%.
  * Point 2: Direct Connect priced from the BoM leased-line rate.
  * Point 3: zero-cost free-tier CUR lines are excluded from the sheet.
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field

from .config import (
    EC2_EXACT, EC2_ANCHOR, WIN_M5A_ANCHORS, WIN_M_FAMILIES,
    ebs_key, NON_BOM_DISC, ADDR_HRS_PER_MONTH,
)
from .cur import CUR


@dataclass
class Row:
    service:    str
    additional: str
    config:     str
    sku:        str
    qty:        int | float
    i_formula:  str          # Excel formula string; {r} replaced at render time
    discount:   float
    is_bom:     bool
    note_n:     int | None = None


@dataclass
class PricingResult:
    rows:          list = field(default_factory=list)
    working_notes: list = field(default_factory=list)   # [(n, text)]


# ── EC2 family helpers ──────────────────────────────────────────────────────
_FAM_RE = re.compile(r"([a-z]+\d+)")

# Standard vCPU per instance *size* (used by the parametric fallback so no
# instance ever resolves to 0 vCPU — Case-2 fix).
_SIZE_VCPU = {
    "nano": 2, "micro": 2, "small": 2, "medium": 2, "large": 2,
    "xlarge": 4, "2xlarge": 8, "3xlarge": 12, "4xlarge": 16,
    "6xlarge": 24, "8xlarge": 32, "9xlarge": 36, "12xlarge": 48,
    "16xlarge": 64, "24xlarge": 96, "32xlarge": 128, "48xlarge": 192,
}
# GB of memory per vCPU by family letter (c = compute, m = general, r/x = mem).
_FAM_MEM_PER_VCPU = {"c": 2, "m": 4, "r": 8, "x": 16, "t": 2, "i": 8, "d": 6, "z": 8}


def _base_family(inst: str) -> str:
    """c6a -> c6, m6a -> m6, r6a -> r6, m5a -> m5, t3a -> t3, etc."""
    fam = inst.split(".")[0]          # e.g. "m6a"
    m = _FAM_RE.match(fam)
    if not m:
        return fam
    token = m.group(1)                # "m6a" style already captured letters+digit
    # strip trailing alpha suffix (a/g/n/d ...) but keep the letter+generation
    mm = re.match(r"([a-z])(\d+)", token)
    return f"{mm.group(1)}{mm.group(2)}" if mm else token


def _os_class(desc: str) -> str:
    """Guess OS from ItemDescription."""
    d = desc.lower()
    if "windows" in d:  return "win"
    if "red hat" in d or "rhel" in d: return "rhel"
    return "linux"


# Client rule: only these families qualify for BoM compute pricing.
# M5<->M6 and C5<->C6 are interchangeable (they share the same vCPU/GB BoM
# lines), so the qualifying base-family set is simply {m5, m6, c5, c6}.
_BOM_COMPUTE_FAMILIES = {"m5", "m6", "c5", "c6"}


def _bom_eligible(inst: str) -> bool:
    """
    True if the instance may be priced from the BoM compute lines.
    Only m5/m6/c5/c6 families qualify. Intel 'i' variants (family string ends
    in 'i', e.g. m6i, c6i) are excluded and billed at standard public price /
    8% (Section-D), per the client's rule.
    """
    fam = inst.split(".")[0]              # raw family token, e.g. "m6i", "c6a"
    if fam.endswith("i"):
        return False
    return _base_family(inst) in _BOM_COMPUTE_FAMILIES


def instance_specs(inst_type: str) -> tuple[int, int]:
    """
    Map an instance type string to (vcpu, mem_gb).
    Table first; parametric fallback second so nothing returns (0, 0).
    """
    table = {
        "t3.nano":(2,1),"t3.micro":(2,1),"t3.small":(2,2),"t3.medium":(2,4),
        "t3.large":(2,8),"t3.xlarge":(4,16),"t3.2xlarge":(8,32),
        "t3a.nano":(2,1),"t3a.micro":(2,1),"t3a.small":(2,2),
        "t3a.medium":(2,4),"t3a.large":(2,8),"t3a.xlarge":(4,16),"t3a.2xlarge":(8,32),
        "c6a.large":(2,4),"c6a.xlarge":(4,8),"c6a.2xlarge":(8,16),
        "c6a.4xlarge":(16,32),"c6a.8xlarge":(32,64),
        "c6g.large":(2,4),"c6g.xlarge":(4,8),"c6g.2xlarge":(8,16),
        "m5.large":(2,8),"m5.xlarge":(4,16),"m5.2xlarge":(8,32),"m5.4xlarge":(16,64),
        "m5a.large":(2,8),"m5a.xlarge":(4,16),"m5a.2xlarge":(8,32),
        "m5a.4xlarge":(16,64),"m5a.8xlarge":(32,128),"m5a.12xlarge":(48,192),
        "m6a.large":(2,8),"m6a.xlarge":(4,16),"m6a.2xlarge":(8,32),
        "m6a.4xlarge":(16,64),"m6a.8xlarge":(32,128),"m6a.16xlarge":(64,256),
        "m6g.large":(2,8),"m6g.xlarge":(4,16),"m6g.2xlarge":(8,32),
        "m6g.4xlarge":(16,64),
        "r6a.large":(2,16),"r6a.xlarge":(4,32),"r6a.2xlarge":(8,64),
        "r6a.4xlarge":(16,128),"r6a.8xlarge":(32,256),
        "r5.large":(2,16),"r5.xlarge":(4,32),"r5.2xlarge":(8,64),
    }
    if inst_type in table:
        return table[inst_type]
    # Parametric fallback: size -> vCPU, family letter -> mem/vCPU ratio.
    parts = inst_type.split(".")
    if len(parts) == 2:
        fam, size = parts
        vcpu = _SIZE_VCPU.get(size, 0)
        letter = fam[0] if fam else "m"
        mem = vcpu * _FAM_MEM_PER_VCPU.get(letter, 4)
        if vcpu:
            return vcpu, mem
    return 0, 0


def _ec2_pricing(inst: str, vcpu: int, mem: int, os_cls: str, bom: dict):
    """
    Client rules for BoM compute substitution:
      * All non-Windows instances are treated as RHEL — the BoM only has Windows
        and RHEL compute lines ("Considering OS as RedHat").
      * Only m5/m6/c5/c6 families qualify (M5<->M6, C5<->C6 interchangeable);
        Intel 'i' variants are excluded and stay at 8%.
      * A qualifying instance is matched to a BoM line by exact (OS, vCPU, GB).
        Anything not listed falls to 8% per the Section-D note.
    Returns (bom_key_or_None, multiplier, anchor_desc). multiplier is always 1.0
    (no cross-size scaling — unlisted sizes go to 8%).
    """
    os_norm = "win" if os_cls == "win" else "rhel"   # all Linux -> RHEL
    if not _bom_eligible(inst):
        return None, 1.0, None
    key = EC2_EXACT.get((os_norm, vcpu, mem))
    if key:
        return key, 1.0, None
    return None, 1.0, None


# ProductName strings that already have a dedicated builder below. The catch-all
# (_other_services) skips these and bills every OTHER ProductName found in the
# CUR at the non-BoM 8% rate, so a service the engine doesn't explicitly model
# is no longer silently dropped from the sheet (v3.1).
HANDLED_PRODUCTS = {
    "Amazon Elastic Compute Cloud",
    "Amazon Simple Storage Service",
    "Amazon Relational Database Service",
    "Amazon Virtual Private Cloud",
    "AWS Direct Connect",
    "Elastic Load Balancing",
    "AWS Network Firewall",
    "Amazon Elastic Container Service",
    "Amazon EC2 Container Registry (ECR)",
    "Amazon Elastic Container Service for Kubernetes",
    "Amazon OpenSearch Service",
    "AWS Key Management Service",
    "AWS WAF",
    "AmazonCloudWatch",
    "AWS Config",
    "AWS CloudTrail",
    "AWS Data Transfer",
    "Amazon GuardDuty",
    "AWS Cost Explorer",
    "AWS Security Hub",
}


class Pricer:
    def __init__(self, cur_df, invoice, bom: dict, rate: float):
        self.li      = cur_df    # already filtered to allowed accounts DataFrame
        self.inv     = invoice
        self.bom     = bom
        self.rate    = rate
        self._note_n = 0

    # ── helpers ─────────────────────────────────────────────────────────────
    def _next_note(self) -> int:
        self._note_n += 1
        return self._note_n

    def _bom_formula(self, key: str, consumption: float, multiplier: float = 1.0) -> str:
        b = self.bom[key]
        unit = b["unit"] * multiplier
        basis = b["basis"]
        if basis == 1:
            return f"=(({unit:.6f}*H{{r}}))/{self.rate}"
        return f"=((({unit:.6f}/{basis})*{round(consumption, 4)})*H{{r}})/{self.rate}"

    def _inv_charge(self, aid: str, service: str) -> float:
        return self.inv.charges(aid, service) if self.inv else 0.0

    def _f(self, aid: str, prod: str) -> "pd.DataFrame":
        """
        Per-account service frame.
        Point 3: zero-cost free-tier lines are dropped here so they never
        reach the sheet — neither as their own row nor as consumption feeding
        a BoM quantity (fixes the 850 GB / 1044 GB free-tier lines in Case 2).
        """
        frame = CUR.svc(self.li[self.li["_acct"] == aid], prod)
        return frame[frame["CostBeforeTax"] > 0]

    @staticmethod
    def _instance_specs(inst_type: str) -> tuple[int, int]:
        return instance_specs(inst_type)

    # ── main entry point ────────────────────────────────────────────────────
    def price_accounts(self, account_list: list[tuple[str, str]]) -> PricingResult:
        res = PricingResult()
        self._build_all(account_list, res)
        self._grossup_nonbom(res, account_list)
        return res

    def _grossup_nonbom(self, res, account_list):
        """
        Non-BoM rows must be priced from the invoice's GROSS charges, then have
        the 8% applied (per the rule book). The CUR's CostBeforeTax is already
        NET of the AWS Distribution Program Discount (~12%), so pricing straight
        from it double-discounts and under-bills every non-BoM row. Here we
        scale each non-BoM row's indicative USD value back up to gross using the
        account's own gross/net ratio (invoice gross ÷ CUR net). BoM rows are
        untouched.
        """
        import re as _re
        num = _re.compile(r"^=(-?\d+(?:\.\d+)?)$")
        for aid, _ in account_list:
            gross = self.inv.total_charges(aid) if self.inv else 0.0
            net   = float(self.li[self.li["_acct"] == aid]["CostBeforeTax"].sum())
            if gross <= 0 or net <= 1e-6:
                continue
            factor = gross / net
            if abs(factor - 1.0) < 1e-9:
                continue
            for _sn, _svc, _is_bom, row in res.rows:
                if getattr(row, "is_bom", False) or getattr(row, "_aid", None) != aid:
                    continue
                m = num.match(row.i_formula.strip())
                if m:
                    row.i_formula = f"={round(float(m.group(1)) * factor, 4)}"

    def _build_all(self, account_list, res):
        """Build all service groups in canonical order."""
        sn = [1]  # mutable counter

        def add_service(svc_name: str, is_bom: bool, rows: list):
            if not rows:
                return
            for row in rows:
                res.rows.append((sn[0], svc_name, is_bom, row))
            sn[0] += 1

        # 1. Compute — split into Windows OS / RHEL OS groups (by row label),
        #    each row already carries is_bom, so mixed BoM+8% rows coexist.
        comp_rows = self._compute(account_list, res)
        comp_grouped: dict[str, list] = {}
        comp_order: list[str] = []
        for row in comp_rows:
            if row.service not in comp_grouped:
                comp_grouped[row.service] = []
                comp_order.append(row.service)
            comp_grouped[row.service].append(row)
        for svc in comp_order:
            grp = comp_grouped[svc]
            add_service(svc, any(r.is_bom for r in grp), grp)
        # 2. Block SSD gp3 (BoM slab)
        add_service("Block SSD Storage", True,
                    self._block_ssd(account_list))
        # 3. EBS Snapshot (separate, non-BoM)
        add_service("EBS Snapshot Storage", False,
                    self._snapshot(account_list))
        # 4. Object Storage S3 (BoM)
        add_service("Object Storage", True,
                    self._s3(account_list))
        # 5. Managed Database RDS (non-BoM)
        add_service("Managed Database", False,
                    self._rds(account_list))
        # 6. Network Data Transfer (BoM, 0 egress)
        add_service("Network Data Transfer", True,
                    self._net_xfer(account_list))
        # 7. NAT Gateway Hours (non-BoM)
        add_service("NAT Gateway – Hours", False,
                    self._nat_hours(account_list))
        # 8. NAT Gateway Data (BoM)
        add_service("NAT Gateway – Data Processed", True,
                    self._nat_data(account_list))
        # 9. Static Public IP (BoM, folds all VPC IPv4)
        add_service("Static Public IP", True,
                    self._static_ip(account_list))
        # 10. VPC Other (TGW + Encryption + Endpoints, non-BoM)
        add_service("VPC – Other Charges", False,
                    self._vpc_other(account_list))
        # 11. Direct Connect (BoM leased-line, disc 75%) — point 2
        add_service("Direct Connect", True,
                    self._direct_connect(account_list, res))
        # 11b. Site-to-Site VPN (BoM line 35, disc 78%)
        add_service("VPN Connection", True,
                    self._vpn(account_list))
        # 12. Application Load Balancing (BoM)
        add_service("Application load balancing", True,
                    self._alb(account_list))
        # 13. Network Firewall (BoM line37 @60%)
        add_service("Network Firewall", True,
                    self._nfw(account_list))
        # 14. ECS Fargate (non-BoM)
        add_service("Elastic Container Service (Fargate)", False,
                    self._ecs(account_list))
        # 15. ECR (non-BoM)
        add_service("Elastic Container Registry (ECR)", False,
                    self._ecr(account_list))
        # 16. EKS (non-BoM)
        add_service("Elastic Container Service for Kubernetes", False,
                    self._eks(account_list))
        # 17. OpenSearch (non-BoM)
        add_service("Amazon OpenSearch Service", False,
                    self._opensearch(account_list))
        # 18. KMS (BoM line39 @99%)
        add_service("Key Management Service", True,
                    self._kms(account_list))
        # 19. WAF (BoM, may be 0)
        add_service("WAF", True,
                    self._waf(account_list))
        # 20. CloudWatch (non-BoM)
        add_service("Cloudwatch", False,
                    self._cloudwatch(account_list))
        # 21. Config (non-BoM)
        add_service("Config", False,
                    self._config(account_list))
        # 22. CloudTrail (non-BoM)
        add_service("CloudTrail", False,
                    self._cloudtrail(account_list))
        # 23. Data Transfer (non-BoM, only if billable)
        add_service("Data Transfer", False,
                    self._data_transfer(account_list))
        # 24. GuardDuty (non-BoM, if present)
        add_service("GuardDuty", False,
                    self._guardduty(account_list))
        # 25. Cost Explorer (non-BoM, if present)
        add_service("Cost Explorer", False,
                    self._cost_explorer(account_list))
        # 26. Security Hub (non-BoM, if present)
        add_service("Security Hub", False,
                    self._security_hub(account_list))
        # 27. Catch-all — any other billable service not modelled above, non-BoM
        #     @8%, so nothing is silently dropped. Each distinct ProductName
        #     becomes its own service group (preserving CUR order).
        other_rows = self._other_services(account_list)
        grouped: dict[str, list] = {}
        order: list[str] = []
        for row in other_rows:
            if row.service not in grouped:
                grouped[row.service] = []
                order.append(row.service)
            grouped[row.service].append(row)
        for svc in order:
            add_service(svc, False, grouped[svc])

    # ── per-service builders ─────────────────────────────────────────────────

    def _compute(self, accts, res) -> list:
        """
        One row per (account, instance-type). Case-1 fixes applied:
          * each account's rows are tagged to THAT account, immediately;
          * CPU credits are billed whenever present.
        """
        all_rows = []
        for aid, aname in accts:
            acct_rows = []
            f = self._f(aid, "Amazon Elastic Compute Cloud")   # cost>0 only
            boxes = f[f["UsageType"].str.contains("BoxUsage", na=False)]
            cpu   = f[f["UsageType"].str.contains("CPUCredits", na=False)]

            for inst_type, grp in boxes.groupby(
                boxes["UsageType"].str.extract(r"BoxUsage:(\S+)")[0]
            ):
                if not inst_type or float(grp["CostBeforeTax"].sum()) < 1e-6:
                    continue
                hours   = float(grp["UsageQuantity"].sum())
                cbt_box = float(grp["CostBeforeTax"].sum())
                desc_sample = str(grp["ItemDescription"].iloc[0]) if len(grp) else ""
                os_cls = _os_class(desc_sample)
                vcpu, mem = self._instance_specs(inst_type)
                bkey, mult, anchor_desc = _ec2_pricing(
                    inst_type, vcpu, mem, os_cls, self.bom)

                if bkey:
                    b = self.bom[bkey]
                    nn = self._next_note()
                    bom_inst = b.get("inst", bkey)
                    res.working_notes.append((nn,
                        f'Actual instance "{inst_type}" ({os_cls.upper()}, {vcpu} vCPU / {mem} GB) '
                        f'priced using BoM "{bom_inst}" (line {b["line"]}, ₹{b["unit"]:,.2f})'
                        + (f' × {mult:.2f} vCPU ratio' if mult != 1.0 else ' — exact match')
                        + f'. Discount {b["disc"]:.2%} unchanged.'
                    ))
                    formula = self._bom_formula(bkey, 1, mult)
                    os_label = "Windows" if os_cls == "win" else "Red Hat Enterprise Linux"
                    os_svc   = "Compute services – Windows OS" if os_cls == "win" else "Compute services – RHEL OS"
                    acct_rows.append(Row(
                        service=os_svc,
                        additional=f"On-Demand {os_label} instance (BoM substitution - Note {nn})",
                        config=f"Chosen instance: {inst_type}  |  {vcpu} vCPU  |  {mem} GiB Memory  |  {hours:.2f} hrs",
                        sku=f"Chosen instance: {inst_type}  |  Family:{_base_family(inst_type)}  |  {vcpu}vCPU  |  {mem} GiB Memory",
                        qty=1, i_formula=formula, discount=b["disc"], is_bom=True, note_n=nn,
                    ))
                else:
                    os_label = "Windows" if os_cls == "win" else "Linux/Unix"
                    os_svc   = "Compute services – Windows OS" if os_cls == "win" else "Compute services – RHEL OS"
                    inst_desc = str(grp["ItemDescription"].iloc[0])[:65] if len(grp) else ""
                    acct_rows.append(Row(
                        service=os_svc,
                        additional=f"1. {os_label}",
                        config=f"Chosen instance: {inst_type}  |  {vcpu} vCPU  |  {mem} GiB Memory  |  {hours:.2f} hrs",
                        sku=f"APS3-BoxUsage:{inst_type} - {inst_desc} ({hours:.2f} hrs)",
                        qty=1, i_formula=f"={round(cbt_box, 4)}",
                        discount=NON_BOM_DISC, is_bom=False,
                    ))

            # CPU credits — always billed when present (Case-1 fix).
            cpu_cbt = float(cpu["CostBeforeTax"].sum())
            if cpu_cbt > 1e-6:
                acct_rows.append(Row(
                    service="Compute services – Linux OS",
                    additional="T3/T3A CPU Credits",
                    config=f"{CUR.qty(f, 'CPUCredits'):.2f} vCPU-hrs",
                    sku=f"APS3-CPUCredits - {CUR.desc(f, 'CPUCredits')} ({CUR.qty(f, 'CPUCredits'):.2f})",
                    qty=1, i_formula=f"={round(cpu_cbt, 4)}",
                    discount=NON_BOM_DISC, is_bom=False,
                ))

            # Tag THIS account's rows to THIS account (Case-1 fix).
            for row in acct_rows:
                row._aid = aid
                row._aname = aname
            all_rows.extend(acct_rows)
        return all_rows

    def _block_ssd(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Elastic Compute Cloud")
            gp3 = CUR.qty(f, "VolumeUsage.gp3")
            if gp3 < 1e-6:
                continue
            key = ebs_key(gp3)
            b = self.bom[key]
            desc = CUR.desc(f, "VolumeUsage.gp3")
            row = Row(
                service="Block SSD Storage",
                additional="EBS SSD block storage (gp3)",
                config=f"{gp3:.3f} GB Disk (gp3)",
                sku=f"APS3-EBS:VolumeUsage.gp3 - {desc} ({gp3:.3f} GB)",
                qty=1, i_formula=self._bom_formula(key, gp3),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _snapshot(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Elastic Compute Cloud")
            snap = CUR.qty(f, "SnapshotUsage")
            cbt  = CUR.cbt(f, "SnapshotUsage")
            if snap < 1e-6 or cbt < 1e-6:
                continue
            desc = CUR.desc(f, "SnapshotUsage")
            row = Row(
                service="EBS Snapshot Storage",
                additional="EBS snapshot storage",
                config=f"{snap:.3f} GB-month snapshot",
                sku=f"APS3-EBS:SnapshotUsage - {desc} ({snap:.3f} GB)",
                qty=1, i_formula=f"={round(cbt, 4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _s3(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Simple Storage Service")
            if CUR.cbt(f) < 1e-6:
                continue
            s3gb = CUR.qty(f, "TimedStorage")
            req1 = CUR.qty(f, "Tier1")
            req2 = CUR.qty(f, "Tier2")
            b = self.bom["s3_hot"]
            row = Row(
                service="Object Storage",
                additional="Hot Tier",
                config=f"{s3gb:.4f} GB",
                sku=(f"PUT, COPY, POST, LIST requests to S3 Standard ({req1:.0f} Requests), "
                     f"GET, SELECT, and all other requests from S3 Standard ({req2:.0f} Requests)"),
                qty=1, i_formula=self._bom_formula("s3_hot", s3gb),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _rds(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Relational Database Service")
            total = CUR.cbt(f)
            if total < 1e-6:
                continue
            inst_f  = f[f["UsageType"].str.contains("InstanceUsage", na=False)]
            stor_gb = CUR.qty(f, "GP3-Storage")
            back_gb = CUR.qty(f, "ChargedBackupUsage")
            inst_cfg = "; ".join(sorted(set(
                f"{r['UsageType'].split(':')[-1]} ({r['UsageQuantity']:.2f} hrs)"
                for _, r in inst_f[inst_f["CostBeforeTax"] > 0].iterrows()
            )))
            row = Row(
                service="Managed Database",
                additional="Backup Storage",
                config=f"{inst_cfg}; GP3 storage {stor_gb:.2f} GB; backup {back_gb:.3f} GB",
                sku="\n".join(CUR.skus(f)),
                qty=1, i_formula=f"={round(total, 4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _net_xfer(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            if CUR.cbt(self._f(aid, "Amazon Elastic Compute Cloud")) < 1e-6:
                continue
            b = self.bom["net_xfer"]
            row = Row(
                service="Network Data Transfer",
                additional="Egress Cost",
                config="\t\n0 GB/month",
                sku="Regional data transfer - in/out/between EC2 AZs or using elastic IPs or ELB\n0 GB",
                qty=1, i_formula=self._bom_formula("net_xfer", 0),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _nat_hours(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Elastic Compute Cloud")
            hrs = CUR.qty(f, "NatGateway-Hours")
            cbt = CUR.cbt(f, "NatGateway-Hours")
            if hrs < 1:
                continue
            desc = CUR.desc(f, "NatGateway-Hours")
            row = Row(
                service="NAT Gateway – Hours",
                additional="Managed NAT Gateway",
                config=f"Number of NAT Gateways (1),NAT Gateway Hour\n{hrs:.0f} Hrs",
                sku=f"APS3-RegionalNatGateway-Hours - {desc} ({hrs:.0f} Hrs)",
                qty=1, i_formula=f"={round(cbt, 4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _nat_data(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Elastic Compute Cloud")
            gb = CUR.qty(f, "NatGateway-Bytes")
            if gb < 1e-6:
                continue
            desc = CUR.desc(f, "NatGateway-Bytes")
            b = self.bom["nat_data"]
            row = Row(
                service="NAT Gateway – Data Processed",
                additional="Managed NAT Gateway",
                config=f"{gb:.4f} GB processed",
                sku=f"APS3-RegionalNatGateway-Bytes - {desc} ({gb:.4f} GB)",
                qty=1, i_formula=self._bom_formula("nat_data", gb),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _static_ip(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Virtual Private Cloud")
            inuse = CUR.qty(f, "InUseAddress")
            idle  = CUR.qty(f, "IdleAddress")
            total_hrs = inuse + idle
            n_ip = round(total_hrs / ADDR_HRS_PER_MONTH)
            if n_ip < 1:
                continue
            b = self.bom["static_ip"]
            row = Row(
                service="Static Public IP",
                additional="Static Public IPs",
                config=f"Number of In-use public IPv4 addresses ({n_ip})",
                sku=(f"APS3-PublicIPv4:InUseAddress - $0.005 per In-use public IPv4 address per hour ({inuse:.2f} hrs)\n"
                     f"APS3-PublicIPv4:IdleAddress - $0.005 per Idle public IPv4 address per hour ({idle:.2f} hrs)"),
                qty=n_ip, i_formula=self._bom_formula("static_ip", 1),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _vpc_other(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Virtual Private Cloud")
            tgw_h   = CUR.qty(f, "TransitGateway-Hours")
            tgw_gb  = CUR.qty(f, "TransitGateway-Bytes")
            enc_h   = CUR.qty(f, "VPCEncryptionControls")
            ep_h    = CUR.qty(f, "VpcEndpoint-Hours")
            total   = round(CUR.cbt(f, "TransitGateway") +
                            CUR.cbt(f, "VPCEncryptionControls") +
                            CUR.cbt(f, "VpcEndpoint"), 4)
            if total < 1e-6:
                continue
            skus = []
            if tgw_h > 0:
                skus.append(f"APS3-TransitGateway-Hours - $0.07 per Transit Gateway VPC Attachment Hour ({tgw_h:.0f} hrs)")
            if tgw_gb > 0:
                skus.append(f"APS3-TransitGateway-Bytes - $0.02 per GB Data Processed by Transit Gateway ({tgw_gb:.4f} GB)")
            if enc_h > 0:
                skus.append(f"APS3-VPCEncryptionControls-hours - $0.19 per hour for VPCEncryptionControls-hours ({enc_h:.0f} hrs)")
            if ep_h > 0:
                skus.append(f"APS3-VpcEndpoint-Hours - $0.013 per VPC Endpoint Hour ({ep_h:.0f} hrs)")
            row = Row(
                service="VPC – Other Charges",
                additional="Transit Gateway + VPC Encryption Controls + VPC Endpoints",
                config=f"Transit Gateway: {tgw_h:.0f} Hrs | VPC Encryption Controls: {enc_h:.0f} Hrs | VPC Endpoints: {ep_h:.0f} Hrs",
                sku="\n".join(skus),
                qty=1, i_formula=f"={total}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _direct_connect(self, accts, res) -> list:
        """
        Point 2 — Direct Connect priced from the BoM leased-line rate.
        Costing: (((19912.5/1000) * actual_mbps) * quantity) / rate ; disc 75%.
        Mbps is parsed from the CUR ItemDescription/UsageType (e.g. "200Mbps",
        "1G", "10Gbps"); quantity = number of connections (port-hours / 730).
        """
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS Direct Connect")
            if CUR.cbt(f) < 1e-6:
                continue
            port = f[f["UsageType"].str.contains("Port", case=False, na=False)]
            port_hrs = float(port["UsageQuantity"].sum())
            n_conn = max(1, round(port_hrs / ADDR_HRS_PER_MONTH)) if port_hrs > 0 else 1

            mbps, speed_src = self._parse_dx_mbps(f)
            nn = self._next_note()
            assumed = " (assumed 1 Gbps — port speed not found in CUR)" if speed_src == "default" else ""
            res.working_notes.append((nn,
                f'Direct Connect priced from BoM line 40 "Dedicated DC-Cloud connect '
                f'leased line charges" (₹19,912.5 per 1 Gbps). Actual link {mbps:.0f} Mbps'
                f'{assumed}, {n_conn} connection(s). Costing = (((19912.5/1000)*{mbps:.0f})*{n_conn})/rate, '
                f'discount 75%.'
            ))
            b = self.bom["direct_connect"]
            row = Row(
                service="Direct Connect",
                additional="Dedicated DC-Cloud connect leased line",
                config=f"{mbps:.0f} Mbps leased line × {n_conn} connection(s)",
                sku="\n".join(CUR.skus(f)) or f"AWS Direct Connect port ({port_hrs:.0f} hrs)",
                qty=n_conn,
                i_formula=self._bom_formula("direct_connect", mbps),
                discount=b["disc"], is_bom=True, note_n=nn,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    @staticmethod
    def _parse_dx_mbps(frame) -> tuple[float, str]:
        """Extract link speed in Mbps from Direct Connect descriptions."""
        import re as _re
        for _, r in frame.iterrows():
            for field_ in (str(r.get("ItemDescription", "")), str(r.get("UsageType", ""))):
                m = _re.search(r"(\d+(?:\.\d+)?)\s*(gbps|g\b|mbps|m\b)", field_, _re.I)
                if m:
                    val = float(m.group(1))
                    unit = m.group(2).lower()
                    return (val * 1000 if unit.startswith("g") else val), "parsed"
        return 1000.0, "default"

    def _vpn(self, accts) -> list:
        """
        Site-to-Site VPN — BoM line 35 (₹10,035 per connection per month, 78%).
        VPN usage sits under "Amazon Virtual Private Cloud" as
        APS3-VPN-Usage-Hours; number of connections = round(hours / 730), min 1.
        """
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Virtual Private Cloud")
            vpn = f[f["UsageType"].str.contains("VPN-Usage", case=False, na=False)]
            hrs = float(vpn["UsageQuantity"].sum())
            if hrs < 1 or float(vpn["CostBeforeTax"].sum()) < 1e-6:
                continue
            n_conn = max(1, round(hrs / ADDR_HRS_PER_MONTH))
            b = self.bom["vpn"]
            row = Row(
                service="VPN Connection",
                additional="Site-to-Site VPN",
                config=f"VPN Connection per month\n{hrs:.0f} connection-hrs → {n_conn} connection(s)",
                sku=f"Number of Site-to-Site VPN Connections ({n_conn}), APS3-VPN-Usage-Hours ({hrs:.0f} hrs)",
                qty=n_conn, i_formula=self._bom_formula("vpn", 1),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _alb(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Elastic Load Balancing")
            alb_h = float(f[f["CostBeforeTax"] > 0]["UsageQuantity"].sum())
            if alb_h < 1:
                continue
            descs = "; ".join(sorted(set(
                str(r["ItemDescription"])[:55].strip()
                for _, r in f[f["CostBeforeTax"] > 0].iterrows()
            )))
            b = self.bom["app_lb"]
            row = Row(
                service="Application load balancing",
                additional="L7 load balancer with HA functionality to host production workloads",
                config=f"Application LoadBalancer-hour\n{alb_h:.0f} Hrs",
                sku=f"Number of Application Load Balancers (1)\nAPS3-LoadBalancerUsage - {descs} ({alb_h:.0f} Hrs)",
                qty=1, i_formula=self._bom_formula("app_lb", alb_h),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _nfw(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS Network Firewall")
            ep_h  = CUR.qty(f, "Endpoint-Hour")
            if ep_h < 1:
                continue
            # Client-approved sheet prices per firewall ENDPOINT (each running
            # endpoint = 1 BoM unit), not per pair of endpoints. e.g. 1488 hrs
            # /730 = 2 endpoints -> qty 2.
            nfw_qty = max(1, round(ep_h / ADDR_HRS_PER_MONTH))
            traf    = CUR.qty(f, "Traffic-GB")
            ep_desc = CUR.desc(f, "Endpoint-Hour")
            tr_desc = CUR.desc(f, "Traffic-GB")
            b = self.bom["net_fw"]
            row = Row(
                service="Network Firewall",
                additional="Network Firewall",
                config=f"{nfw_qty} network firewall endpoint(s)\n{ep_h:.2f} endpoint-hrs → {nfw_qty} endpoint(s)\n{traf:.4f} GB traffic processed",
                sku=(f"APS3-Endpoint-Hour - {ep_desc} ({ep_h:.2f} hrs → {nfw_qty} endpoint(s))\n"
                     f"APS3-Traffic-GB-Processed - {tr_desc} ({traf:.4f} GB)"),
                qty=nfw_qty, i_formula=self._bom_formula("net_fw", 1),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _ecs(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Elastic Container Service")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            vcpu = CUR.qty(f, "vCPU-Hours"); mem = CUR.qty(f, "Fargate-GB-Hours")
            row = Row(
                service="Elastic Container Service (Fargate)",
                additional="AWS Fargate",
                config=f"{vcpu:.2f} vCPU-hrs + {mem:.2f} GB-hrs, Asia Pacific (Mumbai)",
                sku="\n".join(CUR.skus(f)),
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _ecr(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon EC2 Container Registry (ECR)")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            gb = float(f["UsageQuantity"].sum())
            desc = CUR.desc(f, "TimedStorage")
            row = Row(
                service="Elastic Container Registry (ECR)",
                additional="Container image storage",
                config=f"{gb:.3f} GB-month",
                sku=f"APS3-TimedStorage-ByteHrs - {desc} ({gb:.3f} GB)",
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _eks(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon Elastic Container Service for Kubernetes")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            hrs = CUR.qty(f, "perCluster")
            desc = CUR.desc(f, "perCluster")
            row = Row(
                service="Elastic Container Service for Kubernetes",
                additional="EKS cluster",
                config=f"{hrs:.0f} Hrs, Asia Pacific (Mumbai)",
                sku=f"APS3-AmazonEKS-Hours:perCluster - {desc} ({hrs:.0f} Hrs)",
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _opensearch(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon OpenSearch Service")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            inst_h = CUR.qty(f, "ESInstance"); stor_gb = CUR.qty(f, "GP3-Storage")
            row = Row(
                service="Amazon OpenSearch Service",
                additional="OpenSearch cluster",
                config=f"{inst_h:.0f} t3.medium.search hrs + {stor_gb:.2f} GB GP3 storage",
                sku="\n".join(CUR.skus(f)),
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _kms(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS Key Management Service")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            keys = CUR.qty(f, "KMS-Keys"); reqs = CUR.qty(f, "KMS-Requests")
            b = self.bom["kms"]
            row = Row(
                service="Key Management Service",
                additional="50 Encryption keys + 10000 Symmetric Encryption operations/month",
                config=f"{keys:.2f} KMS keys + {reqs:.0f} requests",
                sku="\n".join(CUR.skus(f)),
                qty=1, i_formula=self._bom_formula("kms", 1),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _waf(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS WAF")
            req = CUR.qty(f, "RequestV2")
            if CUR.cbt(f) < 1e-6 and req < 1: continue
            b = self.bom["waf"]
            row = Row(
                service="WAF",
                additional="Highly Available WAF for production workloads",
                config=f"requests processed  \t\n{req:.0f} Request Web ACL Rules(6)\n",
                sku=f"requests processed  \t\n{req:.0f} Request, Web ACL Rules(6)\n",
                qty=1, i_formula=self._bom_formula("waf", req),
                discount=b["disc"], is_bom=True,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _cloudwatch(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AmazonCloudWatch")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            gb = CUR.qty(f, "DataProcessing-Bytes") + CUR.qty(f, "VendedLog-Bytes")
            row = Row(
                service="Cloudwatch",
                additional="Vended log ingestion + storage",
                config=f"vended logs ingested in Standard log class {gb:.4f} GB",
                sku="\n".join(CUR.skus(f)),
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _config(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS Config")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            items = CUR.qty(f, "ConfigurationItemRecorded")
            row = Row(
                service="Config",
                additional=None,
                config=f"ConfigurationItemRecorded\n{items:.0f} APS3-ConfigurationItemRecorded",
                sku=f"\t\n{items:.0f} APS3-ConfigurationItemRecorded",
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _cloudtrail(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS CloudTrail")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            insights = CUR.qty(f, "InsightsEvents")
            row = Row(
                service="CloudTrail",
                additional=None,
                config="CloudTrail APS3-InsightsEvents",
                sku=f"CloudTrail APS3-InsightsEvents {insights:.0f} Events",
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _data_transfer(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS Data Transfer")
            cbt = CUR.cbt(f)
            if cbt < 0.001: continue
            reg_gb = CUR.qty(f, "Regional-Bytes")
            row = Row(
                service="Data Transfer",
                additional="Egress Cost",
                config=f"{reg_gb:.4f} GB regional data transfer",
                sku=f"APS3-DataTransfer-Regional-Bytes - $0.010 per GB regional data transfer ({reg_gb:.4f} GB)",
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _guardduty(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "Amazon GuardDuty")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            events = CUR.qty(f, "PaidEventsAnalyzed")
            row = Row(
                service="GuardDuty",
                additional="GuardDuty CloudTrail event analysis",
                config=f"{events:.0f} events analysed",
                sku="\n".join(CUR.skus(f)),
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _cost_explorer(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS Cost Explorer")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            reqs = CUR.qty(f, "APIRequest")
            row = Row(
                service="Cost Explorer",
                additional="Cost Explorer API requests",
                config=f"{reqs:.0f} API requests",
                sku=f"USE1-APIRequest - $0.01 per Cost Explorer API request ({reqs:.0f})",
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _security_hub(self, accts) -> list:
        rows = []
        for aid, aname in accts:
            f = self._f(aid, "AWS Security Hub")
            cbt = CUR.cbt(f)
            if cbt < 1e-6: continue
            row = Row(
                service="Security Hub",
                additional="AWS Security Hub",
                config="Security Hub findings",
                sku="\n".join(CUR.skus(f)),
                qty=1, i_formula=f"={round(cbt,4)}",
                discount=NON_BOM_DISC, is_bom=False,
            )
            row._aid = aid; row._aname = aname
            rows.append(row)
        return rows

    def _other_services(self, accts) -> list:
        """
        Catch-all (v3.1). Every ProductName that has no dedicated builder above
        is billed here at the non-BoM 8% rate — one row per (account, service) —
        so a service the engine doesn't explicitly model is never silently
        dropped. Zero-cost / free-tier lines stay excluded (CostBeforeTax > 0),
        consistent with R4.

        Scope: this recovers WHOLE unmodelled services (e.g. Route 53,
        CloudFront, Lambda, SNS, SQS, Secrets Manager, EFS, DynamoDB, ...). It
        deliberately does NOT touch usage types inside already-handled products
        (VPN under "Amazon Virtual Private Cloud", RDS BoM substitution, Transit
        Gateway as BoM) — those need the Section B BoM rates and are priced by
        their own builders once those BoM lines are supplied. Applied on every
        tab.
        """
        rows = []
        for aid, aname in accts:
            frame = self.li[self.li["_acct"] == aid]
            frame = frame[frame["CostBeforeTax"] > 0]
            if frame.empty:
                continue
            for prod, grp in frame.groupby("ProductName"):
                prod = str(prod).strip()
                if not prod or prod == "nan" or prod in HANDLED_PRODUCTS:
                    continue
                cbt = float(grp["CostBeforeTax"].sum())
                if cbt < 1e-6:
                    continue
                usage = "; ".join(sorted(set(
                    str(u).split(":")[-1] for u in grp["UsageType"].dropna()
                    if str(u).strip()
                )))[:120]
                row = Row(
                    service=prod,
                    additional="Billed at 8% off public pricing (no BoM line configured for this service)",
                    config=usage or "See SKU",
                    sku="\n".join(CUR.skus(grp)),
                    qty=1, i_formula=f"={round(cbt, 4)}",
                    discount=NON_BOM_DISC, is_bom=False,
                )
                row._aid = aid; row._aname = aname
                rows.append(row)
        return rows
