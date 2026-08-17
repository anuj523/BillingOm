"""
Invoice PDF parser — extracts conversion rate + per-account Charges.
Used only for non-BoM row pricing (I-column = invoice gross Charges).
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

_RATE_RE    = re.compile(r"conversion rate of\s*([\d,]+\.\d+)", re.I)
_PERIOD_RE  = re.compile(r"billing period\s+(.+?\d{4})", re.I)
_ACCT_RE    = re.compile(r"\((\d{12})\)")
_SVC_RE     = re.compile(r"^(.*?)\s+USD\s+([\d,]+\.\d{2})\s*$")
_CHARGES_RE = re.compile(r"^Charges\s+USD\s+([\d,]+\.\d{2})\s*$", re.I)


def _n(s: str) -> float:
    return float(str(s).replace(",", ""))


@dataclass
class AccountInvoice:
    account_id: str
    account_name: str
    services: dict = field(default_factory=dict)

    @property
    def charges_total(self) -> float:
        return round(sum(self.services.values()), 2)


@dataclass
class Invoice:
    conversion_rate: float
    billing_period: str
    accounts: dict = field(default_factory=dict)  # {aid: AccountInvoice}

    def charges(self, aid: str, service: str) -> float:
        a = self.accounts.get(aid)
        return a.services.get(service, 0.0) if a else 0.0

    def total_charges(self, aid: str) -> float:
        a = self.accounts.get(aid)
        return a.charges_total if a else 0.0


def parse(path_or_buffer, allowed_ids: set) -> Invoice:
    if pdfplumber is None:
        raise RuntimeError("pip install pdfplumber")

    lines = []
    with pdfplumber.open(path_or_buffer) as pdf:
        for page in pdf.pages:
            lines.extend((page.extract_text() or "").split("\n"))

    rate, period = None, ""
    for ln in lines:
        if rate is None:
            m = _RATE_RE.search(ln)
            if m:
                rate = _n(m.group(1))
        if not period:
            m = _PERIOD_RE.search(ln)
            if m:
                period = m.group(1).strip()

    if rate is None:
        raise ValueError("Conversion rate not found in invoice PDF.")

    accounts: dict[str, AccountInvoice] = {}
    current: AccountInvoice | None = None
    pending_svc: str | None = None

    for ln in lines:
        ln = ln.strip()
        hdr = _ACCT_RE.search(ln)
        if hdr:
            aid = hdr.group(1)
            if aid in allowed_ids:
                current = accounts.setdefault(aid, AccountInvoice(aid, aid))
                pending_svc = None
            else:
                current = None
            continue
        if current is None:
            continue
        ms = _SVC_RE.match(ln)
        if ms and not any(x in ln for x in ("Charges", "Discount", "GST", "Credits", "total")):
            pending_svc = ms.group(1).strip()
            continue
        mc = _CHARGES_RE.match(ln)
        if mc and pending_svc:
            current.services[pending_svc] = (
                current.services.get(pending_svc, 0.0) + _n(mc.group(1))
            )
            pending_svc = None

    return Invoice(conversion_rate=rate, billing_period=period, accounts=accounts)
