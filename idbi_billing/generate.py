"""
Top-level orchestration for billing sheet generation.
Supports DBD accounts (with invoice PDF) and ITD accounts (with invoice PDF).
"""
from __future__ import annotations
import os, re
from dataclasses import dataclass

from .cur     import CUR
from .invoice import parse as parse_invoice
from .bom     import load  as load_bom
from .pricing import Pricer
from .render  import write
from .config  import DBD_ACCOUNTS, ITD_ACCOUNTS, NOTE_TEXT


@dataclass
class SheetReport:
    account_group: str
    month_label:   str
    rate:          float
    n_rows:        int
    n_accounts:    int
    out_path:      str


def _month_parts(period: str) -> tuple[str, str]:
    """'February 1 - February 28 , 2026' -> ('Feb-2026', '01st February 2026')"""
    m = re.search(r"([A-Za-z]+)\s+\d+.*?(\d{4})", period)
    if not m:
        return "Bill", "01st of month"
    month, year = m.group(1), m.group(2)
    abbr = month[:3] if month.lower() != "september" else "Sept"
    return f"{abbr}-{year}", f"01st {month} {year}"


def build_billing(
    cur_path,
    invoice_path,
    bom_path,
    out_dir:   str,
    group:     str = "dbd",      # "dbd" or "itd"
    date_override: str | None = None,
) -> SheetReport:
    accounts = DBD_ACCOUNTS if group == "dbd" else ITD_ACCOUNTS
    bom      = load_bom(bom_path)
    cur      = CUR.load(cur_path)
    inv      = parse_invoice(invoice_path, set(accounts.keys()))
    rate     = inv.conversion_rate
    month_label, auto_date = _month_parts(inv.billing_period)
    date_label = date_override or auto_date

    li = cur.line_items(set(accounts.keys()))
    active = cur.active_accounts(accounts)

    pricer = Pricer(li, inv, bom, rate)
    result = pricer.price_accounts(active)

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"IDBI_{group.upper()}_AWS_Bill_{month_label}.xlsx")

    write(
        result.rows,
        rate=rate,
        month_label=month_label,
        date_label=date_label,
        working_notes=result.working_notes,
        out_path=out_path,
        sheet_title=month_label,
        footer_note=NOTE_TEXT.format(month=month_label),
        account_group=group.upper(),
    )

    return SheetReport(
        account_group=group.upper(),
        month_label=month_label,
        rate=rate,
        n_rows=len(result.rows),
        n_accounts=len(active),
        out_path=out_path,
    )
