"""
CUR (Cost & Usage Report) CSV parsing.
Filters to LinkedLineItem records for the relevant account group only.
CUR is used exclusively for quantity/config/SKU derivation — NOT for pricing.
"""
from __future__ import annotations
import pandas as pd


def read_tabular(path_or_buffer, skiprows=0, dtype=None) -> pd.DataFrame:
    """
    Read a CSV *or* Excel file into a DataFrame, auto-detecting the format.
    Detection uses the filename/extension when available, and falls back to a
    content sniff (Excel files start with the ZIP magic 'PK') for buffers with
    no usable name. Lets every input accept both .csv and .xlsx/.xls.
    """
    name = str(getattr(path_or_buffer, "name", path_or_buffer) or "").lower()
    is_excel = name.endswith((".xlsx", ".xls", ".xlsm"))
    if not is_excel and not name.endswith(".csv"):
        # Unknown/again ambiguous name (e.g. a NamedTemporaryFile) — sniff bytes.
        try:
            if hasattr(path_or_buffer, "read"):
                pos = path_or_buffer.tell()
                head = path_or_buffer.read(2)
                path_or_buffer.seek(pos)
            else:
                with open(path_or_buffer, "rb") as fh:
                    head = fh.read(2)
            is_excel = head[:2] == b"PK"
        except Exception:
            is_excel = False
    if is_excel:
        return pd.read_excel(path_or_buffer, skiprows=skiprows, dtype=dtype)
    return pd.read_csv(path_or_buffer, skiprows=skiprows, dtype=dtype)


class CUR:
    def __init__(self, df: pd.DataFrame):
        self.df = df

    @classmethod
    def load(cls, path_or_buffer) -> "CUR":
        df = read_tabular(path_or_buffer, dtype=str)
        for col in ("CostBeforeTax", "UsageQuantity"):
            df[col] = pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0.0)
        linked = df.get("LinkedAccountId", pd.Series([""] * len(df))).fillna("")
        payer  = df.get("PayerAccountId",  pd.Series([""] * len(df))).fillna("")
        df["_acct"] = linked.where(linked.str.strip() != "", payer)
        return cls(df)

    def line_items(self, allowed_ids: set) -> pd.DataFrame:
        """Return LinkedLineItem rows for the given account IDs."""
        rt = self.df.get("RecordType", pd.Series([""] * len(self.df)))
        li = self.df[rt == "LinkedLineItem"]
        return li[li["_acct"].isin(allowed_ids)]

    def active_accounts(self, allowed: dict) -> list[tuple[str, str]]:
        """Return [(aid, name)] for accounts that have any usage."""
        li = self.line_items(set(allowed.keys()))
        out = []
        for aid, name in allowed.items():
            if float(li[li["_acct"] == aid]["CostBeforeTax"].sum()) > 1e-6:
                out.append((aid, name))
        return out

    @staticmethod
    def svc(frame: pd.DataFrame, product: str) -> pd.DataFrame:
        return frame[frame["ProductName"] == product]

    @staticmethod
    def qty(frame: pd.DataFrame, ut_contains: str) -> float:
        m = frame[frame["UsageType"].str.contains(ut_contains, case=False, na=False)]
        return float(m["UsageQuantity"].sum())

    @staticmethod
    def cbt(frame: pd.DataFrame, ut_contains: str | None = None) -> float:
        if ut_contains:
            frame = frame[frame["UsageType"].str.contains(ut_contains, case=False, na=False)]
        return float(frame["CostBeforeTax"].sum())

    @staticmethod
    def desc(frame: pd.DataFrame, ut_contains: str) -> str:
        """First non-empty ItemDescription matching ut_contains."""
        m = frame[frame["UsageType"].str.contains(ut_contains, case=False, na=False)]
        for _, r in m.iterrows():
            d = str(r.get("ItemDescription", "")).strip()
            if d:
                return d
        return ""

    @staticmethod
    def skus(frame: pd.DataFrame) -> list[str]:
        """ItemDescription lines for all non-zero cost rows."""
        out = []
        for _, r in frame[frame["CostBeforeTax"] > 0].iterrows():
            ut  = str(r.get("UsageType", "")).strip()
            dsc = str(r.get("ItemDescription", "")).strip()[:70]
            qty = r.get("UsageQuantity", 0)
            if ut or dsc:
                out.append(f"{ut} - {dsc} ({float(qty):.4g})")
        return out
