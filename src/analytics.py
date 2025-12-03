"""Analytics helpers for summarizing Tally vouchers into MIS-friendly numbers."""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Tuple

from tally_client import LedgerEntry, Voucher


@dataclass
class FinancialSnapshot:
    revenue: float
    expenses: float
    gross_margin: float
    profit_loss: float
    assets: float
    liabilities: float
    best_sellers: List[Tuple[str, float]]


def summarize(vouchers: Iterable[Voucher], ledger_groups: Dict[str, str] | None = None) -> FinancialSnapshot:
    """Aggregate the Day Book vouchers into key MIS metrics.

    ledger_groups maps ledger names to categories like "Revenue", "Expense", "Asset", etc.
    This function falls back to basic heuristics when the map is missing entries.
    """

    ledger_groups = ledger_groups or {}
    totals: Dict[str, float] = defaultdict(float)
    product_sales: Dict[str, float] = defaultdict(float)

    for voucher in vouchers:
        for entry in voucher.ledger_entries:
            category = ledger_groups.get(entry.ledger_name, _infer_category(entry.ledger_name))
            amount = entry.amount if entry.is_debit else -entry.amount
            totals[category] += amount

            if category == "Revenue":
                product_sales[entry.ledger_name] += amount

    revenue = totals["Revenue"]
    expenses = totals["Expense"]
    assets = totals["Asset"]
    liabilities = totals["Liability"]
    profit_loss = revenue - expenses
    gross_margin = revenue - totals["Cost of Goods Sold"]
    best_sellers = sorted(product_sales.items(), key=lambda kv: kv[1], reverse=True)[:5]

    return FinancialSnapshot(
        revenue=revenue,
        expenses=expenses,
        gross_margin=gross_margin,
        profit_loss=profit_loss,
        assets=assets,
        liabilities=liabilities,
        best_sellers=best_sellers,
    )


def _infer_category(ledger_name: str) -> str:
    lower_name = ledger_name.lower()
    if "sale" in lower_name or "revenue" in lower_name:
        return "Revenue"
    if "cogs" in lower_name or "cost of goods" in lower_name or "inventory" in lower_name:
        return "Cost of Goods Sold"
    if "expense" in lower_name or "rent" in lower_name or "salary" in lower_name or "marketing" in lower_name:
        return "Expense"
    if "asset" in lower_name or "bank" in lower_name or "cash" in lower_name:
        return "Asset"
    if "loan" in lower_name or "payable" in lower_name or "liability" in lower_name:
        return "Liability"
    return "Expense"


