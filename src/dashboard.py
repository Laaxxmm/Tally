"""Streamlit dashboard that visualizes Tally MIS data in a client-friendly way."""
from __future__ import annotations

from datetime import date, timedelta
from typing import Dict

import pandas as pd
import streamlit as st

from analytics import summarize
from tally_client import TallyClient


st.set_page_config(page_title="Tally MIS Dashboard", layout="wide")
st.title("📊 Tally MIS Dashboard")
st.caption("Connects to Tally over 127.0.0.1:9000 to present client-ready KPIs.")


@st.cache_data(show_spinner=False)
def _load_data(start: date, end: date, host: str, port: int):
    client = TallyClient(host=host, port=port)
    return client.fetch_daybook(start, end)


def _render_kpi(label: str, value: float, delta: float | None = None):
    formatted_value = f"₹{value:,.2f}"
    if delta is not None:
        st.metric(label, formatted_value, f"₹{delta:,.2f}")
    else:
        st.metric(label, formatted_value)


def main() -> None:
    st.sidebar.header("Tally Connection")
    host = st.sidebar.text_input("Host", value="127.0.0.1")
    port = st.sidebar.number_input("Port", value=9000, step=1)

    st.sidebar.header("Date Range")
    today = date.today()
    default_start = today - timedelta(days=30)
    start_date = st.sidebar.date_input("From", value=default_start)
    end_date = st.sidebar.date_input("To", value=today)

    if start_date > end_date:
        st.sidebar.error("Start date must be before end date")
        st.stop()

    if st.sidebar.button("Load from Tally"):
        with st.spinner("Loading Day Book from Tally..."):
            vouchers = _load_data(start_date, end_date, host, int(port))
    else:
        st.info("Click 'Load from Tally' to fetch live data. Showing sample data below.")
        vouchers = []

    if not vouchers:
        vouchers = _sample_vouchers()

    ledger_map = _default_ledger_map()
    snapshot = summarize(vouchers, ledger_map)

    kpi_cols = st.columns(3)
    with kpi_cols[0]:
        _render_kpi("Revenue", snapshot.revenue)
    with kpi_cols[1]:
        _render_kpi("Profit / Loss", snapshot.profit_loss)
    with kpi_cols[2]:
        _render_kpi("Gross Margin", snapshot.gross_margin)

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Balance Sheet Snapshot")
        st.bar_chart(
            pd.DataFrame(
                {
                    "Category": ["Assets", "Liabilities"],
                    "Amount": [snapshot.assets, snapshot.liabilities],
                }
            ).set_index("Category")
        )

    with col2:
        st.subheader("Top Products / Services")
        st.table(pd.DataFrame(snapshot.best_sellers, columns=["Ledger", "Revenue"]))

    st.markdown("---")
    st.subheader("Voucher Details")
    st.dataframe(_voucher_dataframe(vouchers), use_container_width=True)



def _voucher_dataframe(vouchers):
    rows = []
    for voucher in vouchers:
        for entry in voucher.ledger_entries:
            rows.append(
                {
                    "Date": voucher.date,
                    "Voucher Type": voucher.voucher_type,
                    "Ledger": entry.ledger_name,
                    "Debit": entry.amount if entry.is_debit else 0,
                    "Credit": entry.amount if not entry.is_debit else 0,
                }
            )
    return pd.DataFrame(rows)


def _sample_vouchers():
    from tally_client import LedgerEntry, Voucher
    sample = [
        Voucher(
            voucher_type="Sales",
            date=date.today() - timedelta(days=3),
            ledger_entries=[
                LedgerEntry("Product A", 120000.0, True),
                LedgerEntry("CGS", 50000.0, False),
            ],
        ),
        Voucher(
            voucher_type="Sales",
            date=date.today() - timedelta(days=2),
            ledger_entries=[
                LedgerEntry("Product B", 80000.0, True),
                LedgerEntry("CGS", 30000.0, False),
            ],
        ),
        Voucher(
            voucher_type="Payment",
            date=date.today() - timedelta(days=1),
            ledger_entries=[
                LedgerEntry("Rent Expense", 15000.0, True),
                LedgerEntry("Bank", 15000.0, False),
            ],
        ),
    ]
    return sample


def _default_ledger_map() -> Dict[str, str]:
    return {
        "Product A": "Revenue",
        "Product B": "Revenue",
        "CGS": "Cost of Goods Sold",
        "Rent Expense": "Expense",
        "Bank": "Asset",
    }


if __name__ == "__main__":
    main()

