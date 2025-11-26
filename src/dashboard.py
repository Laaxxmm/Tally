"""Streamlit dashboard that visualizes Tally MIS data in a client-friendly way."""
from __future__ import annotations

from datetime import date, timedelta
from typing import Dict
import io

import pandas as pd
import streamlit as st

from analytics import summarize
from tally_client import (
    _fiscal_year_start,
    fetch_companies,
    fetch_daybook,
    fetch_group_master,
    fetch_ledger_master,
)


st.set_page_config(page_title="Tally MIS Dashboard", layout="wide")
st.title("📊 Tally MIS Dashboard")
st.caption("Connects to Tally over 127.0.0.1:9000 to present client-ready KPIs.")


@st.cache_data(show_spinner=False)
def _load_companies(host: str, port: int):
    return fetch_companies(host, port)


@st.cache_data(show_spinner=False)
def _load_data(company: str, host: str, port: int):
    """Load the full Day Book history for the selected company."""

    return fetch_daybook(company, None, None, host, port)


@st.cache_data(show_spinner=False)
def _load_ledger_master(company: str, host: str, port: int):
    return fetch_ledger_master(company, host, port)


@st.cache_data(show_spinner=False)
def _load_group_master(company: str, host: str, port: int):
    return fetch_group_master(company, host, port)


@st.cache_data(show_spinner=False)
def _build_ledger_excel(company: str, host: str, port: int):
    """Fetch ledger masters and return Excel bytes plus ledger count."""

    rows = fetch_ledger_master(company, host, port)
    df = pd.DataFrame(rows, columns=[
        "LedgerName",
        "LedgerParent",
        "OpeningBalanceRaw",
        "OpeningBalanceNormalized",
    ])
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    return buffer.read(), len(df)


@st.cache_data(show_spinner=False)
def _build_group_excel(company: str, host: str, port: int):
    """Fetch group masters and return Excel bytes plus group count."""

    rows = fetch_group_master(company, host, port)
    df = pd.DataFrame(rows, columns=[
        "GroupName",
        "ParentName",
        "BS_or_PnL",
        "Type",
        "AffectsGrossProfit",
    ])
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    return buffer.read(), len(df)


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

    companies: list[str] = st.session_state.get("companies", [])

    if st.sidebar.button("Connect to Tally", type="primary"):
        with st.spinner("Discovering companies..."):
            companies = _load_companies(host, int(port))
            if companies:
                st.session_state.companies = companies
                st.sidebar.success(f"Connected · {len(companies)} companies found")
            else:
                st.sidebar.error("No companies returned. Make sure Tally is open and ODBC/HTTP is enabled.")

    company = None
    if companies:
        company = st.sidebar.selectbox("Company", companies)

    vouchers = []
    if company and st.sidebar.button("Load full Day Book", type="primary"):
        with st.spinner(f"Loading full Day Book for {company}..."):
            try:
                vouchers = _load_data(company, host, int(port))
            except Exception as exc:  # requests or parsing failures
                st.error(f"Tally connection failed: {exc}")

    if not vouchers:
        st.info("Showing sample data. Use the sidebar to connect and load live vouchers from Tally.")
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
    voucher_df = _voucher_dataframe(vouchers)
    st.dataframe(voucher_df, use_container_width=True)
    st.caption(
        f"Voucher nett total: {voucher_df['Nett'].sum():,.2f} (should be 0.00 if balanced)"
    )

    st.markdown("---")
    st.subheader("Dynamic Trial Balance")
    if company:
        tb_col1, tb_col2, tb_col3 = st.columns([1, 1, 1])
        with tb_col1:
            tb_from = st.date_input("From date", value=date.today() - timedelta(days=30))
        with tb_col2:
            tb_to = st.date_input("To date", value=date.today())
        with tb_col3:
            st.write("\n")
            fetch_tb = st.button("Fetch Dynamic Trial Balance", type="primary")

        if fetch_tb:
            if tb_from > tb_to:
                st.error("From date cannot be after To date.")
            else:
                with st.spinner("Computing dynamic trial balance..."):
                    try:
                        tb_df = _build_dynamic_trial_balance(company, host, int(port), tb_from, tb_to)
                    except Exception as exc:
                        st.error(f"Failed to build trial balance: {exc}")
                    else:
                        st.success(f"Dynamic trial balance ready ({len(tb_df):,} ledgers)")
                        st.dataframe(
                            tb_df.style.format(
                                {
                                    "T2Dynamic OB": "₹{:,.2f}",
                                    "DynamicOpening": "₹{:,.2f}",
                                    "DynamicClosing": "₹{:,.2f}",
                                    "OpeningBalance": "₹{:,.2f}",
                                }
                            ),
                            use_container_width=True,
                        )
                        st.download_button(
                            label="Download Dynamic Trial Balance (Excel)",
                            data=_to_excel_bytes(tb_df),
                            file_name=f"Dynamic_Trial_Balance_{company}_{tb_from}_{tb_to}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
    else:
        st.info("Select a company to compute the dynamic trial balance.")

    st.markdown("---")
    st.subheader("Chart of Accounts (Download Only)")
    if company:
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Download Ledger Openings (Excel)", type="primary"):
                with st.spinner("Building ledger master workbook..."):
                    try:
                        excel_bytes, count = _build_ledger_excel(company, host, int(port))
                    except Exception as exc:
                        st.error(f"Failed to load ledgers: {exc}")
                    else:
                        st.success(f"Ready · {count:,} ledgers")
                        st.download_button(
                            label="Download Ledger Master Opening Balances",
                            data=excel_bytes,
                            file_name=f"Ledger_Master_Openings_{company}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
        with col_b:
            if st.button("Download Group Master (Excel)", type="primary"):
                with st.spinner("Building group master workbook..."):
                    try:
                        excel_bytes, count = _build_group_excel(company, host, int(port))
                    except Exception as exc:
                        st.error(f"Failed to load groups: {exc}")
                    else:
                        st.success(f"Ready · {count:,} groups")
                        st.download_button(
                            label="Download Group Master", 
                            data=excel_bytes,
                            file_name=f"Group_Master_{company}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
    else:
        st.info("Select a company to download its ledger and group lists.")


def _voucher_dataframe(vouchers):
    rows = []
    for voucher in vouchers:
        for entry in voucher.ledger_entries:
            debit_amount = entry.amount if entry.is_debit else 0
            credit_amount = entry.amount if not entry.is_debit else 0
            # Column labels need to be flipped: what was shown as "Debit" should
            # now appear under "Credit" and vice versa. The nett value should
            # reflect the corrected columns (Debit minus Credit).
            corrected_debit = credit_amount
            corrected_credit = debit_amount
            net = corrected_debit - corrected_credit
            rows.append(
                {
                    "Date": voucher.date,
                    "Voucher No": voucher.voucher_number,
                    "Voucher Type": voucher.voucher_type,
                    "Ledger": entry.ledger_name,
                    "Debit": corrected_debit,
                    "Credit": corrected_credit,
                    "Nett": net,
                }
            )
    return pd.DataFrame(rows)


def _build_dynamic_trial_balance(company: str, host: str, port: int, from_date: date, to_date: date) -> pd.DataFrame:
    """Assemble a dynamic trial balance using Day Book, ledger master, and group master data."""

    vouchers = _load_data(company, host, port)
    ledger_rows = _load_ledger_master(company, host, port)
    group_rows = _load_group_master(company, host, port)

    ledger_parent_map = {row["LedgerName"]: row["LedgerParent"] for row in ledger_rows}
    ledger_opening_map = {row["LedgerName"]: row["OpeningBalanceNormalized"] for row in ledger_rows}

    group_map = {row["GroupName"]: row for row in group_rows}

    # Build Day Book nets by ledger for the two required windows.
    voucher_df = _voucher_dataframe(vouchers)
    if voucher_df.empty:
        raise RuntimeError("Day Book is empty; cannot compute trial balance.")

    voucher_df["Date"] = pd.to_datetime(voucher_df["Date"]).dt.date
    voucher_df["Nett"] = voucher_df["Nett"].astype(float)

    # Restrict Day Book movements to the fiscal year that aligns with the ledger openings.
    fiscal_start = _fiscal_year_start(from_date)
    voucher_df = voucher_df[voucher_df["Date"] >= fiscal_start]

    t2_mask = (voucher_df["Date"] >= fiscal_start) & (voucher_df["Date"] < from_date)
    in_range_mask = (voucher_df["Date"] >= from_date) & (voucher_df["Date"] <= to_date)

    nets_t2 = voucher_df.loc[t2_mask].groupby("Ledger")["Nett"].sum()
    nets_in_range = voucher_df.loc[in_range_mask].groupby("Ledger")["Nett"].sum()

    rows = []
    for ledger_name, opening in ledger_opening_map.items():
        parent = ledger_parent_map.get(ledger_name, "") or "(Unknown)"
        group_info = group_map.get(parent, {})
        bs_pnl = group_info.get("BS_or_PnL", "")
        gtype = group_info.get("Type", "")
        affects_gp = group_info.get("AffectsGrossProfit", "")

        t2_dynamic_ob = nets_t2.get(ledger_name, 0.0)
        in_range_net = nets_in_range.get(ledger_name, 0.0)

        dynamic_opening = opening + t2_dynamic_ob
        dynamic_closing = dynamic_opening + in_range_net

        rows.append(
            {
                "LedgerName": ledger_name,
                "GroupName": parent,
                "ParentName": group_info.get("ParentName", parent),
                "BS_or_PnL": bs_pnl,
                "Type": gtype,
                "AffectsGrossProfit": affects_gp,
                "OpeningBalance": opening,
                "T2Dynamic OB": t2_dynamic_ob,
                "DynamicOpening": dynamic_opening,
                "DynamicClosing": dynamic_closing,
            }
        )

    return pd.DataFrame(rows).sort_values("LedgerName").reset_index(drop=True)


def _to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    buffer.seek(0)
    return buffer.read()


def _sample_vouchers():
    from tally_client import LedgerEntry, Voucher
    sample = [
        Voucher(
            voucher_type="Sales",
            date=date.today() - timedelta(days=3),
            ledger_entries=[
                LedgerEntry("Product A", 120000.0, True),
                LedgerEntry("Bank", 70000.0, False),
                LedgerEntry("CGS", 50000.0, False),
            ],
        ),
        Voucher(
            voucher_type="Sales",
            date=date.today() - timedelta(days=2),
            ledger_entries=[
                LedgerEntry("Product B", 80000.0, True),
                LedgerEntry("Bank", 50000.0, False),
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

