"""Streamlit dashboard that visualizes Tally MIS data in a client-friendly way."""
from __future__ import annotations

from datetime import date, timedelta
import io

import pandas as pd
import streamlit as st

from tally_client import (
    _fiscal_year_start,
    fetch_companies,
    fetch_daybook,
    fetch_group_master,
    fetch_ledger_master,
)


st.set_page_config(page_title="Tally MIS Dashboard", layout="wide")


def _inject_theme():
    """Inject a simple navy/grey/red theme and typography tweaks."""

    st.markdown(
        """
        <style>
            :root {
                --navy: #0b1f3a;
                --navy-light: #15294d;
                --grey: #f3f4f6;
                --grey-mid: #d6d9de;
                --red: #e74c3c;
                --card-radius: 14px;
                --shadow: 0 10px 25px rgba(0,0,0,0.08);
                --text: #1c1f26;
                --muted: #6b7280;
                --font: "Inter", "Segoe UI", system-ui, -apple-system, sans-serif;
            }

            html, body, [class^="st-"], [class^="css"]  {
                font-family: var(--font);
            }

            .stApp {
                background: linear-gradient(180deg, #f7f8fb 0%, #eef1f6 35%, #e6e9ef 100%);
                color: var(--text);
            }

            .app-shell {
                background: white;
                padding: 18px 22px;
                border-radius: var(--card-radius);
                box-shadow: var(--shadow);
                border: 1px solid var(--grey-mid);
                margin-bottom: 14px;
            }

            .app-header {
                display: flex;
                align-items: center;
                gap: 14px;
                padding: 6px 0 14px 0;
            }

            .logo-chip {
                font-weight: 800;
                letter-spacing: 0.5px;
                color: white;
                background: linear-gradient(135deg, var(--navy), var(--navy-light));
                padding: 10px 14px;
                border-radius: 12px;
                box-shadow: var(--shadow);
                font-size: 18px;
            }

            .title-block small {
                color: var(--muted);
                font-size: 13px;
            }

            .metric-card {
                background: white;
                border-radius: var(--card-radius);
                border: 1px solid var(--grey-mid);
                box-shadow: var(--shadow);
                padding: 14px 16px;
                margin-bottom: 10px;
            }

            .metric-label {
                color: var(--muted);
                font-size: 13px;
                margin-bottom: 4px;
            }

            .metric-value {
                color: var(--navy);
                font-size: 22px;
                font-weight: 700;
            }

            .metric-accent {
                color: var(--red);
            }

            .download-card {
                background: white;
                border-radius: var(--card-radius);
                border: 1px solid var(--grey-mid);
                box-shadow: var(--shadow);
                padding: 16px;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


_inject_theme()

st.markdown(
    """
    <div class="app-header">
        <div class="logo-chip">Indefine.</div>
        <div class="title-block">
            <div style="font-size:22px; font-weight:700; color:var(--navy);">Tally Performance Dashboard</div>
            <small>Client-friendly KPIs, dynamic trial balance, and exports</small>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def _load_companies(host: str, port: int):
    return fetch_companies(host, port)


@st.cache_data(show_spinner=False)
def _load_daybook(company: str, host: str, port: int, start: date | None, end: date | None):
    """Load Day Book vouchers for the selected window (or full history if None)."""

    return fetch_daybook(company, start, end, host, port)


@st.cache_data(show_spinner=False)
def _load_ledger_master(
    company: str, host: str, port: int, from_date: date | None, to_date: date | None
):
    """Load ledger masters; cache keyed on dates so user inputs refresh dependent data."""

    return fetch_ledger_master(company, host, port)


@st.cache_data(show_spinner=False)
def _load_group_master(
    company: str, host: str, port: int, from_date: date | None, to_date: date | None
):
    """Load group masters; cache keyed on dates so user inputs refresh dependent data."""

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


def _render_kpi(label: str, value: float, accent: bool = False):
    """Render a lightly styled KPI card with two-decimal formatting."""

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value{' metric-accent' if accent else ''}">₹{value:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _sum_t2clb(
    tb_df: pd.DataFrame,
    affects_gp: str,
    ledger_type: str,
    exclude_groups: set[str] | None = None,
) -> float:
    """Return the sum of T2Dynamic CLB for rows matching the given filters."""

    if tb_df is None or tb_df.empty:
        return 0.0

    affects_norm = affects_gp.lower()
    type_norm = ledger_type.lower()
    exclude_norm = {g.casefold() for g in exclude_groups} if exclude_groups else set()

    filtered = tb_df[
        (tb_df["AffectsGrossProfit"].astype(str).str.lower() == affects_norm)
        & (tb_df["Type"].astype(str).str.lower() == type_norm)
    ]

    if exclude_norm:
        filtered = filtered[
            ~filtered["GroupName"].astype(str).str.casefold().isin(exclude_norm)
        ]

    if "T2Dynamic CLB" not in filtered:
        return 0.0

    return float(filtered["T2Dynamic CLB"].astype(float).sum())


def _compute_cogs(tb_df: pd.DataFrame, opening_stock: float, closing_stock: float) -> float:
    """Compute COGS using user-entered stock levels and Purchase Accounts T2Dynamic CLB."""

    if tb_df is None or tb_df.empty:
        return 0.0

    purchase_mask = tb_df["GroupName"].astype(str).str.casefold() == "purchase accounts".casefold()
    purchases = 0.0
    if "T2Dynamic CLB" in tb_df:
        purchases = float(tb_df.loc[purchase_mask, "T2Dynamic CLB"].astype(float).sum())

    return float(opening_stock) + purchases - float(closing_stock)


def _render_overview_cards(tb_df: pd.DataFrame, opening_stock: float, closing_stock: float):
    """Render revenue/expense/profit overview cards derived from the dynamic trial balance."""

    revenue = -_sum_t2clb(tb_df, "yes", "income")
    direct_expense = _sum_t2clb(
        tb_df, "yes", "expense", exclude_groups={"Purchase Accounts"}
    )
    cogs = _compute_cogs(tb_df, opening_stock, closing_stock)
    gross_profit = revenue - direct_expense - cogs

    indirect_expense = _sum_t2clb(tb_df, "no", "expense")
    indirect_income = -_sum_t2clb(tb_df, "no", "income")
    net_profit = gross_profit + indirect_income - indirect_expense

    cards = st.columns(3, gap="large")
    with cards[0]:
        _render_kpi("Revenue (Direct)", revenue)
    with cards[1]:
        _render_kpi("Expense (Direct)", direct_expense)
    with cards[2]:
        _render_kpi("COGS", cogs)

    cards2 = st.columns(3, gap="large")
    with cards2[0]:
        _render_kpi("Gross Profit", gross_profit)
    with cards2[1]:
        _render_kpi("Income (Indirect)", indirect_income)
    with cards2[2]:
        _render_kpi("Expense (Indirect)", indirect_expense)

    cards3 = st.columns(1)
    with cards3[0]:
        _render_kpi("Net Profit", net_profit)


def _render_monthly_revenue_chart(voucher_df: pd.DataFrame):
    """Render month-on-month revenue from Day Book nett values (multiplied by -1)."""

    if voucher_df is None or voucher_df.empty:
        st.info("Load vouchers to view month-on-month revenue.")
        return

    df = voucher_df.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    if df.empty:
        st.info("No dated vouchers to chart.")
        return

    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    df["Revenue"] = df["Nett"].astype(float) * -1
    monthly = df.groupby("Month")["Revenue"].sum().reset_index()
    monthly = monthly.sort_values("Month")

    st.markdown("<div style='margin-top:12px;'>", unsafe_allow_html=True)
    st.subheader("Month-on-Month Revenue")
    st.bar_chart(monthly.set_index("Month"))
    st.markdown("</div>", unsafe_allow_html=True)


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

    overview_vouchers: pd.DataFrame | None = st.session_state.get("overview_vouchers_df")
    tb_df = st.session_state.get("tb_df")
    tb_from = st.session_state.get("tb_from")
    tb_to = st.session_state.get("tb_to")
    user_ob_input: float | None = st.session_state.get("user_ob_input")
    user_cb_input: float | None = st.session_state.get("user_cb_input")

    if overview_vouchers is None:
        overview_vouchers = _voucher_dataframe(_sample_vouchers())
        st.session_state.overview_vouchers_df = overview_vouchers

    overview_tab, table_tab = st.tabs(["Overview", "Table"])

    with table_tab:
        st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
        st.subheader("Dynamic Trial Balance Inputs")
        fetch_tb = False

        if company:
            default_from = tb_from or (date.today() - timedelta(days=30))
            default_to = tb_to or date.today()
            tb_col1, tb_col2, tb_col3, tb_col4, tb_col5 = st.columns([1, 1, 1, 1, 1])
            with tb_col1:
                tb_from = st.date_input("From date", value=default_from)
            with tb_col2:
                tb_to = st.date_input("To date", value=default_to)
            with tb_col3:
                user_ob_input = st.number_input(
                    "Opening stock for the period", value=float(user_ob_input or 0.0), step=1000.0, format="%.2f"
                )
            with tb_col4:
                user_cb_input = st.number_input(
                    "Closing stock for the period", value=float(user_cb_input or 0.0), step=1000.0, format="%.2f"
                )
            with tb_col5:
                st.write("\n")
                fetch_tb = st.button("Fetch Dynamic Trial Balance", type="primary")
        else:
            st.info("Select a company to configure the dynamic trial balance inputs.")

        if fetch_tb and company:
            if tb_from > tb_to:
                st.error("From date cannot be after To date.")
            else:
                with st.spinner("Computing dynamic trial balance..."):
                    try:
                        tb_df, tb_vouchers = _build_dynamic_trial_balance(
                            company, host, int(port), tb_from, tb_to
                        )
                    except Exception as exc:
                        st.error(f"Failed to build trial balance: {exc}")
                    else:
                        st.session_state.tb_df = tb_df
                        st.session_state.tb_vouchers_df = tb_vouchers
                        st.session_state.tb_from = tb_from
                        st.session_state.tb_to = tb_to
                        st.session_state.user_ob_input = float(user_ob_input or 0.0)
                        st.session_state.user_cb_input = float(user_cb_input or 0.0)
                        st.success(f"Dynamic trial balance ready ({len(tb_df):,} ledgers)")
                        st.caption(
                            f"User-supplied opening stock: {user_ob_input:,.2f} · "
                            f"User-supplied closing stock: {user_cb_input:,.2f}"
                        )
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
        st.subheader("Dynamic Trial Balance (Table)")
        current_tb = st.session_state.get("tb_df")
        if current_tb is not None and not current_tb.empty:
            from_label = tb_from if tb_from is not None else "NA"
            to_label = tb_to if tb_to is not None else "NA"
            st.download_button(
                label="Download Dynamic Trial Balance (Excel)",
                data=_to_excel_bytes(current_tb),
                file_name=f"Dynamic_Trial_Balance_{company}_{from_label}_{to_label}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.dataframe(
                current_tb.style.format(precision=2),
                use_container_width=True,
                height=520,
            )
        else:
            st.info("Fetch the dynamic trial balance to view the table.")
        st.markdown("</div>", unsafe_allow_html=True)

    # Refresh local variables from session state after processing inputs
    tb_df = st.session_state.get("tb_df")
    tb_from = st.session_state.get("tb_from") or tb_from
    tb_to = st.session_state.get("tb_to") or tb_to
    user_ob_input = st.session_state.get("user_ob_input") or user_ob_input or 0.0
    user_cb_input = st.session_state.get("user_cb_input") or user_cb_input or 0.0
    overview_vouchers = st.session_state.get("overview_vouchers_df") or overview_vouchers

    with overview_tab:
        st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
        st.subheader("Day Book")
        if company:
            if st.button("Load Full Day Book", type="primary"):
                with st.spinner("Fetching full Day Book..."):
                    try:
                        vouchers = _load_daybook(company, host, int(port), None, None)
                        overview_vouchers = _voucher_dataframe(vouchers)
                    except Exception as exc:
                        st.error(f"Failed to load Day Book: {exc}")
                    else:
                        st.session_state.overview_vouchers_df = overview_vouchers
                        st.success(f"Loaded {len(overview_vouchers):,} voucher lines")
        else:
            st.info("Select a company to load the Day Book.")

        if overview_vouchers is not None and not overview_vouchers.empty:
            st.caption(
                f"Voucher nett total: {overview_vouchers['Nett'].sum():,.2f} (should be 0.00 if balanced)"
            )
            st.download_button(
                label="Download Voucher Details (Excel)",
                data=_to_excel_bytes(overview_vouchers),
                file_name=f"Day_Book_{company or 'Sample'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.dataframe(
                overview_vouchers.style.format({"Debit": "{:.2f}", "Credit": "{:.2f}", "Nett": "{:.2f}"}),
                use_container_width=True,
                height=420,
            )
        st.markdown("</div>", unsafe_allow_html=True)

        if tb_df is not None and not tb_df.empty:
            st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
            st.subheader("Performance Overview (Dynamic)")
            opening_stock_val = float(user_ob_input or 0.0)
            closing_stock_val = float(user_cb_input or 0.0)

            _render_overview_cards(tb_df, opening_stock_val, closing_stock_val)
            _render_monthly_revenue_chart(overview_vouchers)
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.info("Select a company to compute the dynamic trial balance.")

        st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
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
        st.markdown("</div>", unsafe_allow_html=True)


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


def _build_dynamic_trial_balance(
    company: str, host: str, port: int, from_date: date, to_date: date
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Assemble a dynamic trial balance and return the supporting Day Book data."""

    fiscal_start = _fiscal_year_start(from_date)

    vouchers = _load_daybook(company, host, port, fiscal_start, to_date)
    ledger_rows = _load_ledger_master(company, host, port, from_date, to_date)
    group_rows = _load_group_master(company, host, port, from_date, to_date)

    ledger_parent_map = {row["LedgerName"]: row["LedgerParent"] for row in ledger_rows}
    ledger_opening_map = {row["LedgerName"]: row["OpeningBalanceNormalized"] for row in ledger_rows}

    group_map = {row["GroupName"]: row for row in group_rows}

    # Build Day Book nets by ledger for the two required windows.
    voucher_df = _voucher_dataframe(vouchers)
    if voucher_df.empty:
        raise RuntimeError("Day Book is empty; cannot compute trial balance.")

    voucher_df["Date"] = pd.to_datetime(voucher_df["Date"]).dt.date
    voucher_df["Nett"] = voucher_df["Nett"].astype(float)

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
        t2_dynamic_clb = nets_in_range.get(ledger_name, 0.0)

        dynamic_opening = opening + t2_dynamic_ob
        dynamic_closing = dynamic_opening + t2_dynamic_clb

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
                "T2Dynamic CLB": t2_dynamic_clb,
                "DynamicClosing": dynamic_closing,
            }
        )

    tb_df = pd.DataFrame(rows).sort_values("LedgerName").reset_index(drop=True)
    return tb_df, voucher_df


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


if __name__ == "__main__":
    main()

