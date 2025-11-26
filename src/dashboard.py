"""Streamlit dashboard that visualizes Tally MIS data in a client-friendly way."""
from __future__ import annotations

from datetime import date, timedelta
import calendar
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

            /* Global button styling */
            .stButton > button, .stDownloadButton > button {
                background: var(--navy);
                color: white !important;
                border-radius: var(--card-radius);
                border: 1px solid var(--navy);
                box-shadow: var(--shadow);
                padding: 10px 16px;
                font-weight: 600;
            }

            .stButton > button:hover, .stDownloadButton > button:hover {
                background: var(--navy-light);
                border-color: var(--navy-light);
            }

            /* KPI buttons inside the overview */
            .kpi-button button {
                background: var(--navy);
                border-radius: var(--card-radius);
                border: 1px solid var(--navy);
                box-shadow: var(--shadow);
                padding: 14px 16px;
                color: white;
                font-weight: 700;
                font-size: 16px;
                text-align: left;
                white-space: pre-line;
            }

            .kpi-button button:hover {
                background: var(--navy-light);
                border-color: var(--navy-light);
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


def _kpi_button(label: str, value: float, key: str):
    """Clickable KPI-style button that preserves the card look."""

    with st.container():
        st.markdown("<div class='kpi-button'>", unsafe_allow_html=True)
        clicked = st.button(
            f"{label}\n₹{value:,.2f}",
            key=key,
            use_container_width=True,
            type="secondary",
            help="Click to view month-on-month trend",
        )
        st.markdown("</div>", unsafe_allow_html=True)
    return clicked


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


def _sum_ytd_nett(
    tb_df: pd.DataFrame,
    affects_gp: str,
    ledger_type: str,
    exclude_groups: set[str] | None = None,
) -> float:
    """Sum Nett YTD for rows matching filters (for YTD overview KPIs)."""

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

    if "Nett YTD" not in filtered:
        return 0.0

    return float(filtered["Nett YTD"].astype(float).sum())


def _merge_vouchers_with_meta(voucher_df: pd.DataFrame, tb_df: pd.DataFrame) -> pd.DataFrame:
    """Attach group/type metadata from the YTD trial balance to voucher rows."""

    if voucher_df is None or voucher_df.empty or tb_df is None or tb_df.empty:
        return pd.DataFrame()

    meta = tb_df[[
        "LedgerName",
        "GroupName",
        "AffectsGrossProfit",
        "Type",
    ]].rename(columns={
        "LedgerName": "Ledger",
    })

    merged = voucher_df.copy()
    merged["Date"] = pd.to_datetime(merged["Date"])
    return merged.merge(meta, on="Ledger", how="left")


def _monthly_series(
    merged_df: pd.DataFrame,
    affects_gp: str | None = None,
    ledger_type: str | None = None,
    exclude_groups: set[str] | None = None,
    group_equals: str | None = None,
    sign: float = 1.0,
) -> pd.Series:
    """Aggregate nett amounts by month using the merged voucher+metadata frame."""

    if merged_df is None or merged_df.empty:
        return pd.Series(dtype=float)

    df = merged_df.copy()
    if affects_gp is not None:
        df = df[df["AffectsGrossProfit"].astype(str).str.lower() == affects_gp.lower()]
    if ledger_type is not None:
        df = df[df["Type"].astype(str).str.lower() == ledger_type.lower()]
    if exclude_groups:
        exclude_norm = {g.casefold() for g in exclude_groups}
        df = df[~df["GroupName"].astype(str).str.casefold().isin(exclude_norm)]
    if group_equals:
        df = df[df["GroupName"].astype(str).str.casefold() == group_equals.casefold()]

    if df.empty:
        return pd.Series(dtype=float)

    df["Month"] = df["Date"].dt.to_period("M").dt.to_timestamp()
    series = df.groupby("Month")["Nett"].sum().sort_index()
    return series * float(sign)


def _monthly_cogs_series(
    merged_df: pd.DataFrame, opening_stock: float, closing_stock: float
) -> pd.Series:
    """Compute month-wise COGS using Purchase Accounts plus stock adjustments."""

    purchases = _monthly_series(
        merged_df,
        affects_gp=None,
        ledger_type=None,
        exclude_groups=None,
        group_equals="Purchase Accounts",
        sign=1.0,
    )

    if purchases.empty:
        return purchases

    months_sorted = purchases.index.sort_values()
    first_month = months_sorted[0]
    last_month = months_sorted[-1]

    purchases.loc[first_month] += float(opening_stock or 0.0)
    purchases.loc[last_month] -= float(closing_stock or 0.0)
    return purchases


def _combine_monthly_series(series_list: list[pd.Series]) -> pd.Series:
    """Align and sum multiple monthly series."""

    if not series_list:
        return pd.Series(dtype=float)

    index_union = pd.Index([])
    for series in series_list:
        index_union = index_union.union(series.index)

    combined = pd.Series(0.0, index=index_union.sort_values())
    for series in series_list:
        combined.loc[series.index] += series
    return combined


def _build_kpi_monthly_series(
    selected_key: str,
    voucher_df: pd.DataFrame,
    tb_df: pd.DataFrame,
    opening_stock: float,
    closing_stock: float,
) -> pd.Series:
    """Return month-wise amounts for the selected KPI."""

    merged = _merge_vouchers_with_meta(voucher_df, tb_df)
    if merged.empty:
        return pd.Series(dtype=float)

    revenue_series = _monthly_series(merged, affects_gp="yes", ledger_type="income", sign=-1)
    direct_exp_series = _monthly_series(
        merged,
        affects_gp="yes",
        ledger_type="expense",
        exclude_groups={"Purchase Accounts"},
    )
    indirect_income_series = _monthly_series(merged, affects_gp="no", ledger_type="income", sign=-1)
    indirect_exp_series = _monthly_series(merged, affects_gp="no", ledger_type="expense")
    cogs_series = _monthly_cogs_series(merged, opening_stock, closing_stock)

    if selected_key == "revenue":
        return revenue_series
    if selected_key == "direct_expense":
        return direct_exp_series
    if selected_key == "cogs":
        return cogs_series
    if selected_key == "gross_profit":
        return _combine_monthly_series([revenue_series, -direct_exp_series, -cogs_series])
    if selected_key == "indirect_income":
        return indirect_income_series
    if selected_key == "indirect_expense":
        return indirect_exp_series
    if selected_key == "net_profit":
        gross = _combine_monthly_series([revenue_series, -direct_exp_series, -cogs_series])
        return _combine_monthly_series([gross, indirect_income_series, -indirect_exp_series])

    return pd.Series(dtype=float)


def _render_ytd_overview_cards(tb_df: pd.DataFrame, opening_stock: float, closing_stock: float, voucher_df: pd.DataFrame):
    """Render revenue/expense/profit overview cards derived from the YTD trial balance."""

    revenue = -_sum_ytd_nett(tb_df, "yes", "income")
    direct_expense = _sum_ytd_nett(tb_df, "yes", "expense", exclude_groups={"Purchase Accounts"})

    purchase_mask = tb_df["GroupName"].astype(str).str.casefold() == "purchase accounts".casefold()
    purchases_ytd = 0.0
    if "Nett YTD" in tb_df:
        purchases_ytd = float(tb_df.loc[purchase_mask, "Nett YTD"].astype(float).sum())

    cogs = float(opening_stock) + purchases_ytd - float(closing_stock)
    gross_profit = revenue - direct_expense - cogs

    indirect_expense = _sum_ytd_nett(tb_df, "no", "expense")
    indirect_income = -_sum_ytd_nett(tb_df, "no", "income")
    net_profit = gross_profit + indirect_income - indirect_expense

    selection_key = "ytd_selected_kpi"
    selected = st.session_state.get(selection_key)

    card_specs = [
        {
            "key": "revenue",
            "label": "Revenue (Direct) YTD",
            "value": revenue,
            "affects": "yes",
            "type": "income",
            "exclude": None,
            "sign": -1.0,
        },
        {
            "key": "direct_expense",
            "label": "Expense (Direct) YTD",
            "value": direct_expense,
            "affects": "yes",
            "type": "expense",
            "exclude": {"Purchase Accounts"},
            "sign": 1.0,
        },
        {
            "key": "cogs",
            "label": "COGS (YTD)",
            "value": cogs,
            "kind": "cogs",
        },
        {
            "key": "gross_profit",
            "label": "Gross Profit (YTD)",
            "value": gross_profit,
            "kind": "gross",
        },
        {
            "key": "indirect_income",
            "label": "Income (Indirect) YTD",
            "value": indirect_income,
            "affects": "no",
            "type": "income",
            "exclude": None,
            "sign": -1.0,
        },
        {
            "key": "indirect_expense",
            "label": "Expense (Indirect) YTD",
            "value": indirect_expense,
            "affects": "no",
            "type": "expense",
            "exclude": None,
            "sign": 1.0,
        },
        {
            "key": "net_profit",
            "label": "Net Profit (YTD)",
            "value": net_profit,
            "kind": "net",
        },
    ]

    rows = [card_specs[:3], card_specs[3:6], card_specs[6:]]

    for row in rows:
        cols = st.columns(len(row), gap="large")
        for col, spec in zip(cols, row):
            with col:
                clicked = _kpi_button(spec["label"], spec["value"], f"kpi-{spec['key']}")
                if clicked:
                    selected = spec["key"]
                    st.session_state[selection_key] = selected

    return selected


def _statement_from_tb(tb_df: pd.DataFrame, statement: str) -> pd.DataFrame:
    """Return a P&L or Balance Sheet slice from the dynamic trial balance."""

    if tb_df is None or tb_df.empty:
        return pd.DataFrame()

    statement_norm = statement.lower()

    if statement_norm == "p&l":
        if "T2Dynamic CLB" not in tb_df.columns:
            return pd.DataFrame()

        def pnl_amount(row: pd.Series) -> float:
            raw = float(row.get("T2Dynamic CLB", 0.0) or 0.0)
            t_val = str(row.get("Type", "")).lower()
            # Incomes are stored negative in the TB cards, flip to display positive figures.
            return -raw if t_val == "income" else raw

        bs_pnl = tb_df.get("BS_or_PnL", pd.Series(dtype=str)).astype(str).str.lower()
        pnl_mask = bs_pnl.str.contains("p&l") | bs_pnl.str.contains("profit")
        type_mask = tb_df.get("Type", pd.Series(dtype=str)).astype(str).str.lower().isin(
            ["income", "expense"]
        )
        mask = pnl_mask | type_mask
        pnl_df = tb_df.loc[mask].copy()
        pnl_df["DisplayAmount"] = pnl_df.apply(pnl_amount, axis=1)

        def summarize(section_mask: pd.Series) -> pd.DataFrame:
            section = pnl_df.loc[section_mask]
            if section.empty:
                return pd.DataFrame()
            grp = (
                section.groupby("GroupName")["DisplayAmount"]
                .sum()
                .reset_index()
                .sort_values("GroupName")
            )
            grp["Line"] = "· " + grp["GroupName"].astype(str)
            grp["Amount"] = grp["DisplayAmount"]
            return grp[["Line", "Amount"]]

        revenue_mask = pnl_df["Type"].astype(str).str.lower() == "income"
        cogs_mask = (pnl_df["Type"].astype(str).str.lower() == "expense") & (
            pnl_df["AffectsGrossProfit"].astype(str).str.lower() == "yes"
        )
        opex_mask = (pnl_df["Type"].astype(str).str.lower() == "expense") & (
            pnl_df["AffectsGrossProfit"].astype(str).str.lower() != "yes"
        )
        other_income_mask = (~cogs_mask) & (
            pnl_df["Type"].astype(str).str.lower() == "income"
        ) & (pnl_df["AffectsGrossProfit"].astype(str).str.lower() != "yes")

        revenue_total = pnl_df.loc[revenue_mask, "DisplayAmount"].sum()
        cogs_total = pnl_df.loc[cogs_mask, "DisplayAmount"].sum()
        opex_total = pnl_df.loc[opex_mask, "DisplayAmount"].sum()
        other_income_total = pnl_df.loc[other_income_mask, "DisplayAmount"].sum()

        lines: list[dict[str, float | str]] = []
        lines.append({"Line": "Revenue", "Amount": revenue_total})
        lines.extend(summarize(revenue_mask).to_dict(orient="records"))
        gross_profit = revenue_total - cogs_total
        lines.append({"Line": "COGS", "Amount": cogs_total})
        lines.extend(summarize(cogs_mask).to_dict(orient="records"))
        lines.append({"Line": "Gross Profit", "Amount": gross_profit})
        lines.append({"Line": "Operating Expenses", "Amount": opex_total})
        lines.extend(summarize(opex_mask).to_dict(orient="records"))
        ebit = gross_profit - opex_total
        lines.append({"Line": "EBIT", "Amount": ebit})
        if other_income_total != 0:
            lines.append({"Line": "Other Income", "Amount": other_income_total})
            lines.extend(summarize(other_income_mask).to_dict(orient="records"))
        net_profit = ebit + other_income_total
        lines.append({"Line": "Net Profit", "Amount": net_profit})

        return pd.DataFrame(lines)

    if statement_norm in {"balance sheet", "balance", "bs"}:
        if "DynamicClosing" not in tb_df.columns:
            return pd.DataFrame()

        def bs_amount(row: pd.Series) -> float:
            raw = float(row.get("DynamicClosing", 0.0) or 0.0)
            t_val = str(row.get("Type", "")).lower()
            # Show assets as positive figures and liabilities as positive by
            # flipping their credit-balance sign for readability.
            if t_val == "liability":
                return -raw
            return raw

        bs_pnl = tb_df.get("BS_or_PnL", pd.Series(dtype=str)).astype(str).str.lower()
        mask = bs_pnl.str.contains("balance") | bs_pnl.str.contains("bs")

        if not mask.any():
            mask = tb_df["Type"].astype(str).str.lower().isin(["asset", "liability"])

        bs_df = tb_df.loc[mask].copy()
        bs_df["DisplayAmount"] = bs_df.apply(bs_amount, axis=1)

        def summarize(section_mask: pd.Series) -> pd.DataFrame:
            section = bs_df.loc[section_mask]
            if section.empty:
                return pd.DataFrame()
            grp = (section.groupby("GroupName")["DisplayAmount"].sum().reset_index().sort_values("GroupName"))
            grp["Line"] = "· " + grp["GroupName"].astype(str)
            grp["Amount"] = grp["DisplayAmount"]
            return grp[["Line", "Amount"]]

        asset_mask = bs_df["Type"].astype(str).str.lower() == "asset"
        liability_mask = bs_df["Type"].astype(str).str.lower() == "liability"

        assets_total = bs_df.loc[asset_mask, "DisplayAmount"].sum()
        liabilities_total = bs_df.loc[liability_mask, "DisplayAmount"].sum()

        lines: list[dict[str, float | str]] = []
        lines.append({"Line": "Assets", "Amount": assets_total})
        lines.extend(summarize(asset_mask).to_dict(orient="records"))
        lines.append({"Line": "Liabilities & Equity", "Amount": liabilities_total})
        lines.extend(summarize(liability_mask).to_dict(orient="records"))
        lines.append({"Line": "Total (Assets - Liabilities)", "Amount": assets_total - liabilities_total})

        return pd.DataFrame(lines)

    bs_pnl = tb_df.get("BS_or_PnL", pd.Series(dtype=str)).astype(str).str.lower()

    if statement_norm == "p&l":
        mask = bs_pnl.str.contains("p&l") | bs_pnl.str.contains("profit")
    else:
        mask = bs_pnl.str.contains("balance") | bs_pnl.str.contains("bs")

    # Fallback: if BS/P&L flag is missing, fall back to Type classification.
    if not mask.any():
        if statement_norm == "p&l":
            mask = tb_df["Type"].astype(str).str.lower().isin(["income", "expense"])
        else:
            mask = tb_df["Type"].astype(str).str.lower().isin(["asset", "liability"])

    cols = [
        "LedgerName",
        "GroupName",
        "ParentName",
        "BS_or_PnL",
        "Type",
        "AffectsGrossProfit",
        "DynamicOpening",
        "T2Dynamic CLB",
        "DynamicClosing",
    ]

    present_cols = [c for c in cols if c in tb_df.columns]
    return tb_df.loc[mask, present_cols].sort_values(["GroupName", "LedgerName"]).reset_index(drop=True)


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
    ytd_tb_df: pd.DataFrame | None = st.session_state.get("ytd_tb_df")

    if overview_vouchers is None:
        overview_vouchers = _voucher_dataframe(_sample_vouchers())
        st.session_state.overview_vouchers_df = overview_vouchers

    overview_tab, table_tab = st.tabs(["Overview", "Table"])

    with table_tab:
        st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
        st.subheader("Dynamic Trial Balance Period")
        fetch_tb = False
        selected_from: date | None = None
        selected_to: date | None = None

        def _period_bounds(label: str, kind: str) -> tuple[date, date]:
            """Return (start, end) dates for the chosen quarter/month in the current fiscal year."""

            today = date.today()
            fy_start = _fiscal_year_start(today)
            fy_year = fy_start.year

            def month_range(year: int, month: int) -> tuple[date, date]:
                last_day = calendar.monthrange(year, month)[1]
                return date(year, month, 1), date(year, month, last_day)

            if kind == "quarter":
                q_map = {
                    "Q1": (4, 6, fy_year),
                    "Q2": (7, 9, fy_year),
                    "Q3": (10, 12, fy_year),
                    "Q4": (1, 3, fy_year + 1),
                }
                start_m, end_m, year_val = q_map[label]
                start_date = date(year_val if start_m != 1 else year_val, start_m, 1)
                end_date = date(year_val if end_m != 12 else year_val, end_m, calendar.monthrange(year_val, end_m)[1])
                return start_date, end_date

            month_map = {
                "April": (fy_year, 4),
                "May": (fy_year, 5),
                "June": (fy_year, 6),
                "July": (fy_year, 7),
                "August": (fy_year, 8),
                "September": (fy_year, 9),
                "October": (fy_year, 10),
                "November": (fy_year, 11),
                "December": (fy_year, 12),
                "January": (fy_year + 1, 1),
                "February": (fy_year + 1, 2),
                "March": (fy_year + 1, 3),
            }
            year_val, month_val = month_map[label]
            return month_range(year_val, month_val)

        if company:
            st.caption("Choose a fiscal quarter or month to auto-set the dynamic trial balance range.")
            q_cols = st.columns(4)
            quarter_clicked = None
            for idx, q_label in enumerate(["Q1", "Q2", "Q3", "Q4"]):
                with q_cols[idx]:
                    if st.button(q_label):
                        quarter_clicked = q_label

            month_labels = [
                "April",
                "May",
                "June",
                "July",
                "August",
                "September",
                "October",
                "November",
                "December",
                "January",
                "February",
                "March",
            ]
            month_cols = st.columns(4)
            month_clicked = None
            for idx, m_label in enumerate(month_labels):
                with month_cols[idx % 4]:
                    if st.button(m_label):
                        month_clicked = m_label

            if quarter_clicked:
                selected_from, selected_to = _period_bounds(quarter_clicked, "quarter")
            elif month_clicked:
                selected_from, selected_to = _period_bounds(month_clicked, "month")

            if selected_from and selected_to:
                fetch_tb = True
                tb_from, tb_to = selected_from, selected_to

        else:
            st.info("Select a company to configure the dynamic trial balance inputs.")

        if fetch_tb and company:
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
                    st.success(
                        f"Dynamic trial balance ready ({len(tb_df):,} ledgers) · Period: {tb_from} to {tb_to}"
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
            st.caption("Dynamic trial balance is hidden; download to view details.")
        else:
            st.info("Fetch the dynamic trial balance to enable downloads and statements.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
        st.subheader("Statements")
        current_tb = st.session_state.get("tb_df")
        col_pl, col_bs = st.columns(2)
        show_pl = col_pl.button("Show Profit & Loss", type="secondary")
        show_bs = col_bs.button("Show Balance Sheet", type="secondary")

        if current_tb is None or current_tb.empty:
            st.info("Build the dynamic trial balance first to view statements.")
        else:
            if show_pl:
                pl_df = _statement_from_tb(current_tb, "p&l")
                if pl_df.empty:
                    st.warning("No Profit & Loss rows found in the dynamic trial balance.")
                else:
                    st.dataframe(pl_df.style.format(precision=2), use_container_width=True, height=480)

            if show_bs:
                bs_df = _statement_from_tb(current_tb, "balance")
                if bs_df.empty:
                    st.warning("No Balance Sheet rows found in the dynamic trial balance.")
                else:
                    st.dataframe(bs_df.style.format(precision=2), use_container_width=True, height=480)

        st.markdown("</div>", unsafe_allow_html=True)

    # Refresh local variables from session state after processing inputs
    tb_df = st.session_state.get("tb_df")
    tb_from = st.session_state.get("tb_from") or tb_from
    tb_to = st.session_state.get("tb_to") or tb_to
    user_ob_input = st.session_state.get("user_ob_input") or user_ob_input or 0.0
    user_cb_input = st.session_state.get("user_cb_input") or user_cb_input or 0.0
    ytd_tb_state = st.session_state.get("ytd_tb_df")
    if ytd_tb_state is not None:
        ytd_tb_df = ytd_tb_state
    # Preserve the existing overview vouchers; avoid truthiness on DataFrame which raises ValueError
    overview_vouchers_state = st.session_state.get("overview_vouchers_df")
    if overview_vouchers_state is not None:
        overview_vouchers = overview_vouchers_state

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
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
        st.subheader("YTD Trial Balance")
        if overview_vouchers is None or overview_vouchers.empty:
            st.info("Load the full Day Book to compute the YTD trial balance.")
        elif company:
            if st.button("Build YTD Trial Balance", type="primary"):
                with st.spinner("Computing YTD trial balance..."):
                    try:
                        ytd_tb_df = _build_ytd_trial_balance(
                            company, host, int(port), overview_vouchers
                        )
                    except Exception as exc:
                        st.error(f"Failed to build YTD trial balance: {exc}")
                    else:
                        st.session_state.ytd_tb_df = ytd_tb_df
                        st.success(f"YTD trial balance ready ({len(ytd_tb_df):,} ledgers)")
        else:
            st.info("Select a company to compute the YTD trial balance.")

        if ytd_tb_df is not None and not ytd_tb_df.empty:
            st.download_button(
                label="Download YTD Trial Balance (Excel)",
                data=_to_excel_bytes(ytd_tb_df),
                file_name=f"YTD_Trial_Balance_{company or 'Sample'}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.dataframe(
                ytd_tb_df.style.format(precision=2),
                use_container_width=True,
                height=480,
            )
        st.markdown("</div>", unsafe_allow_html=True)

        if ytd_tb_df is not None and not ytd_tb_df.empty:
            st.markdown("<div class='app-shell'>", unsafe_allow_html=True)
            st.subheader("Performance Overview (YTD)")
            opening_stock_val = float(user_ob_input or 0.0)
            closing_stock_val = float(user_cb_input or 0.0)
            selected_kpi = _render_ytd_overview_cards(
                ytd_tb_df, opening_stock_val, closing_stock_val, overview_vouchers
            )

            if selected_kpi:
                monthly_series = _build_kpi_monthly_series(
                    selected_kpi,
                    overview_vouchers,
                    ytd_tb_df,
                    opening_stock_val,
                    closing_stock_val,
                )

                st.markdown("<div style='margin-top:10px;'>", unsafe_allow_html=True)
                if monthly_series.empty:
                    st.info("No voucher data available to plot this KPI.")
                else:
                    chart_df = monthly_series.reset_index()
                    chart_df.columns = ["Month", "Amount"]
                    st.bar_chart(chart_df, x="Month", y="Amount")
                st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("</div>", unsafe_allow_html=True)

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


def _build_ytd_trial_balance(
    company: str, host: str, port: int, voucher_df: pd.DataFrame
) -> pd.DataFrame:
    """Compute a YTD trial balance using full Day Book movements."""

    if voucher_df is None or voucher_df.empty:
        raise RuntimeError("Day Book is empty; load vouchers first.")

    ledger_rows = _load_ledger_master(company, host, port, None, None)
    group_rows = _load_group_master(company, host, port, None, None)

    voucher_df = voucher_df.copy()
    voucher_df["Nett"] = voucher_df["Nett"].astype(float)

    ledger_parent_map = {row["LedgerName"]: row["LedgerParent"] for row in ledger_rows}
    ledger_opening_map = {row["LedgerName"]: row["OpeningBalanceNormalized"] for row in ledger_rows}
    group_map = {row["GroupName"]: row for row in group_rows}

    nets_total = voucher_df.groupby("Ledger")["Nett"].sum()

    rows = []
    for ledger_name, opening in ledger_opening_map.items():
        parent = ledger_parent_map.get(ledger_name, "") or "(Unknown)"
        group_info = group_map.get(parent, {})
        bs_pnl = group_info.get("BS_or_PnL", "")
        gtype = group_info.get("Type", "")
        affects_gp = group_info.get("AffectsGrossProfit", "")

        nett_ytd = nets_total.get(ledger_name, 0.0)
        closing = opening + nett_ytd

        rows.append(
            {
                "LedgerName": ledger_name,
                "GroupName": parent,
                "ParentName": group_info.get("ParentName", parent),
                "BS_or_PnL": bs_pnl,
                "Type": gtype,
                "AffectsGrossProfit": affects_gp,
                "OpeningBalance": opening,
                "Nett YTD": nett_ytd,
                "YTDCLB": closing,
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


if __name__ == "__main__":
    main()

