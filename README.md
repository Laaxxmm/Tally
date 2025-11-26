# Tally MIS Dashboard

A lightweight Streamlit dashboard that connects to a running Tally instance (via HTTP XML on `127.0.0.1:9000`) and presents client-ready MIS numbers: revenue, expenses, profit and loss, balance sheet snapshot, and best-selling products/services. The dashboard can also fall back to sample data so you can demo the experience without live books.

## Features
- Connects to Tally over the XML/HTTP interface (`http://127.0.0.1:9000` when Tally is open in the background).
- Pulls Day Book entries for a selectable date range, including Sales/Purchase vouchers that post amounts via inventory lines.
- Exports the chart-of-accounts ledger list as an Excel download (Name, Under, Opening Balance Raw, Opening Balance Normalized) without displaying it in the UI, using Dr/Cr-aware parsing so values mirror Tally.
- Exports a Group master extract (GroupName, ParentName, Balance-Sheet/P&L classification, type, and gross-profit flag) so you can audit chart-of-accounts structure.
- Aggregates ledgers into Revenue, Cost of Goods Sold, Expenses, Assets, and Liabilities with heuristics to keep numbers accurate.
- Displays KPIs, balance snapshot, top products/services, and a searchable voucher grid (with debit/credit/nett columns that should sum to zero) for quick investigation.

## Running locally
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Ensure Tally is running with the HTTP XML interface enabled on port `9000`.
3. Start the dashboard:
   ```bash
   streamlit run src/dashboard.py
   ```
4. Use the sidebar to enter host/port (defaults to `127.0.0.1:9000`), click **Connect to Tally** to load companies, choose the company, set your date range, and click **Load from Tally** to fetch live data.

## Module overview
- `src/tally_client.py`: Minimal HTTP XML client for fetching Day Book vouchers and chart-of-accounts ledgers/groups (including Dr/Cr-normalized opening balances and Excel export helpers).
- `src/analytics.py`: Aggregates vouchers into MIS-friendly KPIs (revenue, expenses, gross margin, assets/liabilities, best sellers).
- `src/dashboard.py`: Streamlit UI that connects to Tally, applies analytics, renders KPIs/charts/tables, and offers ledger and group download buttons.

## Group classification cheatsheet
The group master extract leans on Tally's own flags (see `tally_client.classify_type`, `classify_bs_or_pnl`, and `determine_affects_gross_profit`):
- `ISREVENUE` from the group master dictates Balance Sheet vs P&L (Yes → P&L, No → Balance Sheet). When `ISREVENUE` is absent, the group nature is used to default sensibly.
- `NATUREOFGROUP` is used to tag groups as Asset/Liability/Income/Expense (e.g., "Misc. Expenses (ASSET)" stays Asset → Balance Sheet, "Purchase Accounts" stays Expense → P&L, "Indirect Expenses" stays Expense but does not affect gross profit).
- `AFFECTSGROSSPROFIT` from Tally controls the gross-profit flag; only when that flag is missing do keyword fallbacks apply.

## Security note
The dashboard trusts the Tally instance it connects to. Run it on a secure network and avoid exposing the Tally port publicly.

