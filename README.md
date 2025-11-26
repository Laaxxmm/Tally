# Tally MIS Dashboard

A lightweight Streamlit dashboard that connects to a running Tally instance (via HTTP XML on `127.0.0.1:9000`) and presents client-ready MIS numbers: revenue, expenses, profit and loss, balance sheet snapshot, and best-selling products/services. The dashboard can also fall back to sample data so you can demo the experience without live books.

## Features
- Connects to Tally over the XML/HTTP interface (`http://127.0.0.1:9000` when Tally is open in the background).
- Pulls Day Book entries for a selectable date range, including Sales/Purchase vouchers that post amounts via inventory lines.
- Exports the chart-of-accounts ledger list (Name, Under, Opening Balance) as a CSV download without displaying it in the UI, using Trial Balance data so Dr/Cr opening values match what Tally shows.
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
- `src/tally_client.py`: Minimal HTTP XML client for fetching Day Book vouchers and chart-of-accounts ledgers.
- `src/analytics.py`: Aggregates vouchers into MIS-friendly KPIs (revenue, expenses, gross margin, assets/liabilities, best sellers).
- `src/dashboard.py`: Streamlit UI that connects to Tally, applies analytics, renders KPIs/charts/tables, and offers a ledger download button.

## Security note
The dashboard trusts the Tally instance it connects to. Run it on a secure network and avoid exposing the Tally port publicly.

