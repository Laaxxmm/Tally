import sqlite3
import pandas as pd
from datetime import date, datetime
import tally_client

class DataManager:
    def __init__(self, db_path="tally.db"):
        self.db_path = db_path
        self.init_db()

    def init_db(self):
        """Initialize the SQLite database schema."""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        # Vouchers table
        c.execute('''CREATE TABLE IF NOT EXISTS vouchers (
            voucher_number TEXT,
            date DATE,
            voucher_type TEXT,
            ledger_name TEXT,
            amount REAL,
            is_debit BOOLEAN,
            narration TEXT
        )''')
        
        # Ledgers Master
        c.execute('''CREATE TABLE IF NOT EXISTS ledgers (
            name TEXT PRIMARY KEY,
            parent TEXT,
            opening_balance REAL
        )''')
        
        # Groups Master
        c.execute('''CREATE TABLE IF NOT EXISTS groups (
            name TEXT PRIMARY KEY,
            parent TEXT,
            bs_or_pnl TEXT,
            type TEXT,
            affects_gp TEXT
        )''')
        
        # Sync Status
        c.execute('''CREATE TABLE IF NOT EXISTS sync_status (
            last_sync TIMESTAMP
        )''')
        
        conn.commit()
        conn.close()

    def get_last_sync(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT last_sync FROM sync_status ORDER BY last_sync DESC LIMIT 1")
        row = c.fetchone()
        conn.close()
        return row[0] if row else None

    def sync_data(self, company, host, port):
        """Fetch data from Tally and replace local cache."""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        # 1. Sync Groups
        groups = tally_client.fetch_group_master(company, host, port)
        c.execute("DELETE FROM groups")
        c.executemany("INSERT INTO groups VALUES (?, ?, ?, ?, ?)", 
                      [(g['GroupName'], g['ParentName'], g['BS_or_PnL'], g['Type'], g['AffectsGrossProfit']) for g in groups])
        
        # 2. Sync Ledgers
        ledgers = tally_client.fetch_ledger_master(company, host, port)
        c.execute("DELETE FROM ledgers")
        c.executemany("INSERT INTO ledgers VALUES (?, ?, ?)", 
                      [(l['LedgerName'], l['LedgerParent'], l['OpeningBalanceNormalized']) for l in ledgers])
        
        # 3. Sync Vouchers (Full Daybook for now)
        # In a real scenario, we might do incremental sync based on date
        vouchers = tally_client.fetch_daybook(company, None, None, host, port)
        c.execute("DELETE FROM vouchers")
        
        voucher_rows = []
        for v in vouchers:
            for entry in v.ledger_entries:
                # Store signed amount: Debit is positive, Credit is negative
                signed_amount = entry.amount if entry.is_debit else -entry.amount
                voucher_rows.append((
                    v.voucher_number,
                    v.date,
                    v.voucher_type,
                    entry.ledger_name,
                    signed_amount,
                    entry.is_debit,
                    v.narration
                ))
        
        c.executemany("INSERT INTO vouchers VALUES (?, ?, ?, ?, ?, ?, ?)", voucher_rows)
        
        # Update Sync Status
        c.execute("DELETE FROM sync_status")
        c.execute("INSERT INTO sync_status VALUES (?)", (datetime.now(),))
        
        conn.commit()
        conn.close()

    def get_kpi_data(self, start_date, end_date):
        """Get aggregated KPI data for the given period."""
        conn = sqlite3.connect(self.db_path)
        
        # We need to join vouchers with ledgers and groups to filter by type
        query = """
        SELECT 
            g.type,
            g.affects_gp,
            SUM(v.amount) as total
        FROM vouchers v
        JOIN ledgers l ON v.ledger_name = l.name
        JOIN groups g ON l.parent = g.name
        WHERE v.date BETWEEN ? AND ?
        GROUP BY g.type, g.affects_gp
        """
        
        df = pd.read_sql_query(query, conn, params=(start_date, end_date))
        conn.close()
        
        # Process results
        # Note: The dataframe columns will be lowercase 'type' and 'affects_gp'
        revenue = abs(df[(df['type'] == 'Income') & (df['affects_gp'] == 'Yes')]['total'].sum())
        direct_expense = df[(df['type'] == 'Expense') & (df['affects_gp'] == 'Yes')]['total'].sum()
        
        # COGS logic might need specific group filtering (e.g. "Purchase Accounts")
        # For now, using Direct Expenses as proxy for COGS components excluding stock
        cogs = direct_expense 
        
        gross_profit = revenue - cogs
        
        indirect_expense = df[(df['type'] == 'Expense') & (df['affects_gp'] == 'No')]['total'].sum()
        indirect_income = abs(df[(df['type'] == 'Income') & (df['affects_gp'] == 'No')]['total'].sum())
        
        net_profit = gross_profit + indirect_income - indirect_expense
        
        return {
            "revenue": revenue,
            "cogs": cogs,
            "gross_profit": gross_profit,
            "net_profit": net_profit,
            "opex": indirect_expense
        }

    def get_monthly_trend(self, kpi_type, year):
        """Get monthly trend for a specific KPI."""
        conn = sqlite3.connect(self.db_path)
        
        # Define filters based on KPI
        filters = ""
        if kpi_type == "revenue":
            filters = "g.type = 'Income' AND g.affects_gp = 'Yes'"
        elif kpi_type == "cogs":
            filters = "g.type = 'Expense' AND g.affects_gp = 'Yes'"
        elif kpi_type == "opex":
            filters = "g.type = 'Expense' AND g.affects_gp = 'No'"
            
        query = f"""
        SELECT 
            strftime('%m', v.date) as month,
            SUM(v.amount) as total
        FROM vouchers v
        JOIN ledgers l ON v.ledger_name = l.name
        JOIN groups g ON l.parent = g.name
        WHERE strftime('%Y', v.date) = ? AND {filters}
        GROUP BY month
        ORDER BY month
        """
        
        df = pd.read_sql_query(query, conn, params=(str(year),))
        conn.close()
        
        # Ensure all months are present
        full_months = pd.DataFrame({'month': [f"{i:02d}" for i in range(1, 13)]})
        df = full_months.merge(df, on='month', how='left').fillna(0)
        
        # Flip signs for income/revenue
        if kpi_type == "revenue":
            df['total'] = df['total'].abs()
            
        return df
