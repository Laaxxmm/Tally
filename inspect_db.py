import sqlite3
import pandas as pd

def inspect_db():
    try:
        conn = sqlite3.connect("tally.db")
        c = conn.cursor()
        
        print("--- Table Counts ---")
        tables = ["vouchers", "ledgers", "groups", "sync_status"]
        for t in tables:
            try:
                count = c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                print(f"{t}: {count} rows")
            except Exception as e:
                print(f"{t}: Error - {e}")
                
        print("\n--- Date Range in Vouchers ---")
        try:
            min_date = c.execute("SELECT MIN(date) FROM vouchers").fetchone()[0]
            max_date = c.execute("SELECT MAX(date) FROM vouchers").fetchone()[0]
            print(f"Min Date: {min_date}")
            print(f"Max Date: {max_date}")
        except:
            print("Could not fetch dates")

        print("\n--- Sample Groups ---")
        try:
            groups = pd.read_sql("SELECT * FROM groups LIMIT 5", conn)
            print(groups)
        except:
            print("Could not fetch groups")

        print("\n--- Sample Vouchers ---")
        try:
            vouchers = pd.read_sql("SELECT * FROM vouchers LIMIT 5", conn)
            print(vouchers)
        except:
            print("Could not fetch vouchers")
            
        conn.close()
    except Exception as e:
        print(f"DB Error: {e}")

if __name__ == "__main__":
    inspect_db()
