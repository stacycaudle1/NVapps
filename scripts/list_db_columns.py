import sqlite3

DB = 'business_apps.db'

def list_columns():
    conn = sqlite3.connect(DB)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
    tables = [t[0] for t in cursor.fetchall()]
    for table in tables:
        print(f"Table: {table}")
        cursor.execute(f"PRAGMA table_info({table})")
        cols = cursor.fetchall()
        for col in cols:
            # col tuple: (cid, name, type, notnull, dflt_value, pk)
            print(f"  - {col[1]} ({col[2]})")
        print()
    conn.close()

if __name__ == '__main__':
    list_columns()
