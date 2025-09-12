import sqlite3

DB = 'business_apps.db'

def main():
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = [r[0] for r in c.fetchall()]
    for t in tables:
        print('\nTABLE:', t)
        c.execute(f"PRAGMA table_info({t})")
        cols = c.fetchall()
        if not cols:
            print('  (no columns)')
        for col in cols:
            # col format: (cid, name, type, notnull, dflt_value, pk)
            print(f"  cid={col[0]} name={col[1]} type={col[2]} notnull={col[3]} dflt={col[4]} pk={col[5]}")
    conn.close()

if __name__ == '__main__':
    main()
