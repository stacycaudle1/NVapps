import sqlite3

DB='business_apps.db'
conn=sqlite3.connect(DB)
c=conn.cursor()
for t in ['applications','system_integrations','users','business_units','application_business_units','application_departments','app_departments','departments']:
    try:
        c.execute(f'SELECT COUNT(*) FROM {t}')
        print(f"{t}:", c.fetchone()[0])
    except Exception as e:
        print(f"{t}: ERROR ({e})")
conn.close()
