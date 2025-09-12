import sqlite3
import csv

DB = 'business_apps.db'
OUT = 'schema_export.csv'

conn = sqlite3.connect(DB)
c = conn.cursor()
c.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
tables = [r[0] for r in c.fetchall()]
rows = []
for t in tables:
    c.execute(f"PRAGMA table_info({t})")
    cols = c.fetchall()
    for col in cols:
        cid, name, typ, notnull, dflt, pk = col
        rows.append({'table': t, 'cid': cid, 'name': name, 'type': typ, 'notnull': notnull, 'dflt': dflt, 'pk': pk})
conn.close()

with open(OUT, 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=['table','cid','name','type','notnull','dflt','pk'])
    writer.writeheader()
    for r in rows:
        writer.writerow(r)

print('Wrote', OUT)
