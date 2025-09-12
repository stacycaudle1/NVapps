import sqlite3, shutil, os

DB = 'business_apps.db'
BACKUP = 'business_apps.db.bak'

print('Backing up DB to', BACKUP)
shutil.copyfile(DB, BACKUP)

conn = sqlite3.connect(DB)
c = conn.cursor()

# Read current columns
c.execute("PRAGMA table_info(applications)")
cols = c.fetchall()
print('Current applications columns:', cols)

# Define desired new schema for applications
new_schema = '''CREATE TABLE applications_new (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    vendor TEXT,
    score INTEGER DEFAULT 0,
    need INTEGER DEFAULT 0,
    criticality INTEGER DEFAULT 0,
    installed INTEGER DEFAULT 0,
    disaster_recovery INTEGER DEFAULT 0,
    safety INTEGER DEFAULT 0,
    security INTEGER DEFAULT 0,
    monetary INTEGER DEFAULT 0,
    customer_service INTEGER DEFAULT 0,
    notes TEXT,
    user_id INTEGER,
    risk_score REAL,
    last_modified TEXT
)
'''

print('Creating new table with desired schema')
c.execute(new_schema)

# Build select list from existing columns to map/cast as needed
# We'll select existing columns if present, else use defaults
existing_names = [col[1].lower() for col in cols]

select_parts = []
for col in ['id','name','vendor','score','need','criticality','installed','disaster_recovery','safety','security','monetary','customer_service','notes','user_id','risk_score','last_modified']:
    if col in existing_names:
        # For risk_score, cast to REAL explicitly
        if col == 'risk_score':
            select_parts.append(f"CAST({col} AS REAL) AS risk_score")
        elif col == 'disaster_recovery':
            # ensure disaster_recovery becomes integer (try casting)
            select_parts.append(f"CAST({col} AS INTEGER) AS disaster_recovery")
        else:
            select_parts.append(col)
    else:
        # default values for missing columns
        if col == 'risk_score':
            select_parts.append('NULL AS risk_score')
        elif col == 'disaster_recovery':
            select_parts.append('0 AS disaster_recovery')
        else:
            select_parts.append('NULL')

select_sql = ', '.join(select_parts)
copy_sql = f"INSERT INTO applications_new (id, name, vendor, score, need, criticality, installed, disaster_recovery, safety, security, monetary, customer_service, notes, user_id, risk_score, last_modified) SELECT {select_sql} FROM applications"

print('Copying data...')
c.execute(copy_sql)
conn.commit()

print('Dropping old applications table and renaming new')
c.execute('DROP TABLE applications')
c.execute('ALTER TABLE applications_new RENAME TO applications')
conn.commit()

print('New schema:')
c.execute('PRAGMA table_info(applications)')
for row in c.fetchall():
    print(row)

conn.close()
print('Migration complete. Backup at', BACKUP)
