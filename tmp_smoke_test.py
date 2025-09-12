import sqlite3
import pprint
from datetime import datetime
import database

DB = database.DB_NAME

# Helper to print table rows

def print_table(table):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute(f"SELECT * FROM {table}")
    rows = c.fetchall()
    print(f"\nTABLE: {table} ({len(rows)} rows)")
    for r in rows:
        pprint.pprint(r)
    conn.close()


def run_smoke():
    conn = sqlite3.connect(DB)
    c = conn.cursor()

    # Insert a test business unit
    # Use INSERT OR IGNORE to be idempotent if smoke test ran before
    c.execute("INSERT OR IGNORE INTO business_units (name) VALUES (?)", ('SmokeDept',))
    c.execute("SELECT id FROM business_units WHERE name = ?", ('SmokeDept',))
    unit_id = c.fetchone()[0]
    print('Business unit id=', unit_id)

    # Insert an application linked to that business unit
    name = 'SmokeApp'
    vendor = 'SmokeVendor'
    factors = {'Score': 7, 'Need': 6, 'Criticality': 8, 'Installed': 1, 'DisasterRecovery': 2, 'Safety': 3, 'Security': 4, 'Monetary': 5, 'CustomerService': 6}
    last_mod = datetime.utcnow().isoformat()
    # compute risk_score using new formula (10 - score) * criticality
    risk_score = (10 - int(factors['Score'])) * int(factors['Criticality'])

    # Insert or ignore application (idempotent)
    c.execute('''INSERT OR IGNORE INTO applications (name, vendor, score, need, criticality, installed, disaster_recovery, safety, security, monetary, customer_service, notes, risk_score, last_modified)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                 (name, vendor, factors['Score'], factors['Need'], factors['Criticality'], factors['Installed'], factors['DisasterRecovery'], factors['Safety'], factors['Security'], factors['Monetary'], factors['CustomerService'], 'Smoke notes', risk_score, last_mod))
    # retrieve the app id (existing or new)
    c.execute('SELECT id FROM applications WHERE name = ? AND vendor = ?', (name, vendor))
    app_id = c.fetchone()[0]
    print('Application id=', app_id)

    # Link to business unit
    c.execute('INSERT OR IGNORE INTO application_business_units (app_id, unit_id) VALUES (?, ?)', (app_id, unit_id))

    # Commit and close before calling higher-level DB helpers to avoid locking
    conn.commit()
    conn.close()

    # Insert an integration via database API
    fields = {
        'name': 'SmokeIntegration',
        'vendor': 'SmokeVendor',
        'score': 5,
        'need': 4,
        'criticality': 3,
        'installed': 2,
        'disaster_recovery': 1,
        'safety': 2,
        'security': 3,
        'monetary': 4,
        'customer_service': 5,
        'notes': 'Smoke integration notes'
    }
    integration_id = database.add_system_integration(app_id, fields)
    print('Inserted integration id=', integration_id)

    # Print tables
    print_table('business_units')
    print_table('applications')
    print_table('application_business_units')
    print_table('system_integrations')

    # Call get_system_integrations and show GUI-assembled values
    integrations = database.get_system_integrations(app_id)
    print('\nget_system_integrations returned:')
    for r in integrations:
        pprint.pprint(r)
        # emulate GUI assembly
        name = str(r[2] if r[2] is not None else '')
        vendor = str(r[3] if r[3] is not None else '')
        ratings = [int(r[i]) if r[i] is not None else 0 for i in (4,5,6,7,8,9,10,11,12)]
        risk_text = 'N/A'
        try:
            if r[14] is not None:
                risk_text = f"{float(r[14]):.1f} ({database.dr_priority_band(float(r[14]))})"
        except Exception:
            pass
        last_mod = ''
        try:
            if r[15]:
                last_mod = str(r[15])
        except Exception:
            pass
        values = [name, vendor] + ratings + [risk_text, last_mod]
        print('GUI would show values:')
        pprint.pprint(values)

if __name__ == '__main__':
    run_smoke()
