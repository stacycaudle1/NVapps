"""Utility: Seed minimal demo data if database empty.
Does not alter schema; safe to run multiple times (idempotent for demo names).
"""
import database

def ensure_demo():
    database.initialize_database()
    conn = database.connect_db()
    cur = conn.cursor()
    cur.execute('SELECT COUNT(*) FROM applications')
    app_count = cur.fetchone()[0]
    if app_count > 0:
        conn.close()
        return False
    # Insert a couple of demo apps
    cur.execute("INSERT INTO business_units (name) VALUES ('IT') ON CONFLICT(name) DO NOTHING")
    cur.execute("INSERT INTO business_units (name) VALUES ('Finance') ON CONFLICT(name) DO NOTHING")
    cur.execute("INSERT INTO categories (name) VALUES ('Technology') ON CONFLICT(name) DO NOTHING")
    cur.execute("INSERT INTO categories (name) VALUES ('Operations') ON CONFLICT(name) DO NOTHING")
    # Fetch ids
    bu_it = cur.execute("SELECT id FROM business_units WHERE name='IT'").fetchone()[0]
    bu_fin = cur.execute("SELECT id FROM business_units WHERE name='Finance'").fetchone()[0]
    cat_tech = cur.execute("SELECT id FROM categories WHERE name='Technology'").fetchone()[0]
    cat_ops = cur.execute("SELECT id FROM categories WHERE name='Operations'").fetchone()[0]
    cur.execute("INSERT INTO applications (division, vendor, score, need, criticality, installed, disaster_recovery, safety, security, monetary, customer_service, notes, risk_score, last_modified) VALUES ('DemoApp1','VendorX',5,5,6,7,4,5,5,5,5,'Demo notes', (10-5)*6, CURRENT_TIMESTAMP)")
    app1 = cur.lastrowid
    cur.execute("INSERT INTO applications (division, vendor, score, need, criticality, installed, disaster_recovery, safety, security, monetary, customer_service, notes, risk_score, last_modified) VALUES ('DemoApp2','VendorY',4,6,7,8,3,4,6,4,7,'More notes', (10-4)*7, CURRENT_TIMESTAMP)")
    app2 = cur.lastrowid
    # Link business units
    cur.execute('INSERT OR IGNORE INTO application_business_units (app_id, unit_id) VALUES (?,?)', (app1, bu_it))
    cur.execute('INSERT OR IGNORE INTO application_business_units (app_id, unit_id) VALUES (?,?)', (app2, bu_fin))
    # Link categories
    for cid in (cat_tech, cat_ops):
        cur.execute('INSERT OR IGNORE INTO application_categories (app_id, category_id) VALUES (?,?)', (app1, cid))
    cur.execute('INSERT OR IGNORE INTO application_categories (app_id, category_id) VALUES (?,?)', (app2, cat_tech))
    # Add integrations
    cur.execute("INSERT INTO system_integrations (parent_app_id, name, vendor, score, need, criticality, installed, disaster_recovery, safety, security, monetary, customer_service, notes, risk_score, last_modified) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)", (app1,'DemoInt1','IVendor',6,5,5,5,5,5,5,5,5,'First integration',(10-6)*5))
    int1 = cur.lastrowid
    cur.execute("INSERT INTO system_integrations (parent_app_id, name, vendor, score, need, criticality, installed, disaster_recovery, safety, security, monetary, customer_service, notes, risk_score, last_modified) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)", (app2,'DemoInt2','JVendor',5,6,7,6,5,5,5,6,6,'Second integration',(10-5)*7))
    int2 = cur.lastrowid
    # Link integration categories
    cur.execute('INSERT OR IGNORE INTO integration_categories (integration_id, category_id) VALUES (?,?)', (int1, cat_tech))
    cur.execute('INSERT OR IGNORE INTO integration_categories (integration_id, category_id) VALUES (?,?)', (int2, cat_tech))
    conn.commit()
    conn.close()
    return True

if __name__ == '__main__':
    created = ensure_demo()
    print('Seeded demo data' if created else 'Database already had data')