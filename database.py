def generate_smoke_test_data():
    """
    Populate the database with sample data for reports and UI validation.
    Creates business units, applications, departments, and system integrations with realistic values.
    """
    import random
    from datetime import datetime, timezone
    import time
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            # Create business units
            conn = connect_db()
            c = conn.cursor()
            bu_names = ["Finance", "HR", "IT", "Operations", "Sales"]
            c.executemany("INSERT OR IGNORE INTO business_units (name) VALUES (?)", [(n,) for n in bu_names])
            conn.commit()
            conn.close()

            # Create departments
            conn = connect_db()
            c = conn.cursor()
            dept_names = ["Accounting", "Recruiting", "Support", "Logistics", "Marketing"]
            c.executemany("INSERT OR IGNORE INTO departments (name) VALUES (?)", [(n,) for n in dept_names])
            conn.commit()
            conn.close()

            # Create applications
            app_names = ["ERP System", "Payroll", "CRM", "Inventory", "Helpdesk"]
            for i, app in enumerate(app_names):
                vendor = random.choice(["Oracle", "SAP", "Microsoft", "Custom", "OpenSource"])
                factors = {k: random.randint(1, 10) for k in ['score', 'need', 'criticality', 'installed', 'disaster_recovery', 'safety', 'security', 'monetary', 'customer_service']}
                notes = f"Sample notes for {app}"
                dept_ids = [random.randint(1, len(dept_names))]
                add_application(app, vendor, factors, dept_ids, notes)

            # Link applications to business units
            conn = connect_db()
            c = conn.cursor()
            c.execute("SELECT id FROM applications")
            app_ids = [r[0] for r in c.fetchall()]
            c.execute("SELECT id FROM business_units")
            bu_ids = [r[0] for r in c.fetchall()]
            for app_id in app_ids:
                bu_id = random.choice(bu_ids)
                c.execute("INSERT OR IGNORE INTO application_business_units (app_id, unit_id) VALUES (?, ?)", (app_id, bu_id))
            conn.commit()
            conn.close()

            # Create system integrations for each app
            for app_id in app_ids:
                for j in range(random.randint(2, 4)):
                    fields = {
                        'name': f"Integration {j+1} for App {app_id}",
                        'vendor': random.choice(["AWS", "Azure", "Google", "Internal"]),
                        'score': random.randint(1, 10),
                        'need': random.randint(1, 10),
                        'criticality': random.randint(1, 10),
                        'installed': random.randint(1, 10),
                        'disaster_recovery': random.randint(1, 10),
                        'safety': random.randint(1, 10),
                        'security': random.randint(1, 10),
                        'monetary': random.randint(1, 10),
                        'customer_service': random.randint(1, 10),
                        'notes': f"Integration notes {j+1}",
                        'risk_score': random.uniform(10, 100),
                        'last_modified': datetime.now(timezone.utc).isoformat()
                    }
                    add_system_integration(app_id, fields)
            return
        except sqlite3.OperationalError as e:
            if 'database is locked' in str(e):
                time.sleep(0.5)
                continue
            else:
                raise
    raise RuntimeError('Failed to generate smoke test data: database is locked after multiple attempts')
import sqlite3
from datetime import datetime
from dataclasses import dataclass, field
from typing import Dict, Optional
import math

DB_NAME = 'business_apps.db'


def _to_int_safe(x, default=0):
    try:
        return int(float(x))
    except Exception:
        return default

def dr_priority_band(total_score: float) -> str:
    # Map priority bands to the numeric risk score (0-100) used by the app.
    # New bands: Low: 1-49, Med: 50-69, High: 70-100
    try:
        score = float(total_score)
    except Exception:
        return "Low"

    if score >= 70:
        return "High"
    if score >= 50:
        return "Med"
    # treat 0 and any non-positive values as Low
    return "Low"


def connect_db():
    """Return a sqlite3 connection with Row factory for named access."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    # Ensure applications table exists and has required columns
    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='applications'")
    exists = c.fetchone() is not None
    if not exists:
        # create table with full schema including factor columns and last_modified
        c.execute('''CREATE TABLE applications (
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
        )''')
    else:
        # table exists: ensure factor columns, vendor, and last_modified exist
        c.execute("PRAGMA table_info(applications)")
        columns = [col[1].lower() for col in c.fetchall()]
        factor_columns = ['score', 'need', 'criticality', 'installed', 'disaster_recovery', 'safety', 'security', 'monetary', 'customer_service']
        for col in factor_columns:
            if col not in columns:
                c.execute(f'ALTER TABLE applications ADD COLUMN {col} INTEGER DEFAULT 0')
        if 'vendor' not in columns:
            c.execute('ALTER TABLE applications ADD COLUMN vendor TEXT')
        if 'last_modified' not in columns:
            c.execute("ALTER TABLE applications ADD COLUMN last_modified TEXT")
        # Rename column 'stability' to 'score' in the 'applications' table if it exists
        c.execute("PRAGMA table_info(applications)")
        columns = [col[1].lower() for col in c.fetchall()]
        if 'stability' in columns and 'score' not in columns:
            c.execute("ALTER TABLE applications RENAME COLUMN stability TO score")
        # If 'stability' still exists, remove it by rebuilding the table without that column
        c.execute("PRAGMA table_info(applications)")
        cols_info = c.fetchall()
        col_names_lower = [col[1].lower() for col in cols_info]
        if 'stability' in col_names_lower:
            # Build column defs for new table excluding 'stability'
            cols_keep = [col for col in cols_info if col[1].lower() != 'stability']
            col_defs = []
            for col in cols_keep:
                name = col[1]
                typ = col[2] or 'TEXT'
                notnull = ' NOT NULL' if col[3] else ''
                dflt = f" DEFAULT {col[4]}" if col[4] is not None else ''
                pk = ' PRIMARY KEY' if col[5] else ''
                # handle common INTEGER PRIMARY KEY AUTOINCREMENT pattern
                if col[5] and typ.upper() == 'INTEGER':
                    col_defs.append(f"{name} {typ}{pk}{dflt}")
                else:
                    col_defs.append(f"{name} {typ}{notnull}{dflt}{pk}")
            cols_sql = ', '.join(col_defs)
            # Create a new table, copy data, and replace old table
            c.execute(f"CREATE TABLE IF NOT EXISTS applications_new ({cols_sql})")
            keep_names = ', '.join([f'{col[1]}' for col in cols_keep])
            c.execute(f"INSERT INTO applications_new ({keep_names}) SELECT {keep_names} FROM applications")
            c.execute("DROP TABLE applications")
            c.execute("ALTER TABLE applications_new RENAME TO applications")
            # Refresh PRAGMA info
            c.execute("PRAGMA table_info(applications)")
            cols_info = c.fetchall()
            columns = [col[1].lower() for col in cols_info]
            if 'stability' in columns:
                raise RuntimeError("Migration failed: 'stability' column still exists")
            print("DEBUG: Migration complete, 'stability' column removed from applications table")

    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS business_units (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS application_business_units (
        app_id INTEGER,
        unit_id INTEGER,
        PRIMARY KEY (app_id, unit_id),
        FOREIGN KEY(app_id) REFERENCES applications(id),
        FOREIGN KEY(unit_id) REFERENCES business_units(id)
    )''')
    
    # Create table for system sub-integrations if it doesn't exist
    c.execute('''CREATE TABLE IF NOT EXISTS system_integrations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        parent_app_id INTEGER NOT NULL,
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
        risk_score REAL,
        last_modified TEXT,
        FOREIGN KEY(parent_app_id) REFERENCES applications(id)
    )''')
    # Rename column 'stability' to 'score' in the 'system_integrations' table if it exists
    c.execute("PRAGMA table_info(system_integrations)")
    sys_cols = [col[1].lower() for col in c.fetchall()]
    if 'stability' in sys_cols and 'score' not in sys_cols:
        try:
            c.execute("ALTER TABLE system_integrations RENAME COLUMN stability TO score")
            print("DEBUG: Renamed system_integrations.stability -> score")
        except Exception:
            # If RENAME COLUMN not supported, rebuild table with 'score' instead of 'stability'
            c.execute("PRAGMA table_info(system_integrations)")
            cols_info = c.fetchall()
            cols_keep = []
            for col in cols_info:
                if col[1].lower() == 'stability':
                    # convert entry to 'score' with same defs
                    cols_keep.append((col[0], 'score', col[2], col[3], col[4], col[5]))
                else:
                    cols_keep.append(col)
            col_defs = []
            for col in cols_keep:
                name = col[1]
                typ = col[2] or 'TEXT'
                notnull = ' NOT NULL' if col[3] else ''
                dflt = f" DEFAULT {col[4]}" if col[4] is not None else ''
                pk = ' PRIMARY KEY' if col[5] else ''
                if col[5] and typ.upper() == 'INTEGER':
                    col_defs.append(f"{name} {typ}{pk}{dflt}")
                else:
                    col_defs.append(f"{name} {typ}{notnull}{dflt}{pk}")
            cols_sql = ', '.join(col_defs)
            c.execute(f"CREATE TABLE IF NOT EXISTS system_integrations_new ({cols_sql})")
            keep_names = ', '.join([f'{col[1]}' for col in cols_keep])
            # map stability -> score in select
            select_cols = []
            for col in cols_info:
                if col[1].lower() == 'stability':
                    select_cols.append('stability AS score')
                else:
                    select_cols.append(col[1])
            select_sql = ', '.join(select_cols)
            c.execute(f"INSERT INTO system_integrations_new ({keep_names}) SELECT {select_sql} FROM system_integrations")
            c.execute("DROP TABLE system_integrations")
            c.execute("ALTER TABLE system_integrations_new RENAME TO system_integrations")
            print("DEBUG: Rebuilt system_integrations with 'score' column (from stability)")
    else:
        # If stability isn't present but score is missing, add score column
        if 'score' not in sys_cols:
            c.execute("ALTER TABLE system_integrations ADD COLUMN score INTEGER DEFAULT 0")
            print("DEBUG: Added 'score' column to system_integrations (default 0)")

    # Consolidate duplicate/legacy columns into canonical snake_case names
    # Mapping: canonical -> legacy
    consolidate_map = {
        'disaster_recovery': 'disasterrecovery',
        'customer_service': 'customerservice'
    }

    def _consolidate_table(table_name):
        c.execute(f"PRAGMA table_info({table_name})")
        cols_info = c.fetchall()
        cols = [col[1].lower() for col in cols_info]
        needs_rebuild = False
        for canonical, legacy in consolidate_map.items():
            if legacy in cols and canonical in cols:
                # copy legacy values into canonical where canonical is NULL or empty
                try:
                    # Try numeric/text-agnostic update using COALESCE
                    c.execute(f"UPDATE {table_name} SET {canonical} = CASE WHEN ({canonical} IS NULL OR {canonical} = '') THEN {legacy} ELSE {canonical} END")
                except Exception:
                    # fallback: try without empty-string check
                    c.execute(f"UPDATE {table_name} SET {canonical} = CASE WHEN ({canonical} IS NULL) THEN {legacy} ELSE {canonical} END")
                needs_rebuild = True
            elif legacy in cols and canonical not in cols:
                # rename legacy -> canonical if possible
                try:
                    c.execute(f"ALTER TABLE {table_name} RENAME COLUMN {legacy} TO {canonical}")
                    print(f"DEBUG: Renamed {table_name}.{legacy} -> {canonical}")
                except Exception:
                    # rebuild table mapping legacy AS canonical
                    c.execute(f"PRAGMA table_info({table_name})")
                    cols_info = c.fetchall()
                    cols_keep = []
                    for col in cols_info:
                        if col[1].lower() == legacy:
                            cols_keep.append((col[0], canonical, col[2], col[3], col[4], col[5]))
                        else:
                            cols_keep.append(col)
                    col_defs = []
                    for col in cols_keep:
                        name = col[1]
                        typ = col[2] or 'TEXT'
                        notnull = ' NOT NULL' if col[3] else ''
                        dflt = f" DEFAULT {col[4]}" if col[4] is not None else ''
                        pk = ' PRIMARY KEY' if col[5] else ''
                        if col[5] and typ.upper() == 'INTEGER':
                            col_defs.append(f"{name} {typ}{pk}{dflt}")
                        else:
                            col_defs.append(f"{name} {typ}{notnull}{dflt}{pk}")
                    cols_sql = ', '.join(col_defs)
                    c.execute(f"CREATE TABLE IF NOT EXISTS {table_name}_new ({cols_sql})")
                    keep_names = ', '.join([f'{col[1]}' for col in cols_keep])
                    select_cols = []
                    for col in cols_info:
                        if col[1].lower() == legacy:
                            select_cols.append(f"{legacy} AS {canonical}")
                        else:
                            select_cols.append(col[1])
                    select_sql = ', '.join(select_cols)
                    c.execute(f"INSERT INTO {table_name}_new ({keep_names}) SELECT {select_sql} FROM {table_name}")
                    c.execute(f"DROP TABLE {table_name}")
                    c.execute(f"ALTER TABLE {table_name}_new RENAME TO {table_name}")
                    print(f"DEBUG: Rebuilt {table_name} with {legacy} renamed to {canonical}")
                    return
                    
        # If both canonical existed and we copied values from legacy, drop legacy columns by rebuilding
        if needs_rebuild:
            c.execute(f"PRAGMA table_info({table_name})")
            cols_info = c.fetchall()
            # remove legacy columns from cols_info
            cols_keep = [col for col in cols_info if col[1].lower() not in consolidate_map.values()]
            col_defs = []
            for col in cols_keep:
                name = col[1]
                typ = col[2] or 'TEXT'
                notnull = ' NOT NULL' if col[3] else ''
                dflt = f" DEFAULT {col[4]}" if col[4] is not None else ''
                pk = ' PRIMARY KEY' if col[5] else ''
                if col[5] and typ.upper() == 'INTEGER':
                    col_defs.append(f"{name} {typ}{pk}{dflt}")
                else:
                    col_defs.append(f"{name} {typ}{notnull}{dflt}{pk}")
            cols_sql = ', '.join(col_defs)
            c.execute(f"CREATE TABLE IF NOT EXISTS {table_name}_new ({cols_sql})")
            keep_names = ', '.join([f'{col[1]}' for col in cols_keep])
            c.execute(f"INSERT INTO {table_name}_new ({keep_names}) SELECT {keep_names} FROM {table_name}")
            c.execute(f"DROP TABLE {table_name}")
            c.execute(f"ALTER TABLE {table_name}_new RENAME TO {table_name}")
            print(f"DEBUG: Consolidated legacy columns on {table_name} and removed duplicates")

    # Run consolidation for target tables
    _consolidate_table('applications')
    _consolidate_table('system_integrations')

    conn.commit()
    conn.close()

def calculate_business_risk(app_row):
    """Compute risk for an application row-like object.

    Supports sqlite3.Row (named access) or legacy tuple access by index.
    Returns (risk_score, priority_band).
    """
    def _get(row, key, idx, default=0):
        try:
            # try dict-like/key access first
            if row is None:
                return default
            return _to_int_safe(row[key], default)
        except Exception:
            try:
                return _to_int_safe(row[idx], default)
            except Exception:
                return default

    score = _get(app_row, 'score', 3, 0)
    criticality = _get(app_row, 'criticality', 5, 0)

    computed_risk = (10 - score) * criticality
    priority = dr_priority_band(computed_risk)
    return computed_risk, priority

def link_app_to_departments(app_id, dept_ids):
    print(f"DEBUG: Linking app_id {app_id} to dept_ids {dept_ids}")
    conn = connect_db()
    c = conn.cursor()
    for dept_id in dept_ids:
        c.execute('INSERT OR IGNORE INTO application_business_units (app_id, unit_id) VALUES (?, ?)', (app_id, dept_id))
    conn.commit()
    conn.close()

def get_app_departments(app_id):
    conn = connect_db()
    c = conn.cursor()
    c.execute('SELECT d.name FROM business_units d JOIN application_business_units ad ON d.id = ad.unit_id WHERE ad.app_id = ?', (app_id,))
    depts = [row[0] for row in c.fetchall()]
    conn.close()
    return depts


def get_application(app_id):
    conn = connect_db()
    c = conn.cursor()
    c.execute('''SELECT id, name, vendor, score, need, criticality, installed, disaster_recovery, safety, security, monetary, customer_service, notes, risk_score, last_modified
                 FROM applications WHERE id = ?''', (app_id,))
    row = c.fetchone()
    conn.close()
    return row


def update_application(app_id, fields: Dict[str, object]):
    if not fields:
        return
    allowed = {'name', 'vendor', 'score', 'need', 'criticality', 'installed', 'disaster_recovery', 'safety', 'security', 'monetary', 'customer_service', 'notes', 'risk_score'}
    set_parts = []
    values = []

    # normalize provided fields
    provided = {k.lower(): v for k, v in fields.items()}

    for k, v in fields.items():
        key = k.lower()
        if key in allowed:
            set_parts.append(f"{key} = ?")
            values.append(v)

    # If caller didn't include risk_score, compute it using current values and provided overrides
    if 'risk_score' not in provided:
        conn = connect_db()
        c = conn.cursor()
        c.execute('SELECT score, criticality FROM applications WHERE id = ?', (app_id,))
        cur = c.fetchone()
        conn.close()

        existing_score = _to_int_safe(cur[0] if cur and cur[0] is not None else 0, 0)
        existing_crit = _to_int_safe(cur[1] if cur and cur[1] is not None else 0, 0)

        score_val = _to_int_safe(provided.get('score', existing_score) or existing_score, existing_score)
        crit_val = _to_int_safe(provided.get('criticality', existing_crit) or existing_crit, existing_crit)

        computed_risk = (10 - score_val) * crit_val
        set_parts.append('risk_score = ?')
        values.append(computed_risk)

    # update last_modified
    set_parts.append('last_modified = ?')
    values.append(datetime.utcnow().isoformat())
    values.append(app_id)
    conn = connect_db()
    c = conn.cursor()
    sql = f"UPDATE applications SET {', '.join(set_parts)} WHERE id = ?"
    c.execute(sql, tuple(values))
    conn.commit()
    conn.close()
    return True


def add_application(name: str, vendor: str, factors: Dict[str, int], dept_ids: list, notes: str = ''):
    """Add an application and link it to business units. Returns list of created app ids."""
    conn = connect_db()
    c = conn.cursor()
    # compute risk_score using (10 - score) * criticality
    score_val = _to_int_safe(factors.get('Score', factors.get('score', 0)), 0)
    crit_val = _to_int_safe(factors.get('Criticality', factors.get('criticality', 0)), 0)
    computed_risk = (10 - score_val) * crit_val
    last_mod = datetime.utcnow().isoformat()

    app_ids = []
    for dept_id in dept_ids:
        c.execute('''INSERT INTO applications (name, vendor, score, need, criticality, installed, disaster_recovery, safety, security, monetary, customer_service, notes, risk_score, last_modified)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (
                      name, vendor,
                      _to_int_safe(factors.get('Score', factors.get('score', 0)), 0),
                      _to_int_safe(factors.get('Need', factors.get('need', 0)), 0),
                      _to_int_safe(factors.get('Criticality', factors.get('criticality', 0)), 0),
                      _to_int_safe(factors.get('Installed', factors.get('installed', 0)), 0),
                      _to_int_safe(factors.get('DisasterRecovery', factors.get('disaster_recovery', factors.get('disasterrecovery', 0))), 0),
                      _to_int_safe(factors.get('Safety', factors.get('safety', 0)), 0),
                      _to_int_safe(factors.get('Security', factors.get('security', 0)), 0),
                      _to_int_safe(factors.get('Monetary', factors.get('monetary', 0)), 0),
                      _to_int_safe(factors.get('CustomerService', factors.get('customer_service', 0)), 0),
                      notes,
                      computed_risk,
                      last_mod
                  ))
        app_id = c.lastrowid
        app_ids.append(app_id)
        c.execute('INSERT OR IGNORE INTO application_business_units (app_id, unit_id) VALUES (?, ?)', (app_id, dept_id))

    conn.commit()
    conn.close()
    return app_ids

def purge_database():
    conn = connect_db()
    c = conn.cursor()
    c.execute('DELETE FROM system_integrations')
    c.execute('DELETE FROM application_business_units')
    c.execute('DELETE FROM applications')
    c.execute('DELETE FROM business_units')
    c.execute('DELETE FROM users')
    conn.commit()
    conn.close()

# System Integration Functions
def add_system_integration(parent_app_id, fields: Dict[str, object]):
    """Add a new system integration entry linked to a parent application"""
    if 'name' not in fields or not fields['name']:
        return None
    
    allowed_fields = {'name', 'vendor', 'score', 'need', 'criticality', 'installed', 
                      'disaster_recovery', 'safety', 'security', 'monetary', 'customer_service',
                      'notes', 'risk_score'}
    
    field_keys = []
    field_placeholders = []
    values = [parent_app_id]

    # normalize provided fields for easier lookup
    provided = {k.lower(): v for k, v in fields.items()}

    for k, v in fields.items():
        key = k.lower()
        if key in allowed_fields:
            field_keys.append(key)
            field_placeholders.append('?')
            values.append(v)

    # If caller didn't provide risk_score, compute it using formula: (10 - score) * criticality
    if 'risk_score' not in field_keys:
        score_val = _to_int_safe(provided.get('score', 0) or 0, 0)
        crit_val = _to_int_safe(provided.get('criticality', 0) or 0, 0)
        computed_risk = (10 - score_val) * crit_val
        field_keys.append('risk_score')
        field_placeholders.append('?')
        values.append(computed_risk)

    # Always add last_modified
    field_keys.append('last_modified')
    field_placeholders.append('?')
    values.append(datetime.utcnow().isoformat())
    
    conn = connect_db()
    c = conn.cursor()
    sql = f"INSERT INTO system_integrations (parent_app_id, {', '.join(field_keys)}) VALUES (?, {', '.join(field_placeholders)})"
    c.execute(sql, tuple(values))
    integration_id = c.lastrowid
    conn.commit()
    conn.close()
    
    return integration_id

def get_system_integrations(parent_app_id):
    """Get all system integrations for a parent application"""
    conn = connect_db()
    c = conn.cursor()
    c.execute('''SELECT id, parent_app_id, name, vendor, score, need, criticality, installed,
                 disaster_recovery, safety, security, monetary, customer_service,
                 notes, risk_score, last_modified
                 FROM system_integrations
                 WHERE parent_app_id = ?
                 ORDER BY name''', (parent_app_id,))
    integrations = c.fetchall()
    conn.close()
    return integrations

def get_system_integration(integration_id):
    """Get a specific system integration by ID"""
    conn = connect_db()
    c = conn.cursor()
    c.execute('''SELECT id, parent_app_id, name, vendor, score, need, criticality, installed, 
                 disaster_recovery, safety, security, monetary, customer_service, 
                 notes, risk_score, last_modified
                 FROM system_integrations 
                 WHERE id = ?''', (integration_id,))
    integration = c.fetchone()
    conn.close()
    return integration

def update_system_integration(integration_id, fields: Dict[str, object]):
    """Update a system integration record"""
    if not fields:
        return False
    
    allowed = {'name', 'vendor', 'score', 'need', 'criticality', 'installed', 
               'disaster_recovery', 'safety', 'security', 'monetary', 'customer_service', 
               'notes', 'risk_score'}
    
    set_parts = []
    values = []

    # normalize provided fields
    provided = {k.lower(): v for k, v in fields.items()}

    for k, v in fields.items():
        key = k.lower()
        if key in allowed:
            set_parts.append(f"{key} = ?")
            values.append(v)

    # If caller didn't include risk_score, compute it using current values and provided overrides
    if 'risk_score' not in provided:
        # fetch current row to get existing score/criticality
        conn = connect_db()
        c = conn.cursor()
        c.execute('SELECT score, criticality FROM system_integrations WHERE id = ?', (integration_id,))
        cur = c.fetchone()
        conn.close()

        existing_score = _to_int_safe(cur[0] if cur and cur[0] is not None else 0, 0)
        existing_crit = _to_int_safe(cur[1] if cur and cur[1] is not None else 0, 0)

        # overrides from provided
        score_val = _to_int_safe(provided.get('score', existing_score) or existing_score, existing_score)
        crit_val = _to_int_safe(provided.get('criticality', existing_crit) or existing_crit, existing_crit)

        computed_risk = (10 - score_val) * crit_val
        set_parts.append('risk_score = ?')
        values.append(computed_risk)

    # update last_modified
    set_parts.append('last_modified = ?')
    values.append(datetime.utcnow().isoformat())
    values.append(integration_id)

    conn = connect_db()
    c = conn.cursor()
    sql = f"UPDATE system_integrations SET {', '.join(set_parts)} WHERE id = ?"
    c.execute(sql, tuple(values))
    conn.commit()
    conn.close()
    return True

def delete_system_integration(integration_id):
    """Delete a system integration record"""
    conn = connect_db()
    c = conn.cursor()
    c.execute('DELETE FROM system_integrations WHERE id = ?', (integration_id,))
    conn.commit()
    conn.close()
