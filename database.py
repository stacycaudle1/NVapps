import sqlite3

def get_system_integrations(parent_app_id=None):
    """Get system integrations, optionally filtered by parent application ID."""
    conn = connect_db()
    c = conn.cursor()
    try:
        if parent_app_id is not None:
            c.execute('''SELECT i.*, a.name as parent_app_name
                        FROM system_integrations i
                        LEFT JOIN applications a ON i.parent_app_id = a.id
                        WHERE i.parent_app_id = ?
                        ORDER BY i.name''', (parent_app_id,))
        else:
            c.execute('''SELECT i.*, a.name as parent_app_name
                        FROM system_integrations i
                        LEFT JOIN applications a ON i.parent_app_id = a.id
                        ORDER BY i.name''')
        return c.fetchall()
    finally:
        conn.close()

def purge_database():
    """Delete all data from the database while preserving the schema."""
    conn = connect_db()
    c = conn.cursor()
    try:
        # Get list of all tables
        c.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in c.fetchall() if row[0] != 'sqlite_sequence']
        
        # Disable foreign key constraints temporarily
        c.execute('PRAGMA foreign_keys = OFF')
        
        # Delete all data from each table
        for table in tables:
            try:
                c.execute(f'DELETE FROM {table}')
            except Exception as e:
                print(f"DEBUG: Error purging table {table}: {e}")
                
        # Reset all auto-increment counters
        for table in tables:
            try:
                c.execute(f"DELETE FROM sqlite_sequence WHERE name='{table}'")
            except Exception as e:
                print(f"DEBUG: Error resetting sequence for {table}: {e}")
        
        # Re-enable foreign key constraints
        c.execute('PRAGMA foreign_keys = ON')
        
        conn.commit()
        print("DEBUG: Database purged successfully")
    except Exception as e:
        print(f"DEBUG: Error during database purge: {e}")
        conn.rollback()
        raise
    finally:
        conn.close()

from datetime import datetime
from dataclasses import dataclass, field
from typing import Dict, Optional
import math
import time

DB_NAME = 'business_apps.db'

def connect_db():
    """Get a database connection with proper settings."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def initialize_database():
    """Initialize or upgrade the database schema."""
    max_retries = 3
    last_error = None
    
    for attempt in range(max_retries):
        try:
            conn = connect_db()
            c = conn.cursor()
            
            # Enable foreign key support
            c.execute('PRAGMA foreign_keys = ON')
            
            # Check if business_units table needs last_modified column
            c.execute("PRAGMA table_info(business_units)")
            columns = [col[1] for col in c.fetchall()]
            
            if not columns:  # Table doesn't exist
                print("DEBUG: Creating business_units table")
                c.execute("""CREATE TABLE business_units (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    last_modified TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")
            elif 'last_modified' not in columns:  # Need to add the column
                print("DEBUG: Adding last_modified column to business_units table")
                # Create new table with desired schema
                c.execute("""CREATE TABLE business_units_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    last_modified TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )""")
                # Copy existing data
                c.execute("INSERT INTO business_units_new (id, name) SELECT id, name FROM business_units")
                # Drop old table and rename new one
                c.execute("DROP TABLE business_units")
                c.execute("ALTER TABLE business_units_new RENAME TO business_units")
            
            # Create other required tables
            c.execute("""CREATE TABLE IF NOT EXISTS application_business_units (
                app_id INTEGER,
                unit_id INTEGER,
                PRIMARY KEY (app_id, unit_id),
                FOREIGN KEY(app_id) REFERENCES applications(id),
                FOREIGN KEY(unit_id) REFERENCES business_units(id)
            )""")
            
            c.execute("""CREATE TABLE IF NOT EXISTS applications (
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
            )""")
            
            c.execute("""CREATE TABLE IF NOT EXISTS system_integrations (
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
            )""")
            
            conn.commit()
            conn.close()
            print("DEBUG: Database schema initialized")
            return
            
        except sqlite3.OperationalError as e:
            last_error = e
            if 'database is locked' in str(e) and attempt < max_retries - 1:
                print(f"DEBUG: Database locked, retrying in {(attempt + 1) * 0.5} seconds")
                time.sleep(0.5 * (attempt + 1))
                continue
            else:
                raise
    
    if last_error:
        raise last_error

def add_application(name, vendor, factors, dept_ids, notes=None):
    """Add a new application with the given attributes."""
    conn = connect_db()
    c = conn.cursor()
    try:
        now = datetime.utcnow().isoformat()
        c.execute('''INSERT INTO applications 
                    (name, vendor, score, need, criticality, installed, disaster_recovery,
                     safety, security, monetary, customer_service, notes, last_modified)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                 (name, vendor,
                  factors.get('score', 0),
                  factors.get('need', 0),
                  factors.get('criticality', 0),
                  factors.get('installed', 0),
                  factors.get('disaster_recovery', 0),
                  factors.get('safety', 0),
                  factors.get('security', 0),
                  factors.get('monetary', 0),
                  factors.get('customer_service', 0),
                  notes, now))
        app_id = c.lastrowid
        
        # Link to departments
        if dept_ids:
            for dept_id in dept_ids:
                c.execute('INSERT OR IGNORE INTO application_business_units (app_id, unit_id) VALUES (?, ?)',
                         (app_id, dept_id))
        
        conn.commit()
        return app_id
    finally:
        conn.close()

def add_system_integration(parent_app_id, fields):
    """Add a new system integration linked to the given parent application."""
    conn = connect_db()
    c = conn.cursor()
    try:
        now = datetime.utcnow().isoformat()
        c.execute('''INSERT INTO system_integrations 
                    (parent_app_id, name, vendor, score, need, criticality, installed,
                     disaster_recovery, safety, security, monetary, customer_service,
                     notes, risk_score, last_modified)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                 (parent_app_id,
                  fields.get('name'),
                  fields.get('vendor'),
                  fields.get('score', 0),
                  fields.get('need', 0),
                  fields.get('criticality', 0),
                  fields.get('installed', 0),
                  fields.get('disaster_recovery', 0),
                  fields.get('safety', 0),
                  fields.get('security', 0),
                  fields.get('monetary', 0),
                  fields.get('customer_service', 0),
                  fields.get('notes'),
                  fields.get('risk_score'),
                  now))
        int_id = c.lastrowid
        conn.commit()
        return int_id
    finally:
        conn.close()

def get_system_integration(integration_id):
    """Return a single system integration row by id (sqlite3.Row)."""
    conn = connect_db()
    c = conn.cursor()
    try:
        c.execute("SELECT * FROM system_integrations WHERE id = ?", (integration_id,))
        return c.fetchone()
    finally:
        conn.close()

def update_system_integration(integration_id, fields):
    """Update fields on a system integration and bump last_modified.

    Returns True if a row was updated, False otherwise.
    """
    if not fields:
        fields = {}

    allowed = {
        'name', 'vendor', 'score', 'need', 'criticality', 'installed',
        'disaster_recovery', 'safety', 'security', 'monetary', 'customer_service',
        'notes', 'risk_score'
    }

    set_parts = []
    values = []

    # normalize keys to lowercase
    normalized = {str(k).lower(): v for k, v in fields.items()}

    for key, val in normalized.items():
        if key in allowed:
            set_parts.append(f"{key} = ?")
            values.append(val)

    # Always update last_modified
    set_parts.append('last_modified = ?')
    values.append(datetime.utcnow().isoformat())
    values.append(integration_id)

    sql = f"UPDATE system_integrations SET {', '.join(set_parts)} WHERE id = ?"

    conn = connect_db()
    c = conn.cursor()
    try:
        c.execute(sql, values)
        conn.commit()
        return c.rowcount > 0
    finally:
        conn.close()

def delete_system_integration(integration_id):
    """Delete a system integration by id. Returns True if a row was deleted."""
    conn = connect_db()
    c = conn.cursor()
    try:
        c.execute("DELETE FROM system_integrations WHERE id = ?", (integration_id,))
        conn.commit()
        return c.rowcount > 0
    finally:
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
    """Get business unit names for an application"""
    conn = connect_db()
    c = conn.cursor()
    c.execute('''
        SELECT bu.name
        FROM business_units bu
        JOIN application_business_units abu ON bu.id = abu.unit_id
        WHERE abu.app_id = ?
        ORDER BY bu.name''', (app_id,))
    business_units = [row[0] for row in c.fetchall()]
    conn.close()
    return business_units

def get_application(app_id):
    conn = connect_db()
    c = conn.cursor()
    c.execute('''SELECT id, name, vendor, score, need, criticality, installed,
                 disaster_recovery, safety, security, monetary, customer_service,
                 notes, risk_score, last_modified
                 FROM applications WHERE id = ?''', (app_id,))
    row = c.fetchone()
    conn.close()
    return row

def update_application(app_id, fields: Dict[str, object]):
    if not fields:
        return
    allowed = {'name', 'vendor', 'score', 'need', 'criticality', 'installed',
               'disaster_recovery', 'safety', 'security', 'monetary',
               'customer_service', 'notes', 'risk_score'}
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
    
    if set_parts:
        conn = connect_db()
        c = conn.cursor()
        # Always update last_modified
        set_parts.append('last_modified = ?')
        values.append(datetime.utcnow().isoformat())
        values.append(app_id)
        
        c.execute(f'''UPDATE applications SET {', '.join(set_parts)}
                     WHERE id = ?''', values)
        conn.commit()
        conn.close()

def _to_int_safe(val, default=0):
    """Convert a value to int, with fallback to default."""
    if val is None:
        return default
    try:
        return int(float(str(val)))
    except (ValueError, TypeError):
        return default

def dr_priority_band(risk_score):
    """Map risk score to priority band (high/med/low)."""
    try:
        score = float(risk_score)
    except (ValueError, TypeError):
        return None
        
    if score >= 70:
        return 'High'
    if score >= 50:
        return 'Med'
    if score >= 1:
        return 'Low'
    return None