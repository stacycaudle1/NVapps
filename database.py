import sqlite3
from typing import Dict

__all__ = [
    'DB_NAME',
    'connect_db',
    'initialize_database',
    'purge_database',
    'add_application',
    'add_system_integration',
    'get_system_integration',
    'get_system_integrations',
    'update_system_integration',
    'delete_system_integration',
    'calculate_business_risk',
    'dr_priority_band',
    'link_app_to_departments',
    'get_app_departments',
    'get_application',
    'update_application',
    'generate_smoke_test_data',
    'get_categories',
    'ensure_category',
    'get_category_name',
    'update_category',
    'delete_category',
    'get_app_categories',
    'link_app_to_categories',
    'set_app_categories',
    'clear_app_categories',
    'touch_application',
]

def get_system_integrations(parent_app_id=None):
    """Get system integrations, optionally filtered by parent application ID."""
    conn = connect_db()
    c = conn.cursor()
    try:
        if parent_app_id is not None:
            c.execute('''SELECT i.*, a.division as parent_app_name
                        FROM system_integrations i
                        LEFT JOIN applications a ON i.parent_app_id = a.id
                        WHERE i.parent_app_id = ?
                        ORDER BY i.name''', (parent_app_id,))
        else:
            c.execute('''SELECT i.*, a.division as parent_app_name
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
            
            # Categories table (new)
            c.execute("""CREATE TABLE IF NOT EXISTS categories (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                last_modified TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )""")

            # Create other required tables
            c.execute("""CREATE TABLE IF NOT EXISTS application_business_units (
                app_id INTEGER,
                unit_id INTEGER,
                PRIMARY KEY (app_id, unit_id),
                FOREIGN KEY(app_id) REFERENCES applications(id),
                FOREIGN KEY(unit_id) REFERENCES business_units(id)
            )""")
            # Many-to-many mapping for application categories
            c.execute("""CREATE TABLE IF NOT EXISTS application_categories (
                app_id INTEGER,
                category_id INTEGER,
                PRIMARY KEY (app_id, category_id),
                FOREIGN KEY(app_id) REFERENCES applications(id),
                FOREIGN KEY(category_id) REFERENCES categories(id)
            )""")
            
            # Applications table: prefer 'division' column (formerly 'name')
            c.execute("""CREATE TABLE IF NOT EXISTS applications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                division TEXT NOT NULL,
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

            # Migrate applications table to ensure 'division' and 'category_id' exist
            c.execute("PRAGMA table_info(applications)")
            app_cols = [col[1] for col in c.fetchall()]
            # If legacy schema used 'name', add 'division' and backfill from 'name'
            if 'division' not in app_cols:
                try:
                    c.execute('ALTER TABLE applications ADD COLUMN division TEXT')
                    # Backfill division from name if present
                    if 'name' in app_cols:
                        try:
                            c.execute("UPDATE applications SET division = name WHERE division IS NULL OR TRIM(division) = ''")
                        except Exception:
                            pass
                except Exception as e:
                    print(f"DEBUG: Could not add division to applications: {e}")
            if 'category_id' not in app_cols:
                try:
                    c.execute('ALTER TABLE applications ADD COLUMN category_id INTEGER')
                except Exception as e:
                    # If ALTER fails for some reason, log and continue
                    print(f"DEBUG: Could not add category_id to applications: {e}")
            
            # Optional: create indexes for faster lookups
            try:
                c.execute('CREATE INDEX IF NOT EXISTS idx_applications_category ON applications(category_id)')
            except Exception:
                pass
            try:
                c.execute('CREATE INDEX IF NOT EXISTS idx_applications_division ON applications(division)')
            except Exception:
                pass

            # One-time migration: populate application_categories from existing category_id values
            try:
                c.execute('''
                    INSERT OR IGNORE INTO application_categories (app_id, category_id)
                    SELECT id, category_id FROM applications WHERE category_id IS NOT NULL
                ''')
            except Exception as e:
                print(f"DEBUG: Migration of application_categories failed or not needed: {e}")
            
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

            # Many-to-many mapping for integration categories (added for per-row integration category linkage)
            c.execute("""CREATE TABLE IF NOT EXISTS integration_categories (
                integration_id INTEGER NOT NULL,
                category_id INTEGER NOT NULL,
                PRIMARY KEY (integration_id, category_id),
                FOREIGN KEY (integration_id) REFERENCES system_integrations(id) ON DELETE CASCADE,
                FOREIGN KEY (category_id) REFERENCES categories(id) ON DELETE CASCADE
            )""")
            try:
                c.execute('CREATE INDEX IF NOT EXISTS idx_integration_categories_integration ON integration_categories(integration_id)')
            except Exception:
                pass
            try:
                c.execute('CREATE INDEX IF NOT EXISTS idx_integration_categories_category ON integration_categories(category_id)')
            except Exception:
                pass
            
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

def _now_ts():
    """Return a normalized local timestamp string (YYYY-MM-DD HH:MM:SS)."""
    try:
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        # Fallback to UTC iso without microseconds
        return datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')


def add_application(division, vendor, factors, dept_ids, notes=None):
    """Add a new application (Division) with the given attributes."""
    conn = connect_db()
    c = conn.cursor()
    try:
        # Detect if legacy 'name' column still exists (and may be NOT NULL)
        c.execute("PRAGMA table_info(applications)")
        app_cols = [row[1] for row in c.fetchall()]
        has_legacy_name = 'name' in app_cols
        # Use local wall-clock time for display consistency in UI
        now = _now_ts()
        if has_legacy_name:
            # Insert both division and legacy name for compatibility (legacy 'name' may be NOT NULL)
            c.execute('''INSERT INTO applications 
                        (division, name, vendor, score, need, criticality, installed, disaster_recovery,
                         safety, security, monetary, customer_service, notes, last_modified)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                     (division, division, vendor,
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
        else:
            c.execute('''INSERT INTO applications 
                        (division, vendor, score, need, criticality, installed, disaster_recovery,
                         safety, security, monetary, customer_service, notes, last_modified)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                     (division, vendor,
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
        now = _now_ts()
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
    values.append(_now_ts())
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

def generate_smoke_test_data():
    """Populate database with a small set of demo apps, business units, and integrations."""
    initialize_database()
    import random
    from datetime import datetime, timezone

    conn = connect_db()
    c = conn.cursor()
    try:
        # Seed business units
        bu_names = ["Finance", "HR", "IT"]
        for n in bu_names:
            c.execute("INSERT OR IGNORE INTO business_units (name) VALUES (?)", (n,))

        # Create demo applications
        app_defs = [
            ("TestApp1", "VendorA", "IT"),
            ("TestApp2", "VendorB", "HR"),
            ("TestApp3", "VendorC", "Finance"),
        ]

        app_ids = []
        for name, vendor, bu in app_defs:
            factors: Dict[str, int] = {
                'score': random.randint(1, 10),
                'need': random.randint(1, 10),
                'criticality': random.randint(1, 10),
                'installed': random.randint(1, 10),
                'disaster_recovery': random.randint(1, 10),
                'safety': random.randint(1, 10),
                'security': random.randint(1, 10),
                'monetary': random.randint(1, 10),
                'customer_service': random.randint(1, 10),
            }
            # Find BU id
            c.execute("SELECT id FROM business_units WHERE name = ?", (bu,))
            row = c.fetchone()
            bu_id = row[0] if row else None
            app_id = add_application(name, vendor, factors, [bu_id] if bu_id else [], notes=f"Notes for {name}")
            app_ids.append(app_id)

        # Create a couple of integrations per app
        for app_id in app_ids:
            for j in range(2):
                fields = {
                    'name': f"Int{j+1} for App {app_id}",
                    'vendor': random.choice(["AWS", "Azure", "GCP", "Internal"]),
                    'score': random.randint(1, 10),
                    'need': random.randint(1, 10),
                    'criticality': random.randint(1, 10),
                    'installed': random.randint(1, 10),
                    'disaster_recovery': random.randint(1, 10),
                    'safety': random.randint(1, 10),
                    'security': random.randint(1, 10),
                    'monetary': random.randint(1, 10),
                    'customer_service': random.randint(1, 10),
                    'notes': f"Auto generated integration {j+1}",
                    'risk_score': random.uniform(1.0, 100.0),
                }
                add_system_integration(app_id, fields)
        conn.commit()
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
    # Return division as 'name' for compatibility with existing GUI code
    c.execute('''SELECT id, division AS name, vendor, score, need, criticality, installed,
                 disaster_recovery, safety, security, monetary, customer_service,
                 notes, risk_score, last_modified, category_id
                 FROM applications WHERE id = ?''', (app_id,))
    row = c.fetchone()
    conn.close()
    return row

def update_application(app_id, fields: Dict[str, object]):
    if not fields:
        return
    # Accept both 'division' and legacy 'name' and map to division
    allowed = {'division', 'name', 'vendor', 'score', 'need', 'criticality', 'installed',
               'disaster_recovery', 'safety', 'security', 'monetary',
               'customer_service', 'notes', 'risk_score', 'category_id'}
    set_parts = []
    values = []
    
    # normalize provided fields
    provided = {k.lower(): v for k, v in fields.items()}
    # Map legacy 'name' to 'division'
    if 'name' in provided and 'division' not in provided:
        provided['division'] = provided['name']
    
    for k, v in provided.items():
        key = k.lower()
        if key in allowed:
            # Always write to 'division' for both 'division' and legacy 'name'
            column = 'division' if key == 'name' else key
            set_parts.append(f"{column} = ?")
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
        values.append(_now_ts())
        values.append(app_id)
        
        c.execute(f'''UPDATE applications SET {', '.join(set_parts)}
                     WHERE id = ?''', values)
        conn.commit()
        conn.close()

def get_categories():
    """Return list of all categories as (id, name)."""
    conn = connect_db()
    c = conn.cursor()
    try:
        c.execute('SELECT id, name FROM categories ORDER BY name')
        return c.fetchall()
    finally:
        conn.close()

def ensure_category(name: str):
    """Find or create a category by name. Returns its id or None if name is empty."""
    if not name:
        return None
    conn = connect_db()
    c = conn.cursor()
    try:
        nm = str(name).strip()
        c.execute('SELECT id FROM categories WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))', (nm,))
        row = c.fetchone()
        if row:
            return row[0]
        c.execute('INSERT INTO categories (name, last_modified) VALUES (?, CURRENT_TIMESTAMP)', (nm,))
        conn.commit()
        return c.lastrowid
    finally:
        conn.close()

def get_category_name(category_id: int):
    """Return category name for given id, or None."""
    if not category_id:
        return None
    conn = connect_db()
    c = conn.cursor()
    try:
        c.execute('SELECT name FROM categories WHERE id = ?', (category_id,))
        r = c.fetchone()
        return r[0] if r else None
    finally:
        conn.close()

def update_category(category_id: int, new_name: str):
    """Rename a category."""
    conn = connect_db()
    c = conn.cursor()
    try:
        c.execute('UPDATE categories SET name = ?, last_modified = CURRENT_TIMESTAMP WHERE id = ?', (new_name, category_id))
        conn.commit()
        return c.rowcount > 0
    finally:
        conn.close()

def delete_category(category_id: int):
    """Delete a category and clear references from applications."""
    conn = connect_db()
    c = conn.cursor()
    try:
        # Clear references from applications (legacy single link)
        try:
            c.execute('UPDATE applications SET category_id = NULL WHERE category_id = ?', (category_id,))
        except Exception:
            pass
        # Clear references from application_categories (many-to-many)
        try:
            c.execute('DELETE FROM application_categories WHERE category_id = ?', (category_id,))
        except Exception:
            pass
        c.execute('DELETE FROM categories WHERE id = ?', (category_id,))
        conn.commit()
        return c.rowcount > 0
    finally:
        conn.close()

def get_app_categories(app_id: int):
    """Return list of category names for an application, ordered by name."""
    conn = connect_db()
    c = conn.cursor()
    try:
        print(f"DEBUG: Getting categories for app_id {app_id}")  # Debug log
        c.execute('''
            SELECT c.name, ac.app_id, ac.category_id 
            FROM categories c
            JOIN application_categories ac ON c.id = ac.category_id
            WHERE ac.app_id = ?
            ORDER BY c.name
        ''', (app_id,))
        rows = c.fetchall()
        print(f"DEBUG: Found {len(rows)} categories for app_id {app_id}")  # Debug log
        for row in rows:
            print(f"DEBUG: Category data for app {app_id}: name={row[0]}, app_id={row[1]}, cat_id={row[2]}")
        return [row[0] for row in rows]
    finally:
        conn.close()

def link_app_to_categories(app_id: int, category_ids):
    """Link an application to multiple categories (insert-or-ignore)."""
    if not category_ids:
        return
    conn = connect_db()
    c = conn.cursor()
    try:
        for cid in category_ids:
            if cid:
                c.execute('INSERT OR IGNORE INTO application_categories (app_id, category_id) VALUES (?, ?)', (app_id, cid))
        conn.commit()
    finally:
        conn.close()

def clear_app_categories(app_id: int):
    conn = connect_db()
    c = conn.cursor()
    try:
        c.execute('DELETE FROM application_categories WHERE app_id = ?', (app_id,))
        conn.commit()
    finally:
        conn.close()

def touch_application(app_id: int):
    """Update the last_modified timestamp for an application."""
    if not app_id:
        return
    conn = connect_db()
    c = conn.cursor()
    try:
        c.execute('UPDATE applications SET last_modified = ? WHERE id = ?', (_now_ts(), app_id))
        conn.commit()
    finally:
        conn.close()

def set_app_categories(app_id: int, category_ids):
    """Replace an application's category links with the provided set."""
    max_retries = 5
    retry_delay = 0.1  # Start with 100ms delay
    last_error = None
    
    for attempt in range(max_retries):
        conn = None
        try:
            conn = connect_db()
            c = conn.cursor()
            
            # First get current categories for debug logging
            c.execute('SELECT DISTINCT c.name FROM categories c JOIN application_categories ac ON c.id = ac.category_id WHERE ac.app_id = ?', (app_id,))
            current_cats = [r[0] for r in c.fetchall()]
            print(f"DEBUG: Current categories for app {app_id}: {current_cats}")
            
            # Clear all existing category links within a transaction
            c.execute('BEGIN IMMEDIATE')  # Request immediate transaction to prevent locks
            c.execute('DELETE FROM application_categories WHERE app_id = ?', (app_id,))
            
            # Add new category links
            if category_ids:
                print(f"DEBUG: Setting new category IDs for app {app_id}: {category_ids}")
                for cid in category_ids:
                    if cid:
                        c.execute('INSERT OR IGNORE INTO application_categories (app_id, category_id) VALUES (?, ?)', (app_id, cid))
            
            # Commit all changes at once
            conn.commit()
            
            # Log the new categories
            c.execute('SELECT DISTINCT c.name FROM categories c JOIN application_categories ac ON c.id = ac.category_id WHERE ac.app_id = ?', (app_id,))
            new_cats = [r[0] for r in c.fetchall()]
            print(f"DEBUG: Successfully set new categories for app {app_id}: {new_cats}")
            
            # Success - return from the function
            return
            
        except sqlite3.OperationalError as e:
            last_error = e
            if 'database is locked' in str(e) and attempt < max_retries - 1:
                print(f"DEBUG: Database locked while setting categories for app {app_id}, retry {attempt + 1}")
                if conn:
                    try:
                        conn.rollback()
                    except Exception:
                        pass
                time.sleep(retry_delay)
                retry_delay *= 2  # Exponential backoff
                continue
            raise
        finally:
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass
    
    if last_error:
        raise last_error

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