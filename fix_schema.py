import sqlite3

def fix_database_schema():
    conn = sqlite3.connect('business_apps.db')
    cursor = conn.cursor()
    
    # Drop duplicate tables
    cursor.execute("DROP TABLE IF EXISTS app_departments")
    
    # Ensure system_integrations has the correct structure
    cursor.execute("DROP TABLE IF EXISTS system_integrations")
    cursor.execute('''
        CREATE TABLE system_integrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            parent_app_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            vendor TEXT,
            score INTEGER,
            need INTEGER,
            criticality INTEGER,
            installed INTEGER,
            disaster_recovery INTEGER,
            safety INTEGER,
            security INTEGER,
            monetary INTEGER,
            customer_service INTEGER,
            notes TEXT,
            risk_score REAL,
            last_modified TEXT,
            FOREIGN KEY(parent_app_id) REFERENCES applications(id)
        )
    ''')
    
    conn.commit()
    conn.close()

if __name__ == "__main__":
    fix_database_schema()