import sqlite3
import os

def check_schema():
    conn = sqlite3.connect('business_apps.db')
    cursor = conn.cursor()
    
    # Get all tables
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [x[0] for x in cursor.fetchall()]
    print("Tables in database:", tables)
    
    # For each table, get its schema
    for table in tables:
        cursor.execute(f"PRAGMA table_info({table})")
        columns = cursor.fetchall()
        print(f"\nSchema for {table}:")
        for col in columns:
            print(f"  {col[1]} ({col[2]})")

if __name__ == "__main__":
    check_schema()