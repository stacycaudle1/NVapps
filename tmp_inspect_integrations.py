import sqlite3
import pprint
from datetime import datetime
import database

DB = database.DB_NAME

def print_pragma(table):
    conn = sqlite3.connect(DB)
    c = conn.cursor()
    c.execute(f"PRAGMA table_info({table})")
    rows = c.fetchall()
    print(f"PRAGMA {table} columns (idx,name,type,notnull,dflt,pk):")
    pprint.pprint(rows)
    conn.close()


def main():
    print_pragma('system_integrations')
    rows = database.get_system_integrations(1)
    print('\nReturned rows from get_system_integrations(1):')
    for r in rows:
        print('\nFULL TUPLE:')
        pprint.pprint(r)
        # show index mapping for GUI
        mapping = {
            'id': r[0],
            'parent_app_id': r[1],
            'name': r[2],
            'vendor': r[3],
            'score': r[4],
            'need': r[5],
            'criticality': r[6],
            'installed': r[7],
            'disaster_recovery': r[8],
            'safety': r[9],
            'security': r[10],
            'monetary': r[11],
            'customer_service': r[12],
            'notes': r[13],
            'risk_score': r[14],
            'last_modified': r[15] if len(r) > 15 else None
        }
        print('\nMAPPED by expected SELECT order:')
        pprint.pprint(mapping)

if __name__ == '__main__':
    main()
