"""System report: verify last_modified timestamps for applications and integrations.
Format preserved (two sections):
Applications: ID | Division | Last Modified
Integrations: ID | Parent Division | Name | Last Modified
"""
import database
database.initialize_database()

def fetch_app_last_modified():
    conn = database.connect_db()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT a.id, a.division, a.last_modified
        FROM applications a
        ORDER BY a.id
        """
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def fetch_int_last_modified():
    conn = database.connect_db()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT i.id, a.division AS parent_division, i.name, i.last_modified
        FROM system_integrations i
        LEFT JOIN applications a ON i.parent_app_id = a.id
        ORDER BY i.id
        """
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def print_report():
    apps = fetch_app_last_modified()
    ints = fetch_int_last_modified()
    print("Applications:", flush=True)
    print("ID | Division | Last Modified", flush=True)
    for r in apps:
        print(f"{r[0]} | {r[1]} | {r[2]}", flush=True)
    print("", flush=True)
    print("Integrations:", flush=True)
    print("ID | Parent Division | Name | Last Modified", flush=True)
    for r in ints:
        print(f"{r[0]} | {r[1] or ''} | {r[2]} | {r[3]}", flush=True)

if __name__ == "__main__":
    print_report()
