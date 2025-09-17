import sqlite3

DB_PATH = 'business_apps.db'

def migrate():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS category_notes (
        app_id INTEGER NOT NULL,
        category TEXT NOT NULL,
        notes TEXT,
        PRIMARY KEY (app_id, category)
    )''')
    conn.commit()
    conn.close()

if __name__ == '__main__':
    migrate()
    print('category_notes table created (if not already present).')
