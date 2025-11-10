# tests/test_database.py
import sys
from pathlib import Path

# Add parent directory to path so we can import database module
sys.path.insert(0, str(Path(__file__).parent.parent))

import database

def test_connect_db():
    conn = database.connect_db()
    assert conn is not None
    conn.close()

def test_ensure_category():
    cat_id = database.ensure_category("Test Category")
    assert cat_id is not None