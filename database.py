import sqlite3
from datetime import datetime
from dataclasses import dataclass, field
from typing import Dict, Optional
import math

DB_NAME = 'business_apps.db'

ATTRIBUTES = [
    "Stability",
    "Need",
    "Criticality",
    "Installed",
    "DisasterRecovery",
    "Safety",
    "Security",
    "Monetary",
    "CustomerService",
]

DEFAULT_WEIGHTS = {
    "Stability":         0.10,
    "Need":              0.10,
    "Criticality":       0.20,
    "Installed":         0.05,
    "DisasterRecovery":  0.15,
    "Safety":            0.15,
    "Security":          0.15,
    "Monetary":          0.10,
    "CustomerService":   0.05,
}

@dataclass(frozen=True)
class DRScoreResult:
    total: float
    priority: str
    breakdown: Dict[str, float]
    ratings: Dict[str, int]
    weights: Dict[str, float]

def _normalize_weights(weights: Dict[str, float]) -> Dict[str, float]:
    unknown = set(weights).difference(ATTRIBUTES)
    if unknown:
        raise ValueError(f"Unknown weight keys: {sorted(unknown)}")
    w = {k: float(weights.get(k, 0.0)) for k in ATTRIBUTES}
    total = sum(w.values())
    if total <= 0:
        raise ValueError("Sum of weights must be > 0")
    return {k: v / total for k, v in w.items()}

def _validate_ratings(ratings: Dict[str, int]) -> Dict[str, int]:
    missing = [k for k in ATTRIBUTES if k not in ratings]
    if missing:
        raise ValueError(f"Missing ratings for: {missing}")
    unknown = set(ratings).difference(ATTRIBUTES)
    if unknown:
        raise ValueError(f"Unknown rating keys: {sorted(unknown)}")
    clean = {}
    for k, v in ratings.items():
        if not isinstance(v, (int, float)) or math.isnan(float(v)):
            raise ValueError(f"Rating for {k} must be a number 1..10")
        iv = int(round(float(v)))
        if not (1 <= iv <= 10):
            raise ValueError(f"Rating for {k} must be in 1..10 (got {v})")
        clean[k] = iv
    return clean

def dr_priority_band(total_score: float) -> str:
    if total_score >= 8.5:
        return "Critical"
    if total_score >= 7.0:
        return "High"
    if total_score >= 5.5:
        return "Medium"
    return "Low"

def score_application(
    ratings: Dict[str, int],
    weights: Optional[Dict[str, float]] = None,
) -> DRScoreResult:
    clean_ratings = _validate_ratings(ratings)
    normalized_weights = _normalize_weights(weights or DEFAULT_WEIGHTS)
    breakdown: Dict[str, float] = {}
    total = 0.0
    for k in ATTRIBUTES:
        contrib = clean_ratings[k] * normalized_weights[k]
        breakdown[k] = round(contrib, 4)
        total += contrib
    total = round(total, 2)
    return DRScoreResult(
        total=total,
        priority=dr_priority_band(total),
        breakdown=breakdown,
        ratings=clean_ratings,
        weights=normalized_weights,
    )

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
            stability INTEGER DEFAULT 0,
            need INTEGER DEFAULT 0,
            criticality INTEGER DEFAULT 0,
            installed INTEGER DEFAULT 0,
            disasterrecovery INTEGER DEFAULT 0,
            safety INTEGER DEFAULT 0,
            security INTEGER DEFAULT 0,
            monetary INTEGER DEFAULT 0,
            customerservice INTEGER DEFAULT 0,
            notes TEXT,
            user_id INTEGER,
            risk_score REAL,
            last_modified TEXT
        )''')
    else:
        # table exists: ensure factor columns, vendor, and last_modified exist
        c.execute("PRAGMA table_info(applications)")
        columns = [col[1].lower() for col in c.fetchall()]
        factor_columns = ['stability', 'need', 'criticality', 'installed', 'disasterrecovery', 'safety', 'security', 'monetary', 'customerservice']
        for col in factor_columns:
            if col not in columns:
                c.execute(f'ALTER TABLE applications ADD COLUMN {col} INTEGER DEFAULT 0')
        if 'vendor' not in columns:
            c.execute('ALTER TABLE applications ADD COLUMN vendor TEXT')
        if 'last_modified' not in columns:
            c.execute("ALTER TABLE applications ADD COLUMN last_modified TEXT")
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS departments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS application_departments (
        app_id INTEGER,
        dept_id INTEGER,
        PRIMARY KEY (app_id, dept_id),
        FOREIGN KEY(app_id) REFERENCES applications(id),
        FOREIGN KEY(dept_id) REFERENCES departments(id)
    )''')
    conn.commit()
    conn.close()

def calculate_business_risk(app_row):
    # app_row order: id, name, vendor, Stability, Need, Criticality, DisasterRecovery, Safety, Security, Monetary, CustomerService
    ratings = {
        "Stability": app_row[3],
        "Need": app_row[4],
        "Criticality": app_row[5],
    "Installed": app_row[6],
    "DisasterRecovery": app_row[7],
    "Safety": app_row[8],
    "Security": app_row[9],
    "Monetary": app_row[10],
    "CustomerService": app_row[11],
    }
    result = score_application(ratings)
    return result.total, result.priority

def link_app_to_departments(app_id, dept_ids):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    for dept_id in dept_ids:
        c.execute('INSERT OR IGNORE INTO application_departments (app_id, dept_id) VALUES (?, ?)', (app_id, dept_id))
    conn.commit()
    conn.close()

def get_app_departments(app_id):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('SELECT d.name FROM departments d JOIN application_departments ad ON d.id = ad.dept_id WHERE ad.app_id = ?', (app_id,))
    depts = [row[0] for row in c.fetchall()]
    conn.close()
    return depts


def get_application(app_id):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''SELECT id, name, vendor, stability, need, criticality, installed, disasterrecovery, safety, security, monetary, customerservice, notes, risk_score, last_modified
                 FROM applications WHERE id = ?''', (app_id,))
    row = c.fetchone()
    conn.close()
    return row


def update_application(app_id, fields: Dict[str, object]):
    if not fields:
        return
    allowed = {'name', 'vendor', 'stability', 'need', 'criticality', 'installed', 'disasterrecovery', 'safety', 'security', 'monetary', 'customerservice', 'notes', 'risk_score'}
    set_parts = []
    values = []
    for k, v in fields.items():
        key = k.lower()
        if key in allowed:
            set_parts.append(f"{key} = ?")
            values.append(v)
    # update last_modified
    set_parts.append('last_modified = ?')
    values.append(datetime.utcnow().isoformat())
    values.append(app_id)
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    sql = f"UPDATE applications SET {', '.join(set_parts)} WHERE id = ?"
    c.execute(sql, tuple(values))
    conn.commit()
    conn.close()

def purge_database():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('DELETE FROM application_departments')
    c.execute('DELETE FROM applications')
    c.execute('DELETE FROM departments')
    c.execute('DELETE FROM users')
    conn.commit()
    conn.close()
