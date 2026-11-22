
import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / "student_performance.db"


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT,
            course TEXT,
            attendance REAL,
            study_hours REAL,
            previous_marks REAL,
            assignment_score REAL,
            sleep_hours REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS predictions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_name TEXT,
            attendance REAL,
            study_hours REAL,
            previous_marks REAL,
            assignment_score REAL,
            sleep_hours REAL,
            performance TEXT,
            confidence REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.commit()
    conn.close()


def row_to_dict(row):
    return dict(row) if row else None
