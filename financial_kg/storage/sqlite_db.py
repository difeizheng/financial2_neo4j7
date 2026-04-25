"""SQLite task management for upload tracking and version history."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any


DB_PATH = Path(__file__).parent.parent.parent / "data" / "tasks.db"


def init_db(path: str | Path = DB_PATH) -> None:
    """Create tables if they don't exist."""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            file_hash TEXT NOT NULL UNIQUE,
            upload_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT DEFAULT 'pending',
            sheet_count INTEGER DEFAULT 0,
            cell_count INTEGER DEFAULT 0,
            formula_count INTEGER DEFAULT 0,
            graph_path TEXT,
            error TEXT
        );

        CREATE TABLE IF NOT EXISTS versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            upload_id INTEGER NOT NULL,
            version_name TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            graph_path TEXT,
            FOREIGN KEY (upload_id) REFERENCES uploads(id)
        );

        CREATE TABLE IF NOT EXISTS comparisons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            version_a INTEGER NOT NULL,
            version_b INTEGER NOT NULL,
            diff_path TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (version_a) REFERENCES versions(id),
            FOREIGN KEY (version_b) REFERENCES versions(id)
        );
    """)
    conn.commit()
    conn.close()


def add_upload(filename: str, file_hash: str) -> int:
    """Register a new upload, return ID."""
    init_db()
    conn = sqlite3.connect(str(DB_PATH))
    cur = conn.execute(
        "INSERT OR IGNORE INTO uploads (filename, file_hash) VALUES (?, ?)",
        (filename, file_hash),
    )
    conn.commit()
    upload_id = cur.lastrowid or conn.execute(
        "SELECT id FROM uploads WHERE file_hash = ?", (file_hash,)
    ).fetchone()[0]
    conn.close()
    return upload_id


def update_upload(upload_id: int, **kwargs: Any) -> None:
    """Update upload status."""
    init_db()
    conn = sqlite3.connect(str(DB_PATH))
    fields = ", ".join(f"{k} = ?" for k in kwargs)
    values = list(kwargs.values()) + [upload_id]
    conn.execute(f"UPDATE uploads SET {fields} WHERE id = ?", values)
    conn.commit()
    conn.close()


def get_uploads() -> list[dict]:
    """List all uploads."""
    init_db()
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    rows = conn.execute("SELECT * FROM uploads ORDER BY upload_time DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]
