import json
import os
import sqlite3
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
AUDIT_DB = Path(os.getenv("AUDIT_DB_PATH", BASE_DIR / "audit_log.db"))


def audit_connection():
    connection = sqlite3.connect(AUDIT_DB)
    connection.row_factory = sqlite3.Row
    return connection


def initialize_audit_db():
    with audit_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS audit_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                user_id INTEGER,
                username TEXT,
                role TEXT,
                action TEXT NOT NULL,
                status TEXT NOT NULL,
                details TEXT NOT NULL DEFAULT '{}'
            )
            """
        )


def log_audit_event(user, action, status="success", details=None):
    try:
        initialize_audit_db()
        safe_details = details if isinstance(details, dict) else {}
        with audit_connection() as connection:
            connection.execute(
                """
                INSERT INTO audit_events (user_id, username, role, action, status, details)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    user.get("id") if user else None,
                    user.get("username") if user else None,
                    user.get("role") if user else None,
                    str(action or "unknown"),
                    str(status or "success"),
                    json.dumps(safe_details, ensure_ascii=False, default=str),
                ),
            )
    except Exception:
        return


def list_audit_events(limit=200):
    initialize_audit_db()
    limit = max(1, min(int(limit or 200), 1000))
    with audit_connection() as connection:
        rows = connection.execute(
            """
            SELECT id, created_at, user_id, username, role, action, status, details
            FROM audit_events
            ORDER BY id DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()

    events = []
    for row in rows:
        try:
            details = json.loads(row["details"] or "{}")
        except json.JSONDecodeError:
            details = {}
        events.append({
            "id": row["id"],
            "createdAt": row["created_at"],
            "userId": row["user_id"],
            "username": row["username"],
            "role": row["role"],
            "action": row["action"],
            "status": row["status"],
            "details": details,
        })
    return events


def get_audit_event(event_id):
    initialize_audit_db()
    with audit_connection() as connection:
        row = connection.execute(
            """
            SELECT id, created_at, user_id, username, role, action, status, details
            FROM audit_events
            WHERE id = ?
            """,
            (event_id,),
        ).fetchone()
    if not row:
        return None
    try:
        details = json.loads(row["details"] or "{}")
    except json.JSONDecodeError:
        details = {}
    return {
        "id": row["id"],
        "createdAt": row["created_at"],
        "userId": row["user_id"],
        "username": row["username"],
        "role": row["role"],
        "action": row["action"],
        "status": row["status"],
        "details": details,
    }


initialize_audit_db()
