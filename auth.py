import base64
import hashlib
import hmac
import json
import os
import secrets
import sqlite3
import time
from pathlib import Path

from fastapi import Depends, Header, HTTPException


BASE_DIR = Path(__file__).resolve().parent
USERS_DB = Path(os.getenv("USERS_DB_PATH", BASE_DIR / "users.db"))
ADMIN_CONFIG = BASE_DIR / "admin.local.json"
SESSION_SECRET_FILE = BASE_DIR / ".session_secret"
USER_ROLES = {"vendedor", "supervisor", "administrador"}
PASSWORD_ITERATIONS = 600_000


def fixed_admin_credentials():
    if not ADMIN_CONFIG.exists():
        raise ValueError(
            "Arquivo admin.local.json não encontrado. Configure o administrador local."
        )
    try:
        data = json.loads(ADMIN_CONFIG.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError("Não foi possível ler admin.local.json.") from exc

    username = normalize_username(data.get("username", "admin"))
    password = str(data.get("password", ""))
    if not username or len(password) < 6:
        raise ValueError("Administrador local inválido em admin.local.json.")
    return username, password


def database_connection():
    connection = sqlite3.connect(USERS_DB)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def normalize_username(value):
    return str(value or "").strip().lower()


def hash_password(password, salt=None):
    password = str(password or "")
    if len(password) < 6:
        raise ValueError("A senha deve ter pelo menos 6 caracteres.")
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PASSWORD_ITERATIONS,
    )
    return f"pbkdf2_sha256${PASSWORD_ITERATIONS}${salt.hex()}${digest.hex()}"


def verify_password(password, encoded):
    try:
        algorithm, iterations, salt_hex, digest_hex = encoded.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        candidate = hashlib.pbkdf2_hmac(
            "sha256",
            str(password or "").encode("utf-8"),
            bytes.fromhex(salt_hex),
            int(iterations),
        ).hex()
        return hmac.compare_digest(candidate, digest_hex)
    except (AttributeError, TypeError, ValueError):
        return False


def initialize_users_db():
    with database_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE COLLATE NOCASE,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL CHECK (role IN ('vendedor', 'supervisor', 'administrador')),
                active INTEGER NOT NULL DEFAULT 1,
                token_version INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )

        admin_username, admin_password = fixed_admin_credentials()
        admin = connection.execute(
            "SELECT * FROM users WHERE username = ?",
            (admin_username,),
        ).fetchone()
        if admin:
            admin_changed = (
                not verify_password(admin_password, admin["password_hash"])
                or admin["role"] != "administrador"
                or not admin["active"]
            )
            if admin_changed:
                connection.execute(
                    """
                    UPDATE users
                    SET password_hash = ?, role = 'administrador', active = 1,
                        token_version = token_version + 1, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (hash_password(admin_password), admin["id"]),
                )
        else:
            connection.execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, 'administrador')",
                (admin_username, hash_password(admin_password)),
            )

        bootstrap_users = [
            ("vendedor", os.getenv("SELLER_PASSWORD", ""), "vendedor"),
            ("supervisor", os.getenv("SUPERVISOR_PASSWORD", ""), "supervisor"),
        ]
        for username, password, role in bootstrap_users:
            if not username or not password:
                continue
            exists = connection.execute(
                "SELECT 1 FROM users WHERE username = ?",
                (username,),
            ).fetchone()
            if not exists:
                connection.execute(
                    "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                    (username, hash_password(password), role),
                )


def public_user(row):
    fixed_admin_username, _ = fixed_admin_credentials()
    return {
        "id": row["id"],
        "username": row["username"],
        "role": row["role"],
        "active": bool(row["active"]),
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
        "fixed": row["username"].lower() == fixed_admin_username,
    }


def authenticate(username, password):
    username = normalize_username(username)
    with database_connection() as connection:
        row = connection.execute(
            "SELECT * FROM users WHERE username = ? AND active = 1",
            (username,),
        ).fetchone()
    if not row or not verify_password(password, row["password_hash"]):
        return None
    return {
        "id": row["id"],
        "username": row["username"],
        "role": row["role"],
        "token_version": row["token_version"],
    }


def token_user(user_id):
    with database_connection() as connection:
        row = connection.execute(
            "SELECT * FROM users WHERE id = ? AND active = 1",
            (user_id,),
        ).fetchone()
    if not row:
        return None
    return {
        "id": row["id"],
        "username": row["username"],
        "role": row["role"],
        "token_version": row["token_version"],
    }


def session_secret():
    if SESSION_SECRET_FILE.exists():
        secret = SESSION_SECRET_FILE.read_text(encoding="ascii").strip()
        if secret:
            return secret

    secret = secrets.token_urlsafe(48)
    SESSION_SECRET_FILE.write_text(secret, encoding="ascii")
    return secret


def create_token(user):
    payload = {
        "uid": user["id"],
        "username": user["username"],
        "role": user["role"],
        "ver": user["token_version"],
        "exp": int(time.time()) + 8 * 60 * 60,
    }
    encoded = base64.urlsafe_b64encode(
        json.dumps(payload, separators=(",", ":")).encode("utf-8")
    ).decode("ascii")
    signature = hmac.new(
        session_secret().encode("utf-8"),
        encoded.encode("ascii"),
        hashlib.sha256,
    ).hexdigest()
    return f"{encoded}.{signature}"


def require_user(authorization: str = Header(default="")):
    try:
        scheme, token = authorization.split(" ", 1)
        if scheme.lower() != "bearer":
            raise ValueError
        encoded, signature = token.strip().split(".", 1)
        expected = hmac.new(
            session_secret().encode("utf-8"),
            encoded.encode("ascii"),
            hashlib.sha256,
        ).hexdigest()
        if not hmac.compare_digest(signature, expected):
            raise ValueError
        payload = json.loads(base64.urlsafe_b64decode(encoded).decode("utf-8"))
        if payload.get("exp", 0) < time.time():
            raise ValueError

        with database_connection() as connection:
            row = connection.execute(
                "SELECT * FROM users WHERE id = ? AND active = 1",
                (payload.get("uid"),),
            ).fetchone()
        if not row or row["token_version"] != payload.get("ver"):
            raise ValueError
        return {
            "id": row["id"],
            "username": row["username"],
            "role": row["role"],
        }
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=401, detail="Sessão inválida ou expirada.") from exc


def require_admin(current_user: dict = Depends(require_user)):
    if current_user["role"] != "administrador":
        raise HTTPException(status_code=403, detail="Acesso exclusivo do administrador.")
    return current_user


def list_users():
    with database_connection() as connection:
        rows = connection.execute(
            "SELECT * FROM users ORDER BY username COLLATE NOCASE"
        ).fetchall()
    return [public_user(row) for row in rows]


def create_user(username, password, role):
    username = normalize_username(username)
    if len(username) < 3:
        raise ValueError("O usuário deve ter pelo menos 3 caracteres.")
    if role not in USER_ROLES:
        raise ValueError("Nível de permissão inválido.")
    try:
        with database_connection() as connection:
            cursor = connection.execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                (username, hash_password(password), role),
            )
            row = connection.execute(
                "SELECT * FROM users WHERE id = ?",
                (cursor.lastrowid,),
            ).fetchone()
    except sqlite3.IntegrityError as exc:
        raise ValueError("Já existe um usuário com esse nome.") from exc
    return public_user(row)


def update_user(user_id, username=None, role=None, active=None, password=None):
    with database_connection() as connection:
        current = connection.execute(
            "SELECT * FROM users WHERE id = ?",
            (user_id,),
        ).fetchone()
        if not current:
            raise ValueError("Usuário não encontrado.")
        fixed_admin_username, _ = fixed_admin_credentials()
        if current["username"].lower() == fixed_admin_username:
            raise ValueError("O administrador fixo não pode ser alterado.")

        next_username = normalize_username(username) if username is not None else current["username"]
        next_role = role if role is not None else current["role"]
        next_active = int(active) if active is not None else current["active"]
        if len(next_username) < 3:
            raise ValueError("O usuário deve ter pelo menos 3 caracteres.")
        if next_role not in USER_ROLES:
            raise ValueError("Nível de permissão inválido.")

        removing_admin = current["role"] == "administrador" and (
            next_role != "administrador" or not next_active
        )
        if removing_admin:
            admin_count = connection.execute(
                "SELECT COUNT(*) FROM users WHERE role = 'administrador' AND active = 1"
            ).fetchone()[0]
            if admin_count <= 1:
                raise ValueError("O sistema precisa manter ao menos um administrador ativo.")

        password_hash = current["password_hash"]
        if password:
            password_hash = hash_password(password)
        token_version = current["token_version"] + 1
        try:
            connection.execute(
                """
                UPDATE users
                SET username = ?, password_hash = ?, role = ?, active = ?,
                    token_version = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                """,
                (next_username, password_hash, next_role, next_active, token_version, user_id),
            )
        except sqlite3.IntegrityError as exc:
            raise ValueError("Já existe um usuário com esse nome.") from exc
        row = connection.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    return public_user(row)


def delete_user(user_id, current_user_id):
    if user_id == current_user_id:
        raise ValueError("Você não pode excluir o próprio usuário.")
    with database_connection() as connection:
        row = connection.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            raise ValueError("Usuário não encontrado.")
        fixed_admin_username, _ = fixed_admin_credentials()
        if row["username"].lower() == fixed_admin_username:
            raise ValueError("O administrador fixo não pode ser excluído.")
        if row["role"] == "administrador" and row["active"]:
            admin_count = connection.execute(
                "SELECT COUNT(*) FROM users WHERE role = 'administrador' AND active = 1"
            ).fetchone()[0]
            if admin_count <= 1:
                raise ValueError("O sistema precisa manter ao menos um administrador ativo.")
        connection.execute("DELETE FROM users WHERE id = ?", (user_id,))


initialize_users_db()
