"""Autenticación local y control de acceso basado en roles.

Las contraseñas nunca se guardan en texto plano.  Esta implementación está
pensada para la instalación local; puede sustituirse por SSO corporativo sin
cambiar los permisos de la aplicación.
"""
import hashlib
import hmac
import os
import secrets
import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional, Dict, Any

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_PATH = PROJECT_ROOT / "var" / "security.db"
ROLES = {"ADMINISTRADOR", "SUPERVISOR", "OPERARIO"}


def _connect():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _hash(password: str, salt: Optional[bytes] = None) -> str:
    salt = salt or os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 310000)
    return f"{salt.hex()}${digest.hex()}"


def _verify(password: str, stored: str) -> bool:
    try:
        salt_hex, digest_hex = stored.split("$", 1)
        candidate = _hash(password, bytes.fromhex(salt_hex)).split("$", 1)[1]
        return hmac.compare_digest(candidate, digest_hex)
    except (ValueError, AttributeError):
        return False


def init_db() -> None:
    conn = _connect()
    conn.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
        display_name TEXT NOT NULL, password_hash TEXT NOT NULL, role TEXT NOT NULL,
        active INTEGER NOT NULL DEFAULT 1, must_change_password INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL, updated_at TEXT NOT NULL)
    """)
    conn.execute("""CREATE TABLE IF NOT EXISTS sessions (
        token_hash TEXT PRIMARY KEY, user_id INTEGER NOT NULL, expires_at TEXT NOT NULL,
        created_at TEXT NOT NULL, FOREIGN KEY(user_id) REFERENCES users(id))""")
    
    # Asegurar usuarios iniciales por defecto (admin y operario)
    now = datetime.now(timezone.utc).isoformat()
    admin_user = conn.execute("SELECT id FROM users WHERE username = 'admin'").fetchone()
    if not admin_user:
        conn.execute(
            "INSERT INTO users (username, display_name, password_hash, role, must_change_password, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            ("admin", "Administrador SGA", _hash("admin123"), "ADMINISTRADOR", 1, now, now)
        )
    operario_user = conn.execute("SELECT id FROM users WHERE username = 'operario'").fetchone()
    if not operario_user:
        conn.execute(
            "INSERT INTO users (username, display_name, password_hash, role, must_change_password, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            ("operario", "Operario de Mezclas", _hash("operario123"), "OPERARIO", 1, now, now)
        )
    conn.commit(); conn.close()


def authenticate(username: str, password: str):
    init_db(); conn = _connect()
    row = conn.execute("SELECT * FROM users WHERE username = ? COLLATE NOCASE", (username.strip(),)).fetchone()
    conn.close()
    return dict(row) if row and row["active"] and _verify(password, row["password_hash"]) else None


def create_session(user_id: int) -> str:
    init_db(); token = secrets.token_urlsafe(32)
    now = datetime.now(timezone.utc); expires = now + timedelta(days=7)
    conn = _connect()
    conn.execute("INSERT INTO sessions VALUES (?, ?, ?, ?)", (hashlib.sha256(token.encode()).hexdigest(), user_id, expires.isoformat(), now.isoformat()))
    conn.commit(); conn.close(); return token


def get_user(token: Optional[str]):
    if not token: return None
    init_db(); conn = _connect()
    row = conn.execute("""SELECT u.id, u.username, u.display_name, u.role, u.must_change_password
        FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token_hash=? AND s.expires_at>? AND u.active=1""",
        (hashlib.sha256(token.encode()).hexdigest(), datetime.now(timezone.utc).isoformat())).fetchone()
    conn.close(); return dict(row) if row else None


def delete_session(token: Optional[str]) -> None:
    if token:
        conn = _connect(); conn.execute("DELETE FROM sessions WHERE token_hash=?", (hashlib.sha256(token.encode()).hexdigest(),)); conn.commit(); conn.close()


def update_password(user_id: int, new_password: str) -> bool:
    if len(new_password) < 4:
        raise ValueError("La contraseña debe tener al menos 4 caracteres.")
    init_db()
    conn = _connect()
    now = datetime.now(timezone.utc).isoformat()
    conn.execute(
        "UPDATE users SET password_hash = ?, must_change_password = 0, updated_at = ? WHERE id = ?",
        (_hash(new_password), now, user_id)
    )
    conn.commit()
    conn.close()
    return True


def create_user(username: str, display_name: str, password: str, role: str = "OPERARIO") -> Dict[str, Any]:
    if role not in ROLES:
        role = "OPERARIO"
    if len(password) < 4:
        raise ValueError("La contraseña debe tener al menos 4 caracteres.")
    if len(username.strip()) < 3:
        raise ValueError("El nombre de usuario debe tener al menos 3 caracteres.")
    
    init_db()
    conn = _connect()
    existing = conn.execute("SELECT id FROM users WHERE username = ? COLLATE NOCASE", (username.strip(),)).fetchone()
    if existing:
        conn.close()
        raise ValueError(f"El usuario '{username}' ya está registrado. Por favor inicia sesión o elige otro nombre.")

    now = datetime.now(timezone.utc).isoformat()
    cursor = conn.execute(
        "INSERT INTO users (username, display_name, password_hash, role, must_change_password, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (username.strip(), display_name.strip() or username.strip(), _hash(password), role, 0, now, now)
    )
    user_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return {"id": user_id, "username": username.strip(), "display_name": display_name.strip() or username.strip(), "role": role}


def list_users():
    init_db(); conn = _connect(); rows = conn.execute("SELECT id, username, display_name, role, active, must_change_password, created_at FROM users ORDER BY username").fetchall(); conn.close(); return [dict(r) for r in rows]
