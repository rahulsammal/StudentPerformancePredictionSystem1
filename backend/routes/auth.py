"""Administrator authentication routes for EduPredict.

Uses only Python's standard library for salted password hashing and signed tokens.
The token secret must be configured with EDUPREDICT_SECRET_KEY outside demos.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import time
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from pydantic import BaseModel, Field

router = APIRouter(prefix="/api/auth", tags=["Admin Authentication"])
bearer_scheme = HTTPBearer(auto_error=False)

BACKEND_DIR = Path(__file__).resolve().parents[1]
DATABASE_PATH = Path(
    os.getenv("EDUPREDICT_DB_PATH", str(BACKEND_DIR / "database" / "student_performance.db"))
)
DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)

# Set a long random value in your environment before running FastAPI.
# The fallback is intentionally for a local college demo only.
TOKEN_SECRET = os.getenv("EDUPREDICT_SECRET_KEY", "CHANGE-THIS-DEMO-SECRET-BEFORE-DEPLOYMENT").encode("utf-8")
TOKEN_LIFETIME_SECONDS = 60 * 60 * 12
PBKDF2_ITERATIONS = 310_000


def db_connection() -> sqlite3.Connection:
    connection = sqlite3.connect(DATABASE_PATH, timeout=15)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


def init_auth_db() -> None:
    with db_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS admins (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT NOT NULL UNIQUE COLLATE NOCASE,
                password_hash TEXT NOT NULL,
                password_salt TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        connection.commit()


class RegisterRequest(BaseModel):
    name: str = Field(min_length=2, max_length=80)
    email: str = Field(min_length=5, max_length=254)
    password: str = Field(min_length=8, max_length=128)


class LoginRequest(BaseModel):
    email: str = Field(min_length=5, max_length=254)
    password: str = Field(min_length=1, max_length=128)


def normalise_email(value: str) -> str:
    email = value.strip().lower()
    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
        raise HTTPException(status_code=400, detail="Please enter a valid email address.")
    return email


def hash_password(password: str, salt_hex: str | None = None) -> tuple[str, str]:
    salt = bytes.fromhex(salt_hex) if salt_hex else secrets.token_bytes(16)
    derived = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return derived.hex(), salt.hex()


def verify_password(password: str, stored_hash: str, salt_hex: str) -> bool:
    calculated, _ = hash_password(password, salt_hex)
    return hmac.compare_digest(calculated, stored_hash)


def b64url_encode(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def b64url_decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def create_access_token(admin_id: int, email: str) -> str:
    now = int(time.time())
    header = {"alg": "HS256", "typ": "JWT"}
    payload = {"sub": str(admin_id), "email": email, "iat": now, "exp": now + TOKEN_LIFETIME_SECONDS}
    head = b64url_encode(json.dumps(header, separators=(",", ":")).encode("utf-8"))
    body = b64url_encode(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
    signing_input = f"{head}.{body}".encode("ascii")
    signature = hmac.new(TOKEN_SECRET, signing_input, hashlib.sha256).digest()
    return f"{head}.{body}.{b64url_encode(signature)}"


def decode_access_token(token: str) -> dict[str, Any]:
    try:
        head, body, signature = token.split(".")
        signing_input = f"{head}.{body}".encode("ascii")
        expected = hmac.new(TOKEN_SECRET, signing_input, hashlib.sha256).digest()
        supplied = b64url_decode(signature)
        if not hmac.compare_digest(expected, supplied):
            raise ValueError("Invalid signature")
        header = json.loads(b64url_decode(head))
        payload = json.loads(b64url_decode(body))
        if header.get("alg") != "HS256" or int(payload.get("exp", 0)) < int(time.time()):
            raise ValueError("Expired or unsupported token")
        int(payload["sub"])
        return payload
    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Your admin session is invalid or expired. Please sign in again.",
            headers={"WWW-Authenticate": "Bearer"},
        ) from exc


def public_admin(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "name": row["name"],
        "email": row["email"],
        "role": "Administrator",
        "created_at": row["created_at"],
    }


async def get_current_admin(
    credentials: HTTPAuthorizationCredentials | None = Depends(bearer_scheme),
) -> dict[str, Any]:
    if credentials is None or credentials.scheme.lower() != "bearer":
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Administrator sign-in is required.",
            headers={"WWW-Authenticate": "Bearer"},
        )
    payload = decode_access_token(credentials.credentials)
    with db_connection() as connection:
        row = connection.execute(
            "SELECT id, name, email, created_at FROM admins WHERE id = ?",
            (int(payload["sub"]),),
        ).fetchone()
    if row is None:
        raise HTTPException(status_code=401, detail="Admin account no longer exists. Please sign in again.")
    return public_admin(row)


@router.post("/register", status_code=status.HTTP_201_CREATED)
def register_admin(request: RegisterRequest):
    name = request.name.strip()
    email = normalise_email(request.email)
    if len(name) < 2:
        raise HTTPException(status_code=400, detail="Please enter your full name.")
    password_hash, salt = hash_password(request.password)
    try:
        with db_connection() as connection:
            cursor = connection.execute(
                "INSERT INTO admins (name, email, password_hash, password_salt) VALUES (?, ?, ?, ?)",
                (name, email, password_hash, salt),
            )
            connection.commit()
            admin_id = int(cursor.lastrowid)
            row = connection.execute(
                "SELECT id, name, email, created_at FROM admins WHERE id = ?", (admin_id,)
            ).fetchone()
    except sqlite3.IntegrityError as exc:
        raise HTTPException(status_code=409, detail="An administrator account with this email already exists.") from exc

    admin = public_admin(row)
    return {"access_token": create_access_token(admin_id, email), "token_type": "bearer", "admin": admin}


@router.post("/login")
def login_admin(request: LoginRequest):
    email = normalise_email(request.email)
    with db_connection() as connection:
        row = connection.execute(
            "SELECT id, name, email, password_hash, password_salt, created_at FROM admins WHERE email = ?",
            (email,),
        ).fetchone()
    if row is None or not verify_password(request.password, row["password_hash"], row["password_salt"]):
        raise HTTPException(status_code=401, detail="Incorrect email or password.")

    admin = public_admin(row)
    return {"access_token": create_access_token(row["id"], row["email"]), "token_type": "bearer", "admin": admin}


@router.get("/me")
async def read_current_admin(admin: dict[str, Any] = Depends(get_current_admin)):
    return admin
