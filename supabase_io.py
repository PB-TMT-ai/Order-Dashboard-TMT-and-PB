"""Persistence of uploaded Excel binaries to Supabase Storage.

All functions degrade gracefully when Supabase is not configured: the app still
runs, it just won't persist files between sessions. Errors are captured in
`last_error()` so the UI can surface why a connection failed.
"""
from __future__ import annotations

import os

from dotenv import load_dotenv

load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
SUPABASE_BUCKET = os.getenv("SUPABASE_BUCKET", "jsw-dashboard")

ORDER_FILE = "latest_order.xlsx"
BE_FILE = "latest_be.xlsx"

_client = None
_last_error: str | None = None


def last_error() -> str | None:
    return _last_error


def is_configured() -> bool:
    return bool(SUPABASE_URL and SUPABASE_KEY)


def _get_client():
    global _client, _last_error
    if _client is not None:
        return _client
    if not is_configured():
        _last_error = "SUPABASE_URL / SUPABASE_KEY not set"
        return None
    try:
        from supabase import create_client
        _client = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception as exc:  # noqa: BLE001
        _last_error = f"create_client failed: {exc}"
        _client = None
    return _client


def check_connection() -> tuple[bool, str]:
    """Verify creds + bucket are reachable. Returns (ok, human-readable message)."""
    global _last_error
    if not is_configured():
        return False, "Not configured — set SUPABASE_URL and SUPABASE_KEY."
    client = _get_client()
    if client is None:
        return False, _last_error or "Could not create client."
    try:
        client.storage.from_(SUPABASE_BUCKET).list()
        _last_error = None
        return True, f"Connected · bucket '{SUPABASE_BUCKET}'"
    except Exception as exc:  # noqa: BLE001
        _last_error = str(exc)
        return False, f"Bucket '{SUPABASE_BUCKET}' not reachable: {exc}"


def upload_bytes(name: str, data: bytes) -> bool:
    global _last_error
    client = _get_client()
    if client is None:
        return False
    try:
        client.storage.from_(SUPABASE_BUCKET).upload(
            path=name, file=data,
            file_options={"content-type": "application/octet-stream", "upsert": "true"},
        )
        _last_error = None
        return True
    except Exception as exc:  # noqa: BLE001
        _last_error = str(exc)
        return False


def download_bytes(name: str) -> bytes | None:
    global _last_error
    client = _get_client()
    if client is None:
        return None
    try:
        return client.storage.from_(SUPABASE_BUCKET).download(name)
    except Exception as exc:  # noqa: BLE001
        _last_error = str(exc)
        return None


def save_order_file(data: bytes) -> bool:
    return upload_bytes(ORDER_FILE, data)


def save_be_file(data: bytes) -> bool:
    return upload_bytes(BE_FILE, data)


def load_order_file() -> bytes | None:
    return download_bytes(ORDER_FILE)


def load_be_file() -> bytes | None:
    return download_bytes(BE_FILE)
