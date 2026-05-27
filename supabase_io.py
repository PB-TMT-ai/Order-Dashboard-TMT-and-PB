"""Persistence of uploaded Excel binaries to Supabase Storage.

All functions degrade gracefully when Supabase is not configured: the app still
runs, it just won't persist files between sessions.
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


def is_configured() -> bool:
    return bool(SUPABASE_URL and SUPABASE_KEY)


def _get_client():
    global _client
    if _client is not None:
        return _client
    if not is_configured():
        return None
    try:
        from supabase import create_client
        _client = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception:
        _client = None
    return _client


def upload_bytes(name: str, data: bytes) -> bool:
    client = _get_client()
    if client is None:
        return False
    try:
        client.storage.from_(SUPABASE_BUCKET).upload(
            path=name, file=data,
            file_options={"content-type": "application/octet-stream", "upsert": "true"},
        )
        return True
    except Exception:
        return False


def download_bytes(name: str) -> bytes | None:
    client = _get_client()
    if client is None:
        return None
    try:
        return client.storage.from_(SUPABASE_BUCKET).download(name)
    except Exception:
        return None


def save_order_file(data: bytes) -> bool:
    return upload_bytes(ORDER_FILE, data)


def save_be_file(data: bytes) -> bool:
    return upload_bytes(BE_FILE, data)


def load_order_file() -> bytes | None:
    return download_bytes(ORDER_FILE)


def load_be_file() -> bytes | None:
    return download_bytes(BE_FILE)
