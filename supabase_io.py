"""Persistence of uploaded Excel binaries to Supabase Storage.

All functions degrade gracefully when Supabase is not configured: the app still
runs, it just won't persist files between sessions. Errors are captured in
`last_error()` so the UI can surface why a connection failed.
"""
from __future__ import annotations

import os

from dotenv import load_dotenv

load_dotenv()


def _cfg(key: str, default: str = "") -> str:
    """Read config from OS env (local/.env) or st.secrets (Streamlit Cloud)."""
    val = os.getenv(key)
    if val:
        return val
    try:
        import streamlit as st
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:  # noqa: BLE001 — no secrets file / not in a Streamlit run
        pass
    return default


def supabase_url() -> str:
    return _cfg("SUPABASE_URL")


def supabase_key() -> str:
    return _cfg("SUPABASE_KEY")


def bucket() -> str:
    return _cfg("SUPABASE_BUCKET", "jsw-dashboard")


ORDER_FILE = "latest_order.xlsx"
BE_FILE = "latest_be.xlsx"

_client = None
_last_error: str | None = None


def last_error() -> str | None:
    return _last_error


def _http_detail(exc: BaseException) -> str:
    """Recover the real HTTP status/body from an exception chain.

    storage3 2.30.0 has a bug where an API error response is re-raised as
    'dict object has no attribute text', masking the real cause. The original
    httpx error (with .response) is still in the __context__/__cause__ chain.
    """
    seen: set[int] = set()
    cur: BaseException | None = exc
    while cur is not None and id(cur) not in seen:
        seen.add(id(cur))
        resp = getattr(cur, "response", None)
        if resp is not None:
            try:
                body = resp.text
            except Exception:  # noqa: BLE001
                body = ""
            status = getattr(resp, "status_code", "?")
            return f"HTTP {status}: {body[:300]}".strip()
        cur = cur.__cause__ or cur.__context__
    return str(exc) or exc.__class__.__name__


def is_configured() -> bool:
    return bool(supabase_url() and supabase_key())


def _get_client():
    global _client, _last_error
    if _client is not None:
        return _client
    if not is_configured():
        _last_error = "SUPABASE_URL / SUPABASE_KEY not set"
        return None
    try:
        from supabase import create_client
        _client = create_client(supabase_url(), supabase_key())
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
        client.storage.from_(bucket()).list()
        _last_error = None
        return True, f"Connected · bucket '{bucket()}'"
    except Exception as exc:  # noqa: BLE001
        detail = _http_detail(exc)
        _last_error = detail
        return False, f"Bucket '{bucket()}' check failed — {detail}"


def upload_bytes(name: str, data: bytes) -> bool:
    global _last_error
    client = _get_client()
    if client is None:
        return False
    try:
        client.storage.from_(bucket()).upload(
            path=name, file=data,
            file_options={"content-type": "application/octet-stream", "upsert": "true"},
        )
        _last_error = None
        return True
    except Exception as exc:  # noqa: BLE001
        _last_error = _http_detail(exc)
        return False


def download_bytes(name: str) -> bytes | None:
    global _last_error
    client = _get_client()
    if client is None:
        return None
    try:
        return client.storage.from_(bucket()).download(name)
    except Exception as exc:  # noqa: BLE001
        _last_error = _http_detail(exc)
        return None


def save_order_file(data: bytes) -> bool:
    return upload_bytes(ORDER_FILE, data)


def save_be_file(data: bytes) -> bool:
    return upload_bytes(BE_FILE, data)


def load_order_file() -> bytes | None:
    return download_bytes(ORDER_FILE)


def load_be_file() -> bytes | None:
    return download_bytes(BE_FILE)
