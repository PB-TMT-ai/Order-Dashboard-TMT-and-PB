"""Persistence of uploaded files to Cloudflare R2 (S3-compatible object storage).

R2 has no small per-file cap (multipart up to 5 TB) and no egress fees, so the
full order/BE workbooks can be stored and shared across users.

All functions degrade gracefully when storage is not configured: the app still
runs, it just won't persist files between sessions. Errors are captured in
`last_error()` so the UI can surface why a connection failed.
"""
from __future__ import annotations

import json
import os

from dotenv import load_dotenv

load_dotenv()

ORDER_FILE = "latest_order.xlsx"
BE_FILE = "latest_be.xlsx"
BE_META_FILE = "latest_be_meta.json"

_client = None
_last_error: str | None = None


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


def endpoint() -> str:
    """Full S3 endpoint. Use R2_ENDPOINT, or build it from R2_ACCOUNT_ID."""
    ep = _cfg("R2_ENDPOINT")
    if ep:
        return ep.rstrip("/")
    acct = _cfg("R2_ACCOUNT_ID")
    return f"https://{acct}.r2.cloudflarestorage.com" if acct else ""


def access_key() -> str:
    return _cfg("R2_ACCESS_KEY_ID")


def secret_key() -> str:
    return _cfg("R2_SECRET_ACCESS_KEY")


def bucket() -> str:
    return _cfg("R2_BUCKET", "jsw-dashboard")


def last_error() -> str | None:
    return _last_error


def is_configured() -> bool:
    return bool(endpoint() and access_key() and secret_key() and bucket())


def _err(exc: BaseException) -> str:
    """Readable message from a botocore ClientError (or any exception)."""
    resp = getattr(exc, "response", None)
    if isinstance(resp, dict):
        err = resp.get("Error", {})
        code = err.get("Code", "")
        msg = err.get("Message", "")
        detail = " ".join(p for p in (code, msg) if p).strip()
        if detail:
            return detail
    return str(exc) or exc.__class__.__name__


def _get_client():
    global _client, _last_error
    if _client is not None:
        return _client
    if not is_configured():
        _last_error = "R2 endpoint / keys / bucket not set"
        return None
    try:
        import boto3
        from botocore.config import Config
        _client = boto3.client(
            "s3",
            endpoint_url=endpoint(),
            aws_access_key_id=access_key(),
            aws_secret_access_key=secret_key(),
            region_name="auto",
            config=Config(signature_version="s3v4", retries={"max_attempts": 3}),
        )
    except Exception as exc:  # noqa: BLE001
        _last_error = f"client init failed: {_err(exc)}"
        _client = None
    return _client


def check_connection() -> tuple[bool, str]:
    """Verify creds + bucket are reachable. Returns (ok, human-readable message)."""
    global _last_error
    if not is_configured():
        return False, "Not configured — set R2 endpoint, keys and bucket."
    client = _get_client()
    if client is None:
        return False, _last_error or "Could not create client."
    try:
        client.head_bucket(Bucket=bucket())
        _last_error = None
        return True, f"Connected · R2 bucket '{bucket()}'"
    except Exception as exc:  # noqa: BLE001
        _last_error = _err(exc)
        return False, f"R2 bucket '{bucket()}' not reachable — {_last_error}"


def upload_bytes(name: str, data: bytes) -> bool:
    global _last_error
    client = _get_client()
    if client is None:
        return False
    try:
        client.put_object(Bucket=bucket(), Key=name, Body=data,
                          ContentType="application/octet-stream")
        _last_error = None
        return True
    except Exception as exc:  # noqa: BLE001
        _last_error = _err(exc)
        return False


def download_bytes(name: str) -> bytes | None:
    global _last_error
    client = _get_client()
    if client is None:
        return None
    try:
        obj = client.get_object(Bucket=bucket(), Key=name)
        return obj["Body"].read()
    except Exception as exc:  # noqa: BLE001
        # Missing object is expected (nothing saved yet) — not an error worth surfacing
        code = getattr(exc, "response", {}).get("Error", {}).get("Code", "")
        if code not in ("NoSuchKey", "404", "NoSuchBucket"):
            _last_error = _err(exc)
        return None


def save_order_file(data: bytes) -> bool:
    return upload_bytes(ORDER_FILE, data)


def save_be_file(data: bytes) -> bool:
    return upload_bytes(BE_FILE, data)


def save_be_meta(meta: dict) -> bool:
    return upload_bytes(BE_META_FILE, json.dumps(meta).encode("utf-8"))


def load_order_file() -> bytes | None:
    return download_bytes(ORDER_FILE)


def load_be_file() -> bytes | None:
    return download_bytes(BE_FILE)


def load_be_meta() -> dict | None:
    raw = download_bytes(BE_META_FILE)
    if not raw:
        return None
    try:
        return json.loads(raw)
    except Exception:  # noqa: BLE001
        return None
