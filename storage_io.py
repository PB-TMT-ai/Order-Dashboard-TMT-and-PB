"""Persistence of uploaded files to Cloudflare R2 (S3-compatible object storage).

R2 has no small per-file cap (multipart up to 5 TB) and no egress fees, so the
full order/BE workbooks can be stored and shared across users.

All functions degrade gracefully when storage is not configured: the app still
runs, it just won't persist files between sessions. Errors are captured in
`last_error()` so the UI can surface why a connection failed.
"""
from __future__ import annotations

import io
import json
import os

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# Processed (Parquet) — small, fast, what the dashboard actually consumes.
PROCESSED_ORDERS_FILE = "processed_orders.parquet"
PROCESSED_INVOICES_FILE = "processed_invoices.parquet"
# Legacy single-BE slot (still read for backward compat).
BE_FILE = "latest_be.xlsx"
BE_META_FILE = "latest_be_meta.json"
# Versioned BE slots — one per (month, week). Up to 4 weeks per month.
BE_VERSION_PREFIX = "be_versions/"

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


def _df_to_parquet_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_parquet(buf, engine="pyarrow", compression="snappy", index=False)
    return buf.getvalue()


def _parquet_bytes_to_df(b: bytes) -> pd.DataFrame:
    return pd.read_parquet(io.BytesIO(b), engine="pyarrow")


def _clean_invoices_for_parquet(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce invoice columns to clean, parquet-safe types.

    Source columns can be a mix (e.g. invoice numbers that are sometimes
    numeric, sometimes alphanumeric) which Arrow refuses to infer.
    """
    out = pd.DataFrame(index=df.index)
    if "Order ID" in df.columns:
        out["Order ID"] = df["Order ID"].astype(str)
    if "Invoice date" in df.columns:
        out["Invoice date"] = pd.to_datetime(df["Invoice date"], errors="coerce")
    if "Invoiced qty" in df.columns:
        out["Invoiced qty"] = pd.to_numeric(df["Invoiced qty"], errors="coerce").fillna(0.0)
    if "Invoice number" in df.columns:
        out["Invoice number"] = df["Invoice number"].astype(str)
    return out


def save_processed(orders_df: pd.DataFrame, invoices_df: pd.DataFrame | None) -> bool:
    """Persist the processed order frame + raw invoice lines as compact Parquet.

    Replaces the older approach of saving the full raw Excel — those files
    were too large to load on Streamlit Cloud's RAM. Parquet keeps only what
    the dashboard needs, in a columnar/compressed format. Errors during
    conversion or upload are captured (don't crash the page); the dashboard
    still works for the current session.
    """
    global _last_error
    try:
        orders_bytes = _df_to_parquet_bytes(orders_df)
    except Exception as exc:  # noqa: BLE001
        _last_error = f"orders parquet conversion failed: {exc}"
        return False
    ok = upload_bytes(PROCESSED_ORDERS_FILE, orders_bytes)
    if invoices_df is not None and len(invoices_df):
        try:
            inv_bytes = _df_to_parquet_bytes(_clean_invoices_for_parquet(invoices_df))
        except Exception as exc:  # noqa: BLE001
            _last_error = f"invoices parquet conversion failed: {exc}"
            return False
        ok = upload_bytes(PROCESSED_INVOICES_FILE, inv_bytes) and ok
    return ok


def load_processed() -> tuple[pd.DataFrame | None, pd.DataFrame | None]:
    orders_b = download_bytes(PROCESSED_ORDERS_FILE)
    if not orders_b:
        return None, None
    orders = _parquet_bytes_to_df(orders_b)
    inv_b = download_bytes(PROCESSED_INVOICES_FILE)
    invoices = _parquet_bytes_to_df(inv_b) if inv_b else None
    return orders, invoices


def save_be_file(data: bytes) -> bool:
    return upload_bytes(BE_FILE, data)


def save_be_meta(meta: dict) -> bool:
    return upload_bytes(BE_META_FILE, json.dumps(meta).encode("utf-8"))


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


# ─── Versioned BE storage (multi-week per month) ─────────────────────────────
def _be_version_key(month: str, week: str) -> str:
    """Object key for a (month, week) slot. month=YYYY-MM, week=W1..W4."""
    return f"{BE_VERSION_PREFIX}{month}_{week}.xlsx"


def _be_version_meta_key(month: str, week: str) -> str:
    return f"{BE_VERSION_PREFIX}{month}_{week}_meta.json"


def be_version_exists(month: str, week: str) -> bool:
    """True if a (month, week) BE slot is already stored."""
    global _last_error
    client = _get_client()
    if client is None:
        return False
    try:
        client.head_object(Bucket=bucket(), Key=_be_version_key(month, week))
        return True
    except Exception:  # noqa: BLE001
        return False


def save_be_version(month: str, week: str, data_bytes: bytes,
                    meta: dict) -> bool:
    """Persist a versioned BE file + its meta JSON. Returns True on success."""
    ok = upload_bytes(_be_version_key(month, week), data_bytes)
    if ok:
        ok = upload_bytes(_be_version_meta_key(month, week),
                          json.dumps(meta).encode("utf-8"))
    return ok


def load_be_version(month: str, week: str) -> tuple[bytes | None, dict | None]:
    """Return (file_bytes, meta) for a (month, week) slot, or (None, None)."""
    data_b = download_bytes(_be_version_key(month, week))
    if not data_b:
        return None, None
    raw = download_bytes(_be_version_meta_key(month, week))
    try:
        meta = json.loads(raw) if raw else None
    except Exception:  # noqa: BLE001
        meta = None
    return data_b, meta


def list_be_versions() -> list[dict]:
    """List every stored (month, week) BE slot.

    Returns a list of {"month": "YYYY-MM", "week": "W1", "sheet": str,
    "uploaded": iso, "key": object_key}, sorted by (month desc, week desc).
    """
    global _last_error
    client = _get_client()
    if client is None:
        return []
    try:
        resp = client.list_objects_v2(Bucket=bucket(), Prefix=BE_VERSION_PREFIX)
    except Exception as exc:  # noqa: BLE001
        _last_error = _err(exc)
        return []
    items = resp.get("Contents", []) or []
    versions: list[dict] = []
    for it in items:
        key = it["Key"]
        if not key.endswith("_meta.json"):
            continue
        # Strip prefix and "_meta.json" -> e.g. "2026-05_W3"
        slot = key[len(BE_VERSION_PREFIX):-len("_meta.json")]
        if "_W" not in slot:
            continue
        month, week = slot.rsplit("_", 1)
        meta_raw = download_bytes(key)
        meta = {}
        try:
            meta = json.loads(meta_raw) if meta_raw else {}
        except Exception:  # noqa: BLE001
            pass
        versions.append({
            "month": month, "week": week,
            "sheet": meta.get("sheet", ""),
            "uploaded": meta.get("uploaded", ""),
            "key": _be_version_key(month, week),
        })
    versions.sort(key=lambda v: (v["month"], v["week"]), reverse=True)
    return versions


def latest_be_version() -> dict | None:
    """Latest week of the latest month. None if no versions stored."""
    vs = list_be_versions()
    return vs[0] if vs else None
