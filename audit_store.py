import hashlib
import streamlit as st
from supabase import create_client
from postgrest.exceptions import APIError


def _sb():
    url = st.secrets.get("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        raise RuntimeError("Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY in secrets")
    return create_client(url, key)


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def upload_export_bytes(*, content: bytes, object_path: str) -> str:
    """
    Upload Excel bytes to Supabase Storage.
    This supabase client may NOT support upsert=... argument.
    We do: upload -> if conflict/exists then update.
    """
    sb = _sb()
    bucket = st.secrets.get("SUPABASE_BUCKET", "work-efficiency-exports")

    file_options = {
        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    try:
        sb.storage.from_(bucket).upload(object_path, content, file_options=file_options)
    except APIError as e:
        # 常見：檔案已存在（409 / conflict）→ 改用 update 覆蓋
        msg = str(e)
        if "409" in msg or "Conflict" in msg or "already exists" in msg:
            sb.storage.from_(bucket).update(object_path, content, file_options=file_options)
        else:
            raise

    return object_path


def insert_audit_run(payload: dict) -> dict:
    sb = _sb()
    res = sb.schema("public").table("audit_runs").insert(payload).execute()
    return res.data[0] if res.data else {}
