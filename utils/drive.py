"""
Google Drive 연동: 전체요약 DOCX 업로드/다운로드/목록 조회.
서비스 계정 credentials는 st.secrets["gcp_service_account"] 사용.
"""
import io
import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

_SCOPES = [
    "https://www.googleapis.com/auth/drive",
]
_FOLDER_NAME = "kofeed_reports"
_SUMMARY_FILENAME = "전체요약_보고서.docx"
_MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _service():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=_SCOPES
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _get_or_create_folder(service) -> str:
    """'kofeed_reports' 폴더 ID 반환. 없으면 생성."""
    q = (f"name='{_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder'"
         " and trashed=false")
    res = service.files().list(q=q, fields="files(id)").execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    meta = {
        "name": _FOLDER_NAME,
        "mimeType": "application/vnd.google-apps.folder",
    }
    f = service.files().create(body=meta, fields="id").execute()
    return f["id"]


def upload_summary_docx(docx_bytes: bytes, filename: str = _SUMMARY_FILENAME) -> str:
    """
    DOCX를 Drive kofeed_reports 폴더에 업로드.
    같은 이름 파일이 있으면 내용 덮어쓰기(update), 없으면 새로 생성.
    반환: file_id
    """
    service = _service()
    folder_id = _get_or_create_folder(service)

    # 기존 파일 검색
    q = (f"name='{filename}' and '{folder_id}' in parents and trashed=false")
    res = service.files().list(q=q, fields="files(id)").execute()
    existing = res.get("files", [])

    media = MediaIoBaseUpload(io.BytesIO(docx_bytes), mimetype=_MIME_DOCX, resumable=False)

    if existing:
        file_id = existing[0]["id"]
        service.files().update(fileId=file_id, media_body=media).execute()
        return file_id
    else:
        meta = {"name": filename, "parents": [folder_id]}
        f = service.files().create(body=meta, media_body=media, fields="id").execute()
        return f["id"]


def download_summary_docx(filename: str = _SUMMARY_FILENAME) -> bytes | None:
    """Drive에서 전체요약 DOCX 다운로드. 없으면 None."""
    service = _service()
    folder_id = _get_or_create_folder(service)
    q = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
    res = service.files().list(q=q, fields="files(id,name,modifiedTime)").execute()
    files = res.get("files", [])
    if not files:
        return None
    file_id = files[0]["id"]
    buf = io.BytesIO()
    request = service.files().get_media(fileId=file_id)
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def get_summary_docx_info(filename: str = _SUMMARY_FILENAME) -> dict | None:
    """Drive 파일 메타 정보 반환. {name, modifiedTime} or None."""
    try:
        service = _service()
        folder_id = _get_or_create_folder(service)
        q = f"name='{filename}' and '{folder_id}' in parents and trashed=false"
        res = service.files().list(
            q=q, fields="files(id,name,modifiedTime)"
        ).execute()
        files = res.get("files", [])
        return files[0] if files else None
    except Exception:
        return None
