import json
import time
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

BASE_FIELDS = ["제출일시", "기관명", "담당자명", "이메일", "전화"]


_DATA_SHEET = "제출데이터"


def _get_sheet():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPE
    )
    client = gspread.authorize(creds)
    sp = client.open(st.secrets["sheet"]["name"])
    try:
        return sp.worksheet(_DATA_SHEET)
    except gspread.WorksheetNotFound:
        old = sp.sheet1
        old.update_title(_DATA_SHEET)
        return old


def _fmt(v):
    """숫자면 소수점 2자리 반올림, 아니면 그대로."""
    if v is None:
        return ""
    try:
        f = float(v)
        return round(f, 2)
    except (TypeError, ValueError):
        return v


def submit_data(row: dict):
    sheet = _get_sheet()
    existing = sheet.get_all_values()
    if not existing:
        sheet.append_row(list(row.keys()))
    else:
        existing_headers = existing[0]
        new_keys = [k for k in row.keys() if k not in existing_headers]
        if new_keys:
            sheet.update("A1", [existing_headers + new_keys])
    headers = sheet.row_values(1)
    sheet.append_row([_fmt(row.get(h)) for h in headers])
    get_all_data.clear()


@st.cache_data(ttl=60)
def get_all_data() -> pd.DataFrame:
    sheet = _get_sheet()
    for attempt in range(3):
        try:
            records = sheet.get_all_records()
            return pd.DataFrame(records)
        except gspread.exceptions.APIError as e:
            if attempt == 2:
                raise
            time.sleep(2 ** attempt)
    return pd.DataFrame()


_DRAFT_SHEET = "임시저장"


def _get_draft_sheet():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPE
    )
    client = gspread.authorize(creds)
    sp = client.open(st.secrets["sheet"]["name"])
    try:
        return sp.worksheet(_DRAFT_SHEET)
    except gspread.WorksheetNotFound:
        return sp.add_worksheet(title=_DRAFT_SHEET, rows=200, cols=3)


def save_draft(code: str, saved_at: str, data: dict):
    """참가코드에 해당하는 임시저장 덮어쓰기. data는 session_state 스냅샷."""
    ws = _get_draft_sheet()
    data_json = json.dumps(data, ensure_ascii=False, default=str)
    existing = ws.get_all_records()
    row = [code, saved_at, data_json]
    if not existing:
        ws.update("A1", [["code", "saved_at", "data"], row])
        return
    for i, r in enumerate(existing):
        if str(r.get("code", "")) == code:
            ws.update(f"A{i + 2}", [row])
            return
    ws.append_row(row)


def load_draft(code: str) -> tuple[dict | None, str]:
    """(data_dict, saved_at) 반환. 없으면 (None, "")."""
    try:
        ws = _get_draft_sheet()
        for r in ws.get_all_records():
            if str(r.get("code", "")) == code:
                return json.loads(r.get("data", "{}")), str(r.get("saved_at", ""))
        return None, ""
    except Exception:
        return None, ""


def delete_draft(code: str):
    """참가코드에 해당하는 임시저장 행 삭제."""
    try:
        ws = _get_draft_sheet()
        existing = ws.get_all_records()
        if not existing:
            return
        headers = list(existing[0].keys())
        kept = [r for r in existing if str(r.get("code", "")) != code]
        ws.clear()
        ws.update("A1", [headers] + [[r.get(h, "") for h in headers] for r in kept])
    except Exception:
        pass


def get_submitted_by_institution(institution_name: str, inst_field: str = "기관명") -> pd.DataFrame | None:
    """특정 기관명으로 제출된 데이터 반환 (최신순). 없으면 None."""
    df = get_all_data()
    if df.empty or inst_field not in df.columns:
        return None
    rows = df[df[inst_field].astype(str) == institution_name]
    return rows.reset_index(drop=True) if not rows.empty else None
