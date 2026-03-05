import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

BASE_FIELDS = [
    "제출일시", "이메일", "기관명",
    "ASP", "THR", "SER", "GLU", "GLY", "ALA",
    "VAL", "ISOL", "LEU", "TYR", "PHE",
    "LYS", "HIS", "ARG", "PRO", "MET", "CYS",
]


def _get_sheet():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPE
    )
    client = gspread.authorize(creds)
    sheet = client.open(st.secrets["sheet"]["name"]).sheet1
    return sheet


def submit_data(row: dict):
    sheet = _get_sheet()
    # 헤더가 없으면 첫 행에 추가
    existing = sheet.get_all_values()
    if not existing:
        all_keys = list(row.keys())
        sheet.append_row(all_keys)
    else:
        existing_headers = existing[0]
        # 새 컬럼 헤더 병합
        new_keys = [k for k in row.keys() if k not in existing_headers]
        if new_keys:
            sheet.update(f"A1", [existing_headers + new_keys])
    # 데이터 행 추가
    headers = sheet.row_values(1)
    row_values = [row.get(h, "") for h in headers]
    sheet.append_row(row_values)


def get_all_data() -> pd.DataFrame:
    sheet = _get_sheet()
    records = sheet.get_all_records()
    return pd.DataFrame(records)


def get_custom_fields() -> list:
    sheet = _get_sheet()
    headers = sheet.row_values(1)
    return [h for h in headers if h and h not in BASE_FIELDS]
