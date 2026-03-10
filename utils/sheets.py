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
    sheet.append_row([row.get(h, "") for h in headers])


def get_all_data() -> pd.DataFrame:
    sheet = _get_sheet()
    records = sheet.get_all_records()
    return pd.DataFrame(records)
