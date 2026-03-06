import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# ── 상수 ──────────────────────────────────────────────────────
SAMPLES = ["축우사료", "양계사료", "고양이사료"]
PROXIMATE = ["수분", "조단백질", "조지방", "조섬유", "조회분", "NFE"]
CATTLE_ONLY = ["ADF", "NDF"]
AMINO_ACIDS = [
    "ASP", "THR", "SER", "GLU", "GLY", "ALA",
    "VAL", "ISOL", "LEU", "TYR", "PHE",
    "LYS", "HIS", "ARG", "PRO", "MET", "CYS",
]
NIR_COMPONENTS = PROXIMATE + CATTLE_ONLY

BASE_FIELDS = ["제출일시", "기관명", "담당자명", "이메일", "전화"]


# ── 컬럼 헬퍼 ─────────────────────────────────────────────────
def is_value_col(col: str) -> bool:
    """값 컬럼 여부 (사료종류로 끝나는 컬럼)"""
    return any(col.endswith(f"_{s}") for s in SAMPLES)


def get_component(col: str) -> str | None:
    """값 컬럼에서 성분명 추출 (NIR_ 접두사 포함 처리)"""
    for s in SAMPLES:
        if col.endswith(f"_{s}"):
            prefix = col[: -len(f"_{s}")]
            return prefix[4:] if prefix.startswith("NIR_") else prefix
    return None


def get_sample(col: str) -> str | None:
    """값 컬럼에서 사료종류 추출"""
    for s in SAMPLES:
        if col.endswith(f"_{s}"):
            return s
    return None


def is_nir_col(col: str) -> bool:
    return col.startswith("NIR_") and is_value_col(col)


def method_col(comp: str) -> str:
    return f"{comp}_방법"


def _get_sheet():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPE
    )
    client = gspread.authorize(creds)
    sheet = client.open(st.secrets["sheet"]["name"]).sheet1
    return sheet


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
