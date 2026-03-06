"""
동적 설정 관리 모듈.
Google Sheets 'config' 탭에 설정을 저장/불러옴.
설정이 없으면 기본값으로 자동 초기화.
"""
import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# ── 컬럼 정의 ─────────────────────────────────────────────────
CONFIG_COLS = ["type", "group", "name", "samples", "order", "enabled"]

# ── 기본 설정값 ───────────────────────────────────────────────
_DEFAULT_ROWS = [
    # 사료 종류
    ("sample", "", "축우사료",  "", 1, True),
    ("sample", "", "양계사료",  "", 2, True),
    ("sample", "", "고양이사료","", 3, True),
    # 일반성분
    ("component", "일반성분", "수분",    "all", 1, True),
    ("component", "일반성분", "조단백질","all", 2, True),
    ("component", "일반성분", "조지방",  "all", 3, True),
    ("component", "일반성분", "조섬유",  "all", 4, True),
    ("component", "일반성분", "조회분",  "all", 5, True),
    ("component", "일반성분", "NFE",     "all", 6, True),
    # ADF/NDF
    ("component", "ADF/NDF", "ADF", "축우사료", 1, True),
    ("component", "ADF/NDF", "NDF", "축우사료", 2, True),
    # 아미노산
    ("component", "아미노산", "ASP",  "all", 1, True),
    ("component", "아미노산", "THR",  "all", 2, True),
    ("component", "아미노산", "SER",  "all", 3, True),
    ("component", "아미노산", "GLU",  "all", 4, True),
    ("component", "아미노산", "GLY",  "all", 5, True),
    ("component", "아미노산", "ALA",  "all", 6, True),
    ("component", "아미노산", "VAL",  "all", 7, True),
    ("component", "아미노산", "ISOL", "all", 8, True),
    ("component", "아미노산", "LEU",  "all", 9, True),
    ("component", "아미노산", "TYR",  "all", 10, True),
    ("component", "아미노산", "PHE",  "all", 11, True),
    ("component", "아미노산", "LYS",  "all", 12, True),
    ("component", "아미노산", "HIS",  "all", 13, True),
    ("component", "아미노산", "ARG",  "all", 14, True),
    ("component", "아미노산", "PRO",  "all", 15, True),
    ("component", "아미노산", "MET",  "all", 16, True),
    ("component", "아미노산", "CYS",  "all", 17, True),
]

DEFAULT_CONFIG = pd.DataFrame(_DEFAULT_ROWS, columns=CONFIG_COLS)


# ── Google Sheets 접근 ────────────────────────────────────────
def _get_spreadsheet():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPE
    )
    client = gspread.authorize(creds)
    return client.open(st.secrets["sheet"]["name"])


def _get_or_create_config_sheet():
    sp = _get_spreadsheet()
    try:
        return sp.worksheet("config")
    except gspread.WorksheetNotFound:
        ws = sp.add_worksheet(title="config", rows=300, cols=10)
        df = DEFAULT_CONFIG.copy()
        ws.update([df.columns.tolist()] + df.astype(str).values.tolist())
        return ws


# ── 공개 API ──────────────────────────────────────────────────
@st.cache_data(ttl=120)
def get_config() -> pd.DataFrame:
    """설정 시트에서 config DataFrame 읽기. 없으면 기본값 반환."""
    try:
        ws = _get_or_create_config_sheet()
        records = ws.get_all_records()
        if not records:
            return DEFAULT_CONFIG.copy()
        df = pd.DataFrame(records)
        df["enabled"] = df["enabled"].astype(str).str.lower().isin(["true", "1", "yes"])
        df["order"]   = pd.to_numeric(df["order"], errors="coerce").fillna(0).astype(int)
        return df
    except Exception:
        return DEFAULT_CONFIG.copy()


def save_config(df: pd.DataFrame):
    """DataFrame을 config 시트에 저장 후 캐시 초기화."""
    ws = _get_or_create_config_sheet()
    ws.clear()
    ws.update([df.columns.tolist()] + df.fillna("").astype(str).values.tolist())
    get_config.clear()


# ── 설정 파싱 헬퍼 ────────────────────────────────────────────
def get_samples(cfg: pd.DataFrame = None) -> list[str]:
    """활성화된 사료 종류 목록"""
    if cfg is None:
        cfg = get_config()
    return (
        cfg[(cfg["type"] == "sample") & (cfg["enabled"])]
        .sort_values("order")["name"]
        .tolist()
    )


def get_component_groups(cfg: pd.DataFrame = None) -> dict[str, list[dict]]:
    """
    {그룹명: [{"name": str, "samples": [str]}, ...]}
    samples: 해당 성분을 입력할 사료 종류 목록
    """
    if cfg is None:
        cfg = get_config()
    all_samples = get_samples(cfg)
    comps = (
        cfg[(cfg["type"] == "component") & (cfg["enabled"])]
        .sort_values("order")
    )
    groups: dict[str, list[dict]] = {}
    for _, row in comps.iterrows():
        g = row["group"]
        if g not in groups:
            groups[g] = []
        raw = str(row.get("samples", "")).strip()
        if raw in ("all", "", "전체"):
            applicable = all_samples
        else:
            applicable = [s.strip() for s in raw.split(",") if s.strip() in all_samples]
        groups[g].append({"name": row["name"], "samples": applicable})
    return groups


def get_nir_groups(cfg: pd.DataFrame = None) -> dict[str, list[dict]]:
    """NIR 측정 대상 그룹 (아미노산 제외)"""
    groups = get_component_groups(cfg)
    return {g: items for g, items in groups.items() if g != "아미노산"}


def get_all_value_columns(cfg: pd.DataFrame = None) -> list[str]:
    """제출 폼에서 생성되는 모든 값 컬럼명 목록"""
    groups = get_component_groups(cfg)
    nir    = get_nir_groups(cfg)
    cols = []
    for items in groups.values():
        for item in items:
            for s in item["samples"]:
                cols.append(f"{item['name']}_{s}")
    for items in nir.values():
        for item in items:
            for s in item["samples"]:
                cols.append(f"NIR_{item['name']}_{s}")
    return cols


def is_value_col(col: str, samples: list[str] = None) -> bool:
    if samples is None:
        samples = get_samples()
    return any(col.endswith(f"_{s}") for s in samples)


def get_component_from_col(col: str, samples: list[str] = None) -> str | None:
    if samples is None:
        samples = get_samples()
    for s in samples:
        if col.endswith(f"_{s}"):
            prefix = col[: -len(f"_{s}")]
            return prefix[4:] if prefix.startswith("NIR_") else prefix
    return None


def get_sample_from_col(col: str, samples: list[str] = None) -> str | None:
    if samples is None:
        samples = get_samples()
    for s in samples:
        if col.endswith(f"_{s}"):
            return s
    return None
