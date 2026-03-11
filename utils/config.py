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
CONFIG_COLS = ["type", "group", "name", "samples", "order", "enabled", "use_equip", "use_solvent", "free_decimal"]

# ── 기본 설정값 ───────────────────────────────────────────────
_DEFAULT_ROWS = [
    # 방법 옵션 (드롭다운 선택지)
    ("method_option", "", "AOAC 방법",     "", 1, True, True, True, False),
    ("method_option", "", "식약처 고시법", "", 2, True, True, True, False),
    ("method_option", "", "KS 방법",       "", 3, True, True, True, False),
    ("method_option", "", "자체 분석법",   "", 4, True, True, True, False),
    # 기관 정보 필드 — group=placeholder, samples=flags(required/email)
    ("info_field", "○○ 연구소",        "기관명",   "required",       1, True, True, True, False),
    ("info_field", "홍길동",            "담당자명", "required",       2, True, True, True, False),
    ("info_field", "lab@example.com",  "이메일",   "required,email", 3, True, True, True, False),
    ("info_field", "010-0000-0000",    "전화",     "",               4, True, True, True, False),
    # 사료 종류
    ("sample", "", "축우사료",  "", 1, True, True, True, False),
    ("sample", "", "양계사료",  "", 2, True, True, True, False),
    ("sample", "", "고양이사료","", 3, True, True, True, False),
    # 섹션(그룹) — samples 필드: "nir" = NIR 테이블에도 포함
    ("group", "", "일반성분", "nir", 1, True, True, True, False),
    ("group", "", "ADF/NDF",  "nir", 2, True, True, True, False),
    ("group", "", "아미노산", "",    3, True, True, True, False),
    # 일반성분 (use_equip=True, use_solvent=True)
    ("component", "일반성분", "수분",    "all", 1, True, True, True, False),
    ("component", "일반성분", "조단백질","all", 2, True, True, True, False),
    ("component", "일반성분", "조지방",  "all", 3, True, True, True, False),
    ("component", "일반성분", "조섬유",  "all", 4, True, True, True, False),
    ("component", "일반성분", "조회분",  "all", 5, True, True, True, False),
    ("component", "일반성분", "NFE",     "all", 6, True, True, True, False),
    # ADF/NDF
    ("component", "ADF/NDF", "ADF", "축우사료", 1, True, True, True, False),
    ("component", "ADF/NDF", "NDF", "축우사료", 2, True, True, True, False),
    # 아미노산
    ("component", "아미노산", "ASP",  "all", 1, True, True, True, False),
    ("component", "아미노산", "THR",  "all", 2, True, True, True, False),
    ("component", "아미노산", "SER",  "all", 3, True, True, True, False),
    ("component", "아미노산", "GLU",  "all", 4, True, True, True, False),
    ("component", "아미노산", "GLY",  "all", 5, True, True, True, False),
    ("component", "아미노산", "ALA",  "all", 6, True, True, True, False),
    ("component", "아미노산", "VAL",  "all", 7, True, True, True, False),
    ("component", "아미노산", "ISOL", "all", 8, True, True, True, False),
    ("component", "아미노산", "LEU",  "all", 9, True, True, True, False),
    ("component", "아미노산", "TYR",  "all", 10, True, True, True, False),
    ("component", "아미노산", "PHE",  "all", 11, True, True, True, False),
    ("component", "아미노산", "LYS",  "all", 12, True, True, True, False),
    ("component", "아미노산", "HIS",  "all", 13, True, True, True, False),
    ("component", "아미노산", "ARG",  "all", 14, True, True, True, False),
    ("component", "아미노산", "PRO",  "all", 15, True, True, True, False),
    ("component", "아미노산", "MET",  "all", 16, True, True, True, False),
    ("component", "아미노산", "CYS",  "all", 17, True, True, True, False),
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
        # 기존 시트에 없는 컬럼은 True로 기본값 설정 (빈 값도 True)
        for col in ("use_equip", "use_solvent"):
            if col not in df.columns:
                df[col] = True
            else:
                raw = df[col].astype(str).str.strip().str.lower()
                df[col] = ~raw.isin(["false", "0", "no"])
        # free_decimal: 기존 시트에 없으면 False, 빈 값도 False
        if "free_decimal" not in df.columns:
            df["free_decimal"] = False
        else:
            raw = df["free_decimal"].astype(str).str.strip().str.lower()
            df["free_decimal"] = raw.isin(["true", "1", "yes"])
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
def get_method_options(cfg: pd.DataFrame = None, comp: str = None) -> list[str]:
    """
    방법 드롭다운 선택지 목록.
    comp 지정 시: group이 해당 성분명인 행 우선 반환.
                  해당 성분 전용 행이 없으면 group이 빈 공통 행 반환.
    comp 미지정 시: group이 빈 공통 행만 반환.
    """
    if cfg is None:
        cfg = get_config()
    rows = (
        cfg[(cfg["type"] == "method_option") & (cfg["enabled"])]
        .sort_values("order")
    )
    if comp:
        comp_rows = rows[rows["group"].astype(str).str.strip() == comp]
        if not comp_rows.empty:
            return comp_rows["name"].tolist()
    global_rows = rows[rows["group"].astype(str).str.strip() == ""]
    return global_rows["name"].tolist()


def get_questions(cfg: pd.DataFrame = None) -> list[dict]:
    """
    추가 질문 목록.
    samples 형식:
      "text"                    → 주관식
      "text:힌트"               → 주관식 (힌트 표시)
      "choice:옵션1|옵션2"      → 단일 선택
      "multicheck:옵션1|옵션2"  → 복수 선택
    반환: [{"id", "text", "type", "options", "hint"}, ...]
    """
    if cfg is None:
        cfg = get_config()
    rows = (
        cfg[(cfg["type"] == "question") & (cfg["enabled"])]
        .sort_values("order")
    )
    result = []
    for _, row in rows.iterrows():
        q_id   = str(row.get("group", "")).strip()
        q_text = str(row.get("name", "")).strip()
        raw    = str(row.get("samples", "")).strip()

        if raw.startswith("choice:"):
            q_type = "choice"
            opts   = [o.strip() for o in raw[7:].split("|") if o.strip()]
            hint   = ""
        elif raw.startswith("multicheck:"):
            q_type = "multicheck"
            opts   = [o.strip() for o in raw[11:].split("|") if o.strip()]
            hint   = ""
        elif raw.startswith("text:"):
            q_type = "text"
            opts   = []
            hint   = raw[5:]
        else:
            q_type = "text"
            opts   = []
            hint   = ""

        if not q_text:
            continue
        result.append({
            "id":      q_id or f"q{len(result)+1}",
            "text":    q_text,
            "type":    q_type,
            "options": opts,
            "hint":    hint,
        })
    return result


def get_info_fields(cfg: pd.DataFrame = None) -> list[dict]:
    """
    기관 정보 필드 목록.
    반환: [{"name": str, "placeholder": str, "required": bool, "email": bool}, ...]
    """
    if cfg is None:
        cfg = get_config()
    rows = (
        cfg[(cfg["type"] == "info_field") & (cfg["enabled"])]
        .sort_values("order")
    )
    result = []
    for _, row in rows.iterrows():
        flags = str(row.get("samples", "")).lower()
        result.append({
            "name":        row["name"],
            "placeholder": str(row.get("group", "")),
            "required":    "required" in flags,
            "email":       "email" in flags,
        })
    return result


def get_samples(cfg: pd.DataFrame = None) -> list[str]:
    """활성화된 사료 종류 목록"""
    if cfg is None:
        cfg = get_config()
    return (
        cfg[(cfg["type"] == "sample") & (cfg["enabled"])]
        .sort_values("order")["name"]
        .tolist()
    )


def get_group_order(cfg: pd.DataFrame = None) -> list[dict]:
    """
    섹션 순서 목록: [{"name": str, "nir": bool, "enabled": bool}, ...]
    type="group" 행 기준 정렬.
    """
    if cfg is None:
        cfg = get_config()
    grp_rows = cfg[cfg["type"] == "group"].sort_values("order")
    result = []
    for _, row in grp_rows.iterrows():
        result.append({
            "name":    row["name"],
            "nir":     "nir" in str(row.get("samples", "")).lower(),
            "enabled": bool(row["enabled"]),
        })
    return result


def get_component_groups(cfg: pd.DataFrame = None) -> dict[str, list[dict]]:
    """
    {그룹명: [{"name": str, "samples": [str]}, ...]}
    group 행의 순서를 따름. enabled=False 그룹은 제외.
    """
    if cfg is None:
        cfg = get_config()
    all_samples = get_samples(cfg)
    group_order = get_group_order(cfg)
    enabled_groups = [g["name"] for g in group_order if g["enabled"]]

    comps = (
        cfg[(cfg["type"] == "component") & (cfg["enabled"])]
        .sort_values("order")
    )
    # 그룹 순서대로 딕셔너리 초기화
    groups: dict[str, list[dict]] = {g: [] for g in enabled_groups}

    for _, row in comps.iterrows():
        g = row["group"]
        if g not in groups:
            continue  # 비활성 그룹 소속 성분 제외
        raw = str(row.get("samples", "")).strip()
        if raw in ("all", "", "전체"):
            applicable = all_samples
        else:
            applicable = [s.strip() for s in raw.split(",") if s.strip() in all_samples]
        groups[g].append({
            "name":         row["name"],
            "samples":      applicable,
            "use_equip":    bool(row.get("use_equip",    True)),
            "use_solvent":  bool(row.get("use_solvent",  True)),
            "free_decimal": bool(row.get("free_decimal", False)),
        })
    return groups


def get_nir_groups(cfg: pd.DataFrame = None) -> dict[str, list[dict]]:
    """NIR 측정 대상 그룹 (group 행의 nir 플래그 기준)"""
    if cfg is None:
        cfg = get_config()
    group_order = get_group_order(cfg)
    nir_group_names = {g["name"] for g in group_order if g["nir"] and g["enabled"]}
    groups = get_component_groups(cfg)
    return {g: items for g, items in groups.items() if g in nir_group_names}


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
