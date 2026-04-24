import streamlit as st
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))
from utils.sheets import submit_data, get_submitted_by_institution, save_draft, load_draft, delete_draft
from utils.report import generate_submission_pdf
from utils.config import (
    get_config, get_samples, get_component_groups,
    get_info_fields, get_method_options, get_solvent_options, get_questions,
    get_participant_map, get_participant_password_map,
)
from utils.email_sender import send_confirmation

st.set_page_config(
    page_title="회원사 비교분석 시험 데이터 제출",
    page_icon=None,
    layout="wide",
)

# 숫자 입력 UI 커스텀
st.markdown("""
<style>
/* 브라우저 기본 spin 버튼 제거 */
input[type=number]::-webkit-inner-spin-button,
input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none !important; }
input[type=number] { -moz-appearance: textfield !important; }

/* Streamlit +/- 버튼 숨기기 */
button[data-testid="stNumberInputStepUp"],
button[data-testid="stNumberInputStepDown"] { display: none !important; }

/* "Press Enter to apply" 힌트 숨기기 */
div[data-testid="InputInstructions"] { display: none !important; }

/* 숫자 입력 중앙 정렬 */
div[data-testid="stNumberInput"] input { text-align: center !important; }
</style>
""", unsafe_allow_html=True)

# ── 설정 로드 ─────────────────────────────────────────────────
cfg             = get_config()
SAMPLES         = get_samples(cfg)
GROUPS          = get_component_groups(cfg)
INFO_FIELDS     = get_info_fields(cfg)
QUESTIONS       = get_questions(cfg)
PARTICIPANT_MAP = get_participant_map(cfg)

# 회사명에 해당하는 필드명 사전 결정
_INST_FIELD = next(
    (f["name"] for f in INFO_FIELDS if any(kw in f["name"] for kw in ("기관", "회사", "기업", "업체"))),
    INFO_FIELDS[0]["name"] if INFO_FIELDS else "회사명"
)

# ── 참가코드로 접속 제한 ──────────────────────────────────────
if PARTICIPANT_MAP:
    PARTICIPANT_PASSWORDS = get_participant_password_map(cfg)

    if "app_code" not in st.session_state:
        st.session_state.app_code = ""

    if not st.session_state.app_code:
        st.title("한국사료협회 비교분석 시험")
        st.caption("해당 페이지 저작권은 한국사료협회 사료기술연구소에 있습니다.")
        code_try = st.text_input("참가코드", placeholder="예: A99", key="_code_input").strip()
        pw_try   = st.text_input("비밀번호", placeholder="예: 1234", type="password", key="_pw_input").strip()
        if st.button("입장", type="primary"):
            if code_try not in PARTICIPANT_MAP:
                st.error("등록되지 않은 코드입니다. 다시 확인해주세요.")
            else:
                expected_pw = PARTICIPANT_PASSWORDS.get(code_try, "")
                if expected_pw and pw_try != expected_pw:
                    st.error("비밀번호가 올바르지 않습니다.")
                else:
                    st.session_state.app_code = code_try
                    st.rerun()
        st.stop()

    # 인증된 코드와 회사명
    _entered_code    = st.session_state.app_code
    _company_from_code = PARTICIPANT_MAP.get(_entered_code, "")
else:
    _entered_code      = ""
    _company_from_code = ""

# ── 임시저장 스냅샷 수집 ──────────────────────────────────────
def _take_draft_snapshot() -> dict:
    snap = {}
    for field in INFO_FIELDS:
        k = f"info_{field['name']}"
        if k in st.session_state:
            snap[k] = st.session_state[k]
    for group_name, items in GROUPS.items():
        for item in items:
            comp  = item["name"]
            extra = st.session_state.get(f"{group_name}_{comp}_extra", 0)
            if extra:
                snap[f"{group_name}_{comp}_extra"] = extra
            for sfx in [""] + [f"_{i+2}" for i in range(extra)]:
                for s in item["samples"]:
                    k = f"{group_name}_{comp}_{s}{sfx}"
                    if k in st.session_state:
                        snap[k] = st.session_state[k]
                for sub in ("method", "equip", "solvent"):
                    k = f"{group_name}_{comp}_{sub}{sfx}"
                    if k in st.session_state:
                        snap[k] = st.session_state[k]
    for q in QUESTIONS:
        if q["type"] == "multicheck":
            for opt in q["options"]:
                k = f"q_{q['id']}_{opt}"
                if k in st.session_state:
                    snap[k] = st.session_state[k]
        else:
            k = f"q_{q['id']}"
            if k in st.session_state:
                snap[k] = st.session_state[k]
    return snap


# ── 검증 + 데이터 수집 ────────────────────────────────────────
def _collect_and_validate(info_values: dict) -> tuple[list[str], dict | None]:
    import math
    errors, all_data = [], {}
    for group_name, items in GROUPS.items():
        for item in items:
            comp = item["name"]
            extra_count = st.session_state.get(f"{group_name}_{comp}_extra", 0)
            suffixes = [""] + [f"_{i+2}" for i in range(extra_count)]
            for sfx in suffixes:
                for s in item["samples"]:
                    v = st.session_state.get(f"{group_name}_{comp}_{s}{sfx}")
                    if v is not None:
                        all_data[f"{comp}_{s}{sfx}"] = v
                for sub, dk in [("method", f"{comp}_방법{sfx}"),
                                 ("equip",  f"{comp}_기기{sfx}"),
                                 ("solvent",f"{comp}_용매{sfx}")]:
                    v = st.session_state.get(f"{group_name}_{comp}_{sub}{sfx}")
                    if v is not None:
                        all_data[dk] = v
    for field in INFO_FIELDS:
        val = info_values.get(field["name"], "").strip()
        if field["required"] and not val:
            errors.append(f"{field['name']}을(를) 입력해주세요.")
        elif field["email"] and val and "@" not in val:
            errors.append(f"{field['name']}: 올바른 이메일 주소를 입력해주세요.")
    for q in QUESTIONS:
        if not q.get("required"):
            continue
        val = all_data.get(f"Q_{q['id']}", "")
        if q["type"] == "choice" and (not val or val == "(선택 안 함)"):
            errors.append(f"[추가설문] '{q['text']}' 항목은 필수입니다.")
        elif q["type"] in ("text", "multicheck") and not str(val).strip():
            errors.append(f"[추가설문] '{q['text']}' 항목은 필수입니다.")
    for group_name, group_items in GROUPS.items():
        for item in group_items:
            comp = item["name"]
            extra_count = st.session_state.get(f"{group_name}_{comp}_extra", 0)
            suffixes = [""] + [f"_{i+2}" for i in range(extra_count)]
            for sfx in suffixes:
                method_val  = st.session_state.get(f"{group_name}_{comp}_method{sfx}", "") or ""
                solvent_val = st.session_state.get(f"{group_name}_{comp}_solvent{sfx}", "") or ""
                any_value   = any(
                    st.session_state.get(f"{group_name}_{comp}_{s}{sfx}") is not None
                    for s in item["samples"]
                )
                label = comp if sfx == "" else f"{comp}(추가{sfx})"
                if any_value and bool(get_method_options(cfg, comp=comp)) and not method_val:
                    errors.append(f"[{label}] 값을 입력한 경우 방법을 선택해야 합니다.")
                if any_value and item.get("use_solvent", True) and not solvent_val:
                    errors.append(f"[{label}] 값을 입력한 경우 용매를 선택해야 합니다.")
    if errors:
        return errors, None
    inst_field_name = next(
        (f["name"] for f in INFO_FIELDS if any(kw in f["name"] for kw in ("기관", "회사", "기업", "업체"))),
        INFO_FIELDS[0]["name"] if INFO_FIELDS else "기관명"
    )
    email_field_name = next((f["name"] for f in INFO_FIELDS if f["email"]), None)
    inst_name = info_values.get(inst_field_name, "").strip()
    email_to  = info_values.get(email_field_name, "").strip() if email_field_name else ""
    row = {"제출일시": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")}
    for field in INFO_FIELDS:
        row[field["name"]] = info_values.get(field["name"], "").strip()
    for k, v in all_data.items():
        if v is None:
            row[k] = ""
        elif isinstance(v, float) and math.isnan(v):
            row[k] = ""
        else:
            row[k] = v
    return [], {"row": row, "inst_name": inst_name, "email_to": email_to}


@st.dialog("제출 확인")
def _submit_dialog():
    st.markdown("정말로 제출하시겠습니까?")
    st.markdown("**제출 후에는 수정이 불가합니다.**")
    st.markdown("")
    _d1, _d2 = st.columns(2)
    with _d1:
        if st.button("✅ 제출", type="primary", use_container_width=True):
            _p     = st.session_state.get("_pending_submission", {})
            _row   = _p.get("row", {})
            _inst  = _p.get("inst_name", "")
            _email = _p.get("email_to", "")
            with st.spinner("제출 중..."):
                try:
                    submit_data(_row)
                    if _entered_code:
                        try:
                            delete_draft(_entered_code)
                            st.session_state._draft_state = "none"
                        except Exception:
                            pass
                    st.session_state._submit_confirm = False
                    st.session_state._just_submitted = True
                    st.session_state._submit_inst    = _inst
                    st.session_state._submit_email   = _email
                    st.session_state._submit_row     = _row
                    st.rerun()
                except Exception as _e:
                    st.error(f"제출 중 오류가 발생했습니다: {_e}")
    with _d2:
        if st.button("❌ 취소", use_container_width=True):
            st.session_state._submit_confirm = False
            st.rerun()


# ── 값 파싱 헬퍼 ──────────────────────────────────────────────
def parse_float(s: str, free_decimal: bool = False):
    """문자열 → float, 빈 값이면 None. free_decimal=False면 소수점 2자리로 반올림."""
    s = s.strip().replace(",", ".")
    if not s:
        return None
    try:
        v = float(s)
        return v if free_decimal else round(v, 2)
    except ValueError:
        return "ERR"


# ── 성분 입력 테이블 ──────────────────────────────────────────
def component_table(items: list[dict], prefix: str = "") -> dict:
    data = {}

    all_sample_set: list[str] = []
    for item in items:
        for s in item["samples"]:
            if s not in all_sample_set:
                all_sample_set.append(s)

    col_ratios = [2, 2, 2, 2] + [2] * len(all_sample_set)
    hcols = st.columns(col_ratios)
    hcols[0].markdown("**성분**")
    hcols[1].markdown("**방법**")
    hcols[2].markdown("**기기명**")
    hcols[3].markdown("**용매**")
    for i, s in enumerate(all_sample_set):
        hcols[4 + i].markdown(f"**{s}**")
    st.markdown("<hr style='margin:2px 0 6px 0; border-color:#e0e0e0'>",
                unsafe_allow_html=True)

    for item in items:
        comp        = item["name"]
        samples     = item["samples"]
        use_equip   = item.get("use_equip", True)
        use_solvent = item.get("use_solvent", True)
        free_dec    = item.get("free_decimal", False)
        fmt         = "%.4f" if free_dec else "%.2f"
        step        = 0.0001 if free_dec else 0.01
        method_opts  = get_method_options(cfg, comp=comp)
        use_method   = bool(method_opts)
        method_list  = [""] + method_opts
        solvent_list = [""] + get_solvent_options(cfg, comp=comp)

        # 이 성분에 추가된 방법 행 수 (0 = 추가 없음, 기본 1행만)
        extra_key   = f"{prefix}{comp}_extra"
        extra_count = st.session_state.get(extra_key, 0)

        def _render_row(suffix: str):
            """suffix="" → 기존 행, suffix="_2","_3"... → 추가 행"""
            cols = st.columns(col_ratios)
            cols[0].markdown(
                f"**{comp}**" if suffix == "" else
                f"<span style='color:#888;font-size:0.85em;padding-left:12px'>↳ {comp}</span>",
                unsafe_allow_html=True,
            )
            with cols[1]:
                if use_method:
                    sel = st.selectbox(
                        "방법", method_list,
                        key=f"{prefix}{comp}_method{suffix}",
                        label_visibility="collapsed",
                    )
                    data[f"{comp}_방법{suffix}"] = sel
                else:
                    st.markdown("<div style='color:#ccc;text-align:center'>—</div>", unsafe_allow_html=True)
                    data[f"{comp}_방법{suffix}"] = ""
            with cols[2]:
                if use_equip:
                    data[f"{comp}_기기{suffix}"] = st.text_input(
                        "기기", key=f"{prefix}{comp}_equip{suffix}",
                        label_visibility="collapsed", placeholder="기기명",
                    )
                else:
                    st.markdown("<div style='color:#ccc;text-align:center'>—</div>", unsafe_allow_html=True)
                    data[f"{comp}_기기{suffix}"] = ""
            with cols[3]:
                if use_solvent:
                    data[f"{comp}_용매{suffix}"] = st.selectbox(
                        "용매", solvent_list,
                        key=f"{prefix}{comp}_solvent{suffix}",
                        label_visibility="collapsed",
                    )
                else:
                    st.markdown("<div style='color:#ccc;text-align:center'>—</div>", unsafe_allow_html=True)
                    data[f"{comp}_용매{suffix}"] = ""
            for i, s in enumerate(all_sample_set):
                with cols[4 + i]:
                    if s in samples:
                        data[f"{comp}_{s}{suffix}"] = st.number_input(
                            s, key=f"{prefix}{comp}_{s}{suffix}",
                            value=None, min_value=0.0,
                            step=step, format=fmt,
                            placeholder="0.00",
                            label_visibility="collapsed",
                        )
                    else:
                        st.markdown("<div style='color:#ccc;text-align:center'>—</div>", unsafe_allow_html=True)

        # 기본 행 (suffix="" → 기존 키 그대로)
        _render_row("")

        # 추가 행들 (suffix="_2", "_3", ...)
        for idx in range(extra_count):
            _render_row(f"_{idx + 2}")

        # + / - 버튼 (allow_multi 설정이 True인 성분만 표시)
        if item.get("allow_multi", False):
            btn_cols = st.columns([8, 1, 1])
            with btn_cols[1]:
                if st.button("＋", key=f"add_{prefix}{comp}", help=f"{comp} 방법 추가",
                             use_container_width=True):
                    st.session_state[extra_key] = extra_count + 1
                    st.rerun()
            with btn_cols[2]:
                if extra_count > 0:
                    if st.button("－", key=f"del_{prefix}{comp}", help="마지막 행 삭제",
                                 use_container_width=True):
                        st.session_state[extra_key] = extra_count - 1
                        st.rerun()

    return data



# ── 입력 폼 ──────────────────────────────────────────────────
_title_col, _draft_col = st.columns([8, 2])
with _title_col:
    st.title("한국사료협회 비교분석 시험")
    if _company_from_code:
        st.caption(f"참가코드: **{_entered_code}** | 회사명: **{_company_from_code}**")
    else:
        st.caption("제출하신 데이터는 적어주신 이메일로 발송됩니다.")
with _draft_col:
    st.markdown("<div style='padding-top:1.4rem'></div>", unsafe_allow_html=True)
    if _entered_code and st.button("💾 임시저장", key="_draft_save_btn", use_container_width=True):
        _snap = _take_draft_snapshot()
        _now  = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")
        try:
            save_draft(_entered_code, _now, _snap)
            st.session_state._draft_state = "none"
            st.toast(f"임시저장 완료  ({_now})", icon="💾")
        except Exception as _e:
            st.error(f"임시저장 실패: {_e}")
st.divider()

# ── 임시저장 복원 배너 (세션당 1회 API 호출) ─────────────────
if _entered_code and "_draft_state" not in st.session_state:
    _d, _at = load_draft(_entered_code)
    if _d:
        st.session_state._draft_state = "pending"
        st.session_state._draft_ss    = _d
        st.session_state._draft_at    = _at
    else:
        st.session_state._draft_state = "none"

if st.session_state.get("_draft_state") == "pending":
    _at_disp = st.session_state.get("_draft_at", "")
    st.info(f"💾 이전에 임시저장한 데이터가 있습니다.  ·  저장일시: **{_at_disp}**")
    _dc1, _dc2, _ = st.columns([2, 2, 6])
    with _dc1:
        if st.button("📂 불러오기", use_container_width=True):
            for _k, _v in st.session_state._draft_ss.items():
                st.session_state[_k] = _v
            st.session_state._draft_state = "loaded"
            st.rerun()
    with _dc2:
        if st.button("✕ 무시", use_container_width=True):
            st.session_state._draft_state = "dismissed"
            st.rerun()

# ── 이미 제출한 데이터 표시 ───────────────────────────────────
if _company_from_code:
    _prev = get_submitted_by_institution(_company_from_code, _INST_FIELD)
    if _prev is not None:
        latest = _prev.iloc[-1].to_dict()
        submitted_at = latest.get("제출일시", "-")
        st.success(f"✅ 이미 데이터를 제출하셨습니다. (제출일시: {submitted_at})")
        _pdf_bytes = generate_submission_pdf(latest, cfg, generated_at=submitted_at)
        st.download_button(
            label="📄 데이터 제출 확인서 다운로드",
            data=_pdf_bytes,
            file_name=f"제출확인서_{_entered_code}_{submitted_at[:10]}.pdf",
            mime="application/pdf",
        )
        st.divider()

st.subheader("기관 정보")
info_cols = st.columns(max(len(INFO_FIELDS), 1))
info_values: dict[str, str] = {}
for i, field in enumerate(INFO_FIELDS):
    with info_cols[i]:
        label = field["name"] + (" *" if field["required"] else "")
        if PARTICIPANT_MAP and field["name"] == _INST_FIELD:
            info_values[field["name"]] = _company_from_code
            st.text_input(label, value=_company_from_code, disabled=True, key=f"info_{field['name']}")
        else:
            info_values[field["name"]] = st.text_input(
                label, placeholder=field["placeholder"], key=f"info_{field['name']}"
            )
st.divider()

all_data: dict = {}

for group_name, items in GROUPS.items():
    st.subheader(group_name)
    all_data.update(component_table(items, prefix=f"{group_name}_"))
    st.divider()

if QUESTIONS:
    st.subheader("추가 설문")
    for q in QUESTIONS:
        label = q["text"] + (" *" if q.get("required") else "")
        if q["type"] == "text":
            all_data[f"Q_{q['id']}"] = st.text_area(
                label, key=f"q_{q['id']}", placeholder=q["hint"],
            )
        elif q["type"] == "choice":
            all_data[f"Q_{q['id']}"] = st.radio(
                label, ["(선택 안 함)"] + q["options"],
                key=f"q_{q['id']}", horizontal=True,
            )
        elif q["type"] == "multicheck":
            st.markdown(f"**{label}**")
            selected = [
                opt for opt in q["options"]
                if st.checkbox(opt, key=f"q_{q['id']}_{opt}")
            ]
            all_data[f"Q_{q['id']}"] = ", ".join(selected)
    st.divider()

# ── 데이터 제출 버튼 ──────────────────────────────────────────
if st.button("데이터 제출", type="primary", use_container_width=True):
    _errs, _result = _collect_and_validate(info_values)
    if _errs:
        st.session_state._validation_errors = _errs
        st.session_state._submit_confirm = False
    else:
        st.session_state._validation_errors = []
        st.session_state._pending_submission = _result
        st.session_state._submit_confirm = True

for _e in st.session_state.get("_validation_errors", []):
    st.error(_e)

if st.session_state.get("_submit_confirm"):
    _submit_dialog()

# ── 제출 성공 처리 ────────────────────────────────────────────
if st.session_state.get("_just_submitted"):
    st.session_state._just_submitted = False
    _inst_ok  = st.session_state.pop("_submit_inst",  "기관")
    _email_ok = st.session_state.pop("_submit_email", "")
    _row_ok   = st.session_state.pop("_submit_row",   {})
    st.success(f"{_inst_ok or '기관'}의 데이터가 성공적으로 제출되었습니다!")
    st.balloons()
    if _email_ok:
        with st.spinner("접수 확인 이메일 발송 중..."):
            try:
                send_confirmation(_email_ok, _inst_ok, _row_ok, cfg)
                st.info(f"접수 확인 이메일이 {_email_ok}으로 발송되었습니다.")
            except Exception as _e:
                st.warning(f"이메일 발송 중 오류가 발생했습니다: {_e}")
