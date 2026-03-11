import streamlit as st
from datetime import datetime
from utils.sheets import submit_data
from utils.config import (
    get_config, get_samples, get_component_groups, get_nir_groups,
    get_info_fields, get_method_options, get_questions,
)
from utils.email_sender import send_confirmation

st.set_page_config(
    page_title="회원사 비교분석 시험 데이터 제출",
    page_icon=None,
    layout="wide",
)

# +/- 스핀 버튼 숨기기
st.markdown("""
<style>
input[type=number]::-webkit-inner-spin-button,
input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none !important; }
input[type=number] { -moz-appearance: textfield !important; }
</style>
""", unsafe_allow_html=True)

st.title("회원사 비교분석 시험")
st.caption("제출하신 데이터는 적어주신 이메일로 발송됩니다.")
st.divider()

# ── 설정 로드 ─────────────────────────────────────────────────
cfg            = get_config()
SAMPLES        = get_samples(cfg)
GROUPS         = get_component_groups(cfg)
NIR_GRP        = get_nir_groups(cfg)
INFO_FIELDS    = get_info_fields(cfg)
QUESTIONS      = get_questions(cfg)

# ── 기관 정보 ─────────────────────────────────────────────────
st.subheader("기관 정보")
info_cols   = st.columns(max(len(INFO_FIELDS), 1))
info_values: dict[str, str] = {}
for i, field in enumerate(INFO_FIELDS):
    with info_cols[i]:
        label = field["name"] + (" *" if field["required"] else "")
        info_values[field["name"]] = st.text_input(
            label, placeholder=field["placeholder"], key=f"info_{field['name']}"
        )

st.divider()


# ── 값 파싱 헬퍼 ──────────────────────────────────────────────
def parse_float(s: str):
    """문자열 → 소수점 2자리 float, 빈 값이면 None"""
    s = s.strip().replace(",", ".")
    if not s:
        return None
    try:
        return round(float(s), 2)
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
        comp    = item["name"]
        samples = item["samples"]
        use_equip   = item.get("use_equip", True)
        use_solvent = item.get("use_solvent", True)
        cols    = st.columns(col_ratios)
        cols[0].markdown(f"**{comp}**")

        method_list = [""] + get_method_options(cfg, comp=comp)
        with cols[1]:
            sel = st.selectbox(
                "방법", method_list,
                key=f"{prefix}{comp}_method",
                label_visibility="collapsed",
            )
            data[f"{comp}_방법"] = sel

        with cols[2]:
            if use_equip:
                data[f"{comp}_기기"] = st.text_input(
                    "기기", key=f"{prefix}{comp}_equip",
                    label_visibility="collapsed", placeholder="기기명",
                )
            else:
                st.markdown(
                    "<div style='color:#ccc;text-align:center'>—</div>",
                    unsafe_allow_html=True,
                )
                data[f"{comp}_기기"] = ""

        with cols[3]:
            if use_solvent:
                data[f"{comp}_용매"] = st.text_input(
                    "용매", key=f"{prefix}{comp}_solvent",
                    label_visibility="collapsed", placeholder="용매",
                )
            else:
                st.markdown(
                    "<div style='color:#ccc;text-align:center'>—</div>",
                    unsafe_allow_html=True,
                )
                data[f"{comp}_용매"] = ""

        for i, s in enumerate(all_sample_set):
            with cols[4 + i]:
                if s in samples:
                    raw = st.text_input(
                        s, key=f"{prefix}{comp}_{s}",
                        label_visibility="collapsed", placeholder="0.00",
                    )
                    parsed = parse_float(raw)
                    if parsed == "ERR":
                        st.markdown(
                            "<small style='color:red'>숫자만 입력</small>",
                            unsafe_allow_html=True,
                        )
                        data[f"{comp}_{s}"] = None
                    else:
                        data[f"{comp}_{s}"] = parsed
                else:
                    st.markdown(
                        "<div style='color:#ccc;text-align:center'>—</div>",
                        unsafe_allow_html=True,
                    )
    return data


def nir_table(items: list[dict]) -> dict:
    data = {}

    all_sample_set: list[str] = []
    for item in items:
        for s in item["samples"]:
            if s not in all_sample_set:
                all_sample_set.append(s)

    col_ratios = [2, 3] + [2] * len(all_sample_set)
    hcols = st.columns(col_ratios)
    hcols[0].markdown("**성분**")
    hcols[1].markdown("**기기명**")
    for i, s in enumerate(all_sample_set):
        hcols[2 + i].markdown(f"**{s}**")
    st.markdown("<hr style='margin:2px 0 6px 0; border-color:#e0e0e0'>",
                unsafe_allow_html=True)

    for item in items:
        comp    = item["name"]
        samples = item["samples"]
        cols    = st.columns(col_ratios)
        cols[0].markdown(f"**{comp}**")

        with cols[1]:
            data[f"NIR_{comp}_기기"] = st.text_input(
                "기기", key=f"NIR_{comp}_equip",
                label_visibility="collapsed", placeholder="기기명",
            )
        for i, s in enumerate(all_sample_set):
            with cols[2 + i]:
                if s in samples:
                    raw = st.text_input(
                        s, key=f"NIR_{comp}_{s}",
                        label_visibility="collapsed", placeholder="0.00",
                    )
                    parsed = parse_float(raw)
                    data[f"NIR_{comp}_{s}"] = None if parsed == "ERR" else parsed
                else:
                    st.markdown(
                        "<div style='color:#ccc;text-align:center'>—</div>",
                        unsafe_allow_html=True,
                    )
    return data


# ── 성분 그룹별 폼 ────────────────────────────────────────────
all_data: dict = {}

for group_name, items in GROUPS.items():
    st.subheader(group_name)
    all_data.update(component_table(items, prefix=f"{group_name}_"))
    st.divider()

# ── NIR 측정값 ───────────────────────────────────────────────
if NIR_GRP:
    st.subheader("NIR 측정값")
    st.caption("NIR 기기로 측정한 값을 입력하세요.")
    for group_name, items in NIR_GRP.items():
        st.markdown(f"**{group_name}**")
        all_data.update(nir_table(items))
    st.divider()

# ── 추가 질문 ─────────────────────────────────────────────────
if QUESTIONS:
    st.subheader("추가 설문")
    for q in QUESTIONS:
        if q["type"] == "text":
            all_data[f"Q_{q['id']}"] = st.text_area(
                q["text"], key=f"q_{q['id']}", placeholder=q["hint"],
            )
        elif q["type"] == "choice":
            all_data[f"Q_{q['id']}"] = st.radio(
                q["text"], ["(선택 안 함)"] + q["options"],
                key=f"q_{q['id']}", horizontal=True,
            )
        elif q["type"] == "multicheck":
            st.markdown(f"**{q['text']}**")
            selected = [
                opt for opt in q["options"]
                if st.checkbox(opt, key=f"q_{q['id']}_{opt}")
            ]
            all_data[f"Q_{q['id']}"] = ", ".join(selected)
    st.divider()

# ── 제출 ─────────────────────────────────────────────────────
if st.button("데이터 제출", type="primary", use_container_width=True):
    errors = []

    # 기관 정보 검증
    for field in INFO_FIELDS:
        val = info_values.get(field["name"], "").strip()
        if field["required"] and not val:
            errors.append(f"{field['name']}을(를) 입력해주세요.")
        elif field["email"] and val and "@" not in val:
            errors.append(f"{field['name']}: 올바른 이메일 주소를 입력해주세요.")

    # 값 입력 시 방법 필수 검증
    for group_items in GROUPS.values():
        for item in group_items:
            comp = item["name"]
            method_val = all_data.get(f"{comp}_방법", "")
            any_value  = any(
                all_data.get(f"{comp}_{s}") is not None
                for s in item["samples"]
            )
            if any_value and not method_val:
                errors.append(f"[{comp}] 값을 입력한 경우 방법을 선택해야 합니다.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        inst_field_name = next((f["name"] for f in INFO_FIELDS if "기관" in f["name"]), "기관명")
        email_field_name = next((f["name"] for f in INFO_FIELDS if f["email"]), None)
        inst_name = info_values.get(inst_field_name, "").strip()
        email_to  = info_values.get(email_field_name, "").strip() if email_field_name else ""

        row = {"제출일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        for field in INFO_FIELDS:
            row[field["name"]] = info_values.get(field["name"], "").strip()

        for k, v in all_data.items():
            row[k] = "" if v is None else v

        with st.spinner("제출 중..."):
            try:
                submit_data(row)
                st.success(f"{inst_name or '기관'}의 데이터가 성공적으로 제출되었습니다!")
                st.balloons()
            except Exception as e:
                st.error(f"제출 중 오류가 발생했습니다: {e}")
                st.stop()

        # ── 접수 확인 이메일 즉시 발송 ───────────────────────
        if email_to:
            with st.spinner("접수 확인 이메일 발송 중..."):
                try:
                    send_confirmation(email_to, inst_name, row, cfg)
                    st.info(f"접수 확인 이메일이 {email_to}으로 발송되었습니다.")
                except Exception as e:
                    st.warning(f"이메일 발송 중 오류가 발생했습니다: {e}")
