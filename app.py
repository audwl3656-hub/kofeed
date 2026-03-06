import streamlit as st
from datetime import datetime
from utils.sheets import submit_data
from utils.config import get_config, get_samples, get_component_groups, get_nir_groups, get_info_fields

st.set_page_config(
    page_title="사료 숙련도 시험 데이터 제출",
    page_icon=None,
    layout="wide",
)

st.title("사료 숙련도 시험")
st.caption("제출하신 데이터는 Robust Z-score 분석 후 보고서로 발송됩니다.")
st.divider()

# ── 설정 로드 ─────────────────────────────────────────────────
cfg         = get_config()
SAMPLES     = get_samples(cfg)
GROUPS      = get_component_groups(cfg)
NIR_GRP     = get_nir_groups(cfg)
INFO_FIELDS = get_info_fields(cfg)

# ── 기관 정보 ─────────────────────────────────────────────────
st.subheader("기관 정보")
info_cols = st.columns(max(len(INFO_FIELDS), 1))
info_values: dict[str, str] = {}
for i, field in enumerate(INFO_FIELDS):
    with info_cols[i]:
        label = field["name"] + (" *" if field["required"] else "")
        info_values[field["name"]] = st.text_input(
            label, placeholder=field["placeholder"], key=f"info_{field['name']}"
        )

st.divider()


# ── 성분 입력 테이블 헬퍼 ─────────────────────────────────────
def component_table(items: list[dict], prefix: str = "") -> dict:
    """
    items: [{"name": str, "samples": [str]}, ...]
    모든 사료가 같으면 하나의 통합 테이블, 아니면 각 항목별 칼럼 다름.
    """
    data = {}

    # 이 그룹에서 실제 사용되는 사료 종류 (순서 유지)
    all_sample_set = []
    for item in items:
        for s in item["samples"]:
            if s not in all_sample_set:
                all_sample_set.append(s)

    col_ratios = [2, 2, 2, 2] + [2] * len(all_sample_set)

    # 헤더
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
        cols    = st.columns(col_ratios)
        cols[0].markdown(f"**{comp}**")

        with cols[1]:
            data[f"{comp}_방법"] = st.text_input(
                "방법", key=f"{prefix}{comp}_method",
                label_visibility="collapsed", placeholder="예: AOAC 950.01",
            )
        with cols[2]:
            data[f"{comp}_기기"] = st.text_input(
                "기기", key=f"{prefix}{comp}_equip",
                label_visibility="collapsed", placeholder="기기명",
            )
        with cols[3]:
            data[f"{comp}_용매"] = st.text_input(
                "용매", key=f"{prefix}{comp}_solvent",
                label_visibility="collapsed", placeholder="용매",
            )
        for i, s in enumerate(all_sample_set):
            with cols[4 + i]:
                if s in samples:
                    val = st.number_input(
                        s, min_value=0.0, value=None,
                        step=0.0001, format="%.4f", placeholder="미입력",
                        key=f"{prefix}{comp}_{s}",
                        label_visibility="collapsed",
                    )
                    data[f"{comp}_{s}"] = val
                else:
                    st.markdown("<div style='color:#ccc;text-align:center'>—</div>",
                                unsafe_allow_html=True)
    return data


def nir_table(items: list[dict]) -> dict:
    """NIR 기기명 + 사료별 값"""
    data = {}

    all_sample_set = []
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
                    val = st.number_input(
                        s, min_value=0.0, value=None,
                        step=0.0001, format="%.4f", placeholder="미입력",
                        key=f"NIR_{comp}_{s}",
                        label_visibility="collapsed",
                    )
                    data[f"NIR_{comp}_{s}"] = val
                else:
                    st.markdown("<div style='color:#ccc;text-align:center'>—</div>",
                                unsafe_allow_html=True)
    return data


# ── 성분 그룹별 폼 렌더링 ─────────────────────────────────────
all_data: dict = {}

for group_name, items in GROUPS.items():
    st.subheader(group_name)
    group_data = component_table(items, prefix=f"{group_name}_")
    all_data.update(group_data)
    st.divider()

# ── NIR 측정값 ───────────────────────────────────────────────
if NIR_GRP:
    st.subheader("NIR 측정값")
    st.caption("NIR 기기로 측정한 값을 입력하세요.")
    for group_name, items in NIR_GRP.items():
        st.markdown(f"**{group_name}**")
        nir_data = nir_table(items)
        all_data.update(nir_data)
    st.divider()

# ── 제출 ─────────────────────────────────────────────────────
if st.button("데이터 제출", type="primary", use_container_width=True):
    errors = []
    for field in INFO_FIELDS:
        val = info_values.get(field["name"], "").strip()
        if field["required"] and not val:
            errors.append(f"{field['name']}을(를) 입력해주세요.")
        elif field["email"] and val and "@" not in val:
            errors.append(f"{field['name']}: 올바른 이메일 주소를 입력해주세요.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        # 기관명, 이메일 추출 (보고서 발송용)
        inst_name  = info_values.get("기관명", "").strip()
        inst_email = info_values.get("이메일", "").strip()

        row = {"제출일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        for field in INFO_FIELDS:
            row[field["name"]] = info_values.get(field["name"], "").strip()

        for k, v in all_data.items():
            if v is None or v == "":
                row[k] = ""
            elif isinstance(v, float):
                row[k] = round(v, 4)
            else:
                row[k] = v

        with st.spinner("제출 중..."):
            try:
                submit_data(row)
                st.success(f"{inst_name or '기관'}의 데이터가 성공적으로 제출되었습니다!")
                st.info("분석 완료 후 보고서를 이메일로 발송해 드립니다.")
                st.balloons()
            except Exception as e:
                st.error(f"제출 중 오류가 발생했습니다: {e}")
