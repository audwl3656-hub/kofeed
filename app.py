import streamlit as st
from datetime import datetime
from utils.sheets import (
    submit_data,
    SAMPLES, PROXIMATE, CATTLE_ONLY, AMINO_ACIDS, NIR_COMPONENTS,
)

st.set_page_config(
    page_title="사료 숙련도 시험 데이터 제출",
    page_icon="🧪",
    layout="wide",
)

st.title("🧪 사료 숙련도 시험")
st.caption("제출하신 데이터는 Robust Z-score 분석 후 보고서로 발송됩니다.")
st.divider()

# ── 기관 정보 ─────────────────────────────────────────────────
st.subheader("🏢 기관 정보")
c1, c2, c3, c4 = st.columns(4)
with c1:
    institution = st.text_input("기관명 *", placeholder="○○ 연구소")
with c2:
    person = st.text_input("담당자명 *", placeholder="홍길동")
with c3:
    email = st.text_input("이메일 *", placeholder="lab@example.com")
with c4:
    phone = st.text_input("전화번호", placeholder="010-0000-0000")

st.divider()


# ── 성분 입력 테이블 헬퍼 ─────────────────────────────────────
def component_table(components: list, samples: list, prefix: str = "") -> dict:
    """
    성분 × (방법/기기/용매 + 사료별 값) 입력 테이블.
    반환: {field_name: value}
    """
    data = {}
    col_ratios = [2, 2, 2, 2] + [2] * len(samples)

    # 헤더 행
    hcols = st.columns(col_ratios)
    hcols[0].markdown("**성분**")
    hcols[1].markdown("**방법**")
    hcols[2].markdown("**기기명**")
    hcols[3].markdown("**용매**")
    for i, s in enumerate(samples):
        hcols[4 + i].markdown(f"**{s}**")

    st.markdown(
        "<hr style='margin:2px 0 6px 0; border-color:#e0e0e0'>",
        unsafe_allow_html=True,
    )

    for comp in components:
        cols = st.columns(col_ratios)
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
        for i, sample in enumerate(samples):
            with cols[4 + i]:
                val = st.number_input(
                    sample, min_value=0.0, value=None,
                    step=0.0001, format="%.4f", placeholder="미입력",
                    key=f"{prefix}{comp}_{sample}",
                    label_visibility="collapsed",
                )
                data[f"{comp}_{sample}"] = val

    return data


def nir_table(components: list, samples: list) -> dict:
    """NIR 측정값 입력 테이블 (기기명 + 사료별 값)"""
    data = {}
    col_ratios = [2, 3] + [2] * len(samples)

    hcols = st.columns(col_ratios)
    hcols[0].markdown("**성분**")
    hcols[1].markdown("**기기명**")
    for i, s in enumerate(samples):
        hcols[2 + i].markdown(f"**{s}**")

    st.markdown(
        "<hr style='margin:2px 0 6px 0; border-color:#e0e0e0'>",
        unsafe_allow_html=True,
    )

    for comp in components:
        cols = st.columns(col_ratios)
        cols[0].markdown(f"**{comp}**")

        with cols[1]:
            data[f"NIR_{comp}_기기"] = st.text_input(
                "기기", key=f"NIR_{comp}_equip",
                label_visibility="collapsed", placeholder="기기명",
            )
        for i, sample in enumerate(samples):
            with cols[2 + i]:
                val = st.number_input(
                    sample, min_value=0.0, value=None,
                    step=0.0001, format="%.4f", placeholder="미입력",
                    key=f"NIR_{comp}_{sample}",
                    label_visibility="collapsed",
                )
                data[f"NIR_{comp}_{sample}"] = val

    return data


# ── 일반성분 ─────────────────────────────────────────────────
st.subheader("📊 일반성분 (g/kg 건물 기준)")
prox_data = component_table(PROXIMATE, SAMPLES)
st.divider()

# ── ADF / NDF ─────────────────────────────────────────────────
st.subheader("📊 ADF / NDF — 축우사료 전용 (g/kg 건물 기준)")
cattle_data = component_table(CATTLE_ONLY, ["축우사료"], prefix="CF_")
st.divider()

# ── 아미노산 ─────────────────────────────────────────────────
st.subheader("🔬 아미노산 (g/kg 건물 기준)")
aa_data = component_table(AMINO_ACIDS, SAMPLES, prefix="AA_")
st.divider()

# ── NIR 측정값 ───────────────────────────────────────────────
st.subheader("📡 NIR 측정값")
st.caption("NIR 기기로 측정한 일반성분 및 ADF/NDF 값을 입력하세요.")

# NIR — ADF/NDF는 축우사료만
nir_data = {}
nir_proximate_data = nir_table(PROXIMATE, SAMPLES)
nir_data.update(nir_proximate_data)

st.markdown("**ADF / NDF (축우사료 전용)**")
nir_cattle_data = nir_table(CATTLE_ONLY, ["축우사료"])
nir_data.update(nir_cattle_data)

st.divider()

# ── 제출 ─────────────────────────────────────────────────────
if st.button("📤 데이터 제출", type="primary", use_container_width=True):
    errors = []
    if not institution.strip():
        errors.append("기관명을 입력해주세요.")
    if not person.strip():
        errors.append("담당자명을 입력해주세요.")
    if not email.strip() or "@" not in email:
        errors.append("올바른 이메일 주소를 입력해주세요.")

    if errors:
        for e in errors:
            st.error(e)
    else:
        row = {
            "제출일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "기관명":   institution.strip(),
            "담당자명": person.strip(),
            "이메일":   email.strip(),
            "전화":     phone.strip(),
        }

        def add_data(d: dict):
            for k, v in d.items():
                if v is None or v == "":
                    row[k] = ""
                elif isinstance(v, float):
                    row[k] = round(v, 4)
                else:
                    row[k] = v

        add_data(prox_data)
        add_data(cattle_data)
        add_data(aa_data)
        add_data(nir_data)

        with st.spinner("제출 중..."):
            try:
                submit_data(row)
                st.success(f"✅ {institution}의 데이터가 성공적으로 제출되었습니다!")
                st.info("분석 완료 후 보고서를 이메일로 발송해 드립니다.")
                st.balloons()
            except Exception as e:
                st.error(f"제출 중 오류가 발생했습니다: {e}")
