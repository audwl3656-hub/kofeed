import streamlit as st
from datetime import datetime
from utils.sheets import submit_data, get_custom_fields

st.set_page_config(
    page_title="아미노산 숙련도 시험 데이터 제출",
    page_icon="🧪",
    layout="centered",
)

# 기본 아미노산 항목
BASE_AMINO_ACIDS = [
    "ASP", "THR", "SER", "GLU", "GLY",
    "ALA", "VAL", "ISOL", "LEU", "TYR",
    "PHE", "LYS", "HIS", "ARG", "PRO",
    "MET", "CYS",
]

st.title("🧪 아미노산 숙련도 시험")
st.caption("제출하신 데이터는 Robust Z-score 분석 후 보고서로 발송됩니다.")
st.divider()

# ── 기관 정보 ────────────────────────────────────────────────
st.subheader("🏢 기관 정보")
col1, col2 = st.columns(2)
with col1:
    email = st.text_input("이메일 *", placeholder="lab@example.com")
with col2:
    institution = st.text_input("기관명 *", placeholder="○○ 연구소")

st.divider()

# ── 아미노산 데이터 입력 ─────────────────────────────────────
st.subheader("🔬 아미노산 분석값 입력")
st.caption("단위: g/kg (건물 기준) — 미분석 항목은 0으로 입력하거나 비워두세요.")

aa_data = {}
cols = st.columns(3)
for i, aa in enumerate(BASE_AMINO_ACIDS):
    with cols[i % 3]:
        val = st.number_input(
            aa, min_value=0.0, value=None,
            step=0.0001, format="%.4f",
            placeholder="미입력",
            key=f"aa_{aa}",
        )
        aa_data[aa] = val

st.divider()

# ── 추가 항목 ─────────────────────────────────────────────────
st.subheader("➕ 추가 항목 (선택)")
st.caption("기관에서 추가로 분석한 항목을 입력할 수 있습니다. 여기서 추가한 항목은 이후 제출 양식에도 표시됩니다.")

# 기존 커스텀 필드 불러오기
try:
    existing_custom = get_custom_fields()
except Exception:
    existing_custom = []

custom_data = {}

if existing_custom:
    c_cols = st.columns(3)
    for i, field in enumerate(existing_custom):
        with c_cols[i % 3]:
            val = st.number_input(
                field, min_value=0.0, value=None,
                step=0.0001, format="%.4f",
                placeholder="미입력",
                key=f"custom_{field}",
            )
            custom_data[field] = val

with st.expander("🆕 새 항목 추가하기"):
    st.info("새로 추가한 항목은 향후 모든 제출 양식에도 표시됩니다.")
    new_fields = []
    for i in range(3):
        nc1, nc2 = st.columns([2, 1])
        with nc1:
            fname = st.text_input(f"항목명 {i+1}", key=f"nf_{i}", placeholder="예: HYP")
        with nc2:
            fval = st.number_input(
                "값", min_value=0.0, value=None,
                step=0.0001, format="%.4f",
                key=f"nv_{i}", placeholder="미입력",
            )
        if fname.strip():
            new_fields.append((fname.strip().upper(), fval))

st.divider()

# ── 제출 ─────────────────────────────────────────────────────
if st.button("📤 데이터 제출", type="primary", use_container_width=True):
    if not email.strip() or "@" not in email:
        st.error("올바른 이메일 주소를 입력해주세요.")
    elif not institution.strip():
        st.error("기관명을 입력해주세요.")
    else:
        row = {
            "제출일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "이메일":   email.strip(),
            "기관명":   institution.strip(),
        }
        # 아미노산 값 (None이면 빈 문자열로)
        for aa, val in aa_data.items():
            row[aa] = "" if val is None else round(val, 4)

        # 기존 커스텀 필드
        for field, val in custom_data.items():
            row[field] = "" if val is None else round(val, 4)

        # 새 커스텀 필드
        for fname, fval in new_fields:
            row[fname] = "" if fval is None else round(fval, 4)

        with st.spinner("제출 중..."):
            try:
                submit_data(row)
                st.success(f"✅ {institution}의 데이터가 성공적으로 제출되었습니다!")
                st.info("분석 완료 후 보고서를 이메일로 발송해 드립니다.")
                st.balloons()
            except Exception as e:
                st.error(f"제출 중 오류가 발생했습니다: {e}")
