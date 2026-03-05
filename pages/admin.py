import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from utils.sheets import get_all_data, BASE_FIELDS
from utils.zscore import compute_zscores, zscore_flag, zscore_color
from utils.report import generate_pdf
from utils.email_sender import send_all_reports

st.set_page_config(
    page_title="관리자 페이지",
    page_icon="🔐",
    layout="wide",
)

# ── 관리자 인증 ───────────────────────────────────────────────
if "admin_auth" not in st.session_state:
    st.session_state.admin_auth = False

if not st.session_state.admin_auth:
    st.title("🔐 관리자 로그인")
    pw = st.text_input("비밀번호", type="password")
    if st.button("로그인"):
        if pw == st.secrets["admin"]["password"]:
            st.session_state.admin_auth = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    st.stop()

# ── 메인 ─────────────────────────────────────────────────────
st.title("🔬 숙련도 시험 관리자")
st.caption(f"로그인됨 | {datetime.now().strftime('%Y-%m-%d %H:%M')}")

if st.button("🚪 로그아웃"):
    st.session_state.admin_auth = False
    st.rerun()

st.divider()

# ── 데이터 로드 ───────────────────────────────────────────────
@st.cache_data(ttl=60)
def load_data():
    return get_all_data()

with st.spinner("데이터 불러오는 중..."):
    try:
        df = load_data()
    except Exception as e:
        st.error(f"데이터 로드 실패: {e}")
        st.stop()

if df.empty:
    st.warning("제출된 데이터가 없습니다.")
    st.stop()

st.success(f"총 {len(df)}개 기관 데이터 로드 완료")

# 분석 항목 컬럼 추출 (기본 정보 컬럼 제외, 값이 있는 것만)
analyte_cols = [c for c in df.columns if c not in BASE_FIELDS and c.strip()]

# 숫자형으로 변환
for col in analyte_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# ── 탭 구성 ──────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊 제출 현황", "📈 Z-score 분석", "📧 보고서 발송"])

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab1:
    st.subheader("제출 현황")
    col1, col2, col3 = st.columns(3)
    col1.metric("제출 기관 수", len(df))
    col2.metric("분석 항목 수", len(analyte_cols))
    col3.metric("마지막 제출", df["제출일시"].max() if "제출일시" in df.columns else "-")

    st.dataframe(df[["제출일시", "이메일", "기관명"] + analyte_cols], use_container_width=True)

    if st.button("🔄 데이터 새로고침"):
        st.cache_data.clear()
        st.rerun()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab2:
    st.subheader("Robust Z-score 분석 결과")

    if len(df) < 3:
        st.warning("Z-score 계산을 위해 최소 3개 기관의 데이터가 필요합니다.")
    else:
        zscore_df = compute_zscores(df, analyte_cols)

        # 그룹 통계
        st.markdown("#### 그룹 통계 (중앙값 / MAD)")
        stats_rows = []
        for col in analyte_cols:
            vals = pd.to_numeric(df[col], errors="coerce").dropna()
            if len(vals) < 1:
                continue
            median = np.median(vals)
            mad = np.median(np.abs(vals - median))
            stats_rows.append({
                "항목": col,
                "n": len(vals),
                "중앙값": round(median, 4),
                "MAD": round(mad, 4),
                "최솟값": round(vals.min(), 4),
                "최댓값": round(vals.max(), 4),
            })
        st.dataframe(pd.DataFrame(stats_rows), use_container_width=True)

        st.markdown("#### 기관별 Z-score")
        display_z = zscore_df.copy()
        display_z.insert(0, "기관명", df["기관명"].values)
        display_z.insert(0, "이메일", df["이메일"].values)

        # 색상 스타일 적용
        def color_zscore(val):
            try:
                z = float(val)
                return f"background-color: {zscore_color(z)}"
            except Exception:
                return ""

        styled = display_z.style.applymap(color_zscore, subset=analyte_cols)
        st.dataframe(styled, use_container_width=True)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab3:
    st.subheader("보고서 생성 및 일괄 발송")

    if len(df) < 3:
        st.warning("Z-score 계산을 위해 최소 3개 기관의 데이터가 필요합니다.")
        st.stop()

    zscore_df = compute_zscores(df, analyte_cols)

    # 그룹 통계 dict
    group_stats = {}
    for col in analyte_cols:
        vals = pd.to_numeric(df[col], errors="coerce").dropna()
        if len(vals) < 1:
            continue
        median = float(np.median(vals))
        mad = float(np.median(np.abs(vals - median)))
        group_stats[col] = {"median": median, "mad": mad, "n": len(vals)}

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 개별 PDF 미리보기
    st.markdown("#### 개별 보고서 미리보기 / 다운로드")
    for idx, row in df.iterrows():
        email = row.get("이메일", "")
        institution = row.get("기관명", "")
        row_data = {col: row.get(col, "") for col in analyte_cols}
        zscore_row = {col: zscore_df.loc[idx, col] for col in analyte_cols}

        with st.expander(f"📄 {institution} ({email})"):
            # Z-score 요약 표시
            summary = []
            for col in analyte_cols:
                z = zscore_row.get(col, np.nan)
                try:
                    z_f = float(z)
                    flag = zscore_flag(z_f)
                    summary.append({"항목": col, "제출값": row_data.get(col, "-"),
                                    "Z-score": round(z_f, 2), "판정": flag})
                except Exception:
                    pass
            if summary:
                st.dataframe(pd.DataFrame(summary), use_container_width=True)

            # PDF 다운로드 버튼
            pdf_bytes = generate_pdf(
                email, institution, row_data, zscore_row,
                group_stats, analyte_cols, generated_at
            )
            st.download_button(
                label=f"⬇️ PDF 다운로드",
                data=pdf_bytes,
                file_name=f"report_{institution}.pdf",
                mime="application/pdf",
                key=f"dl_{idx}",
            )

    st.divider()

    # 일괄 발송
    st.markdown("#### 📧 전체 기관 일괄 발송")
    st.warning("발송 후에는 취소할 수 없습니다. 모든 제출 데이터가 확정되었을 때만 실행하세요.")

    if st.button("🚀 전체 보고서 이메일 발송", type="primary"):
        report_list = []
        for idx, row in df.iterrows():
            email = row.get("이메일", "")
            institution = row.get("기관명", "")
            if not email:
                continue
            row_data = {col: row.get(col, "") for col in analyte_cols}
            zscore_row = {col: zscore_df.loc[idx, col] for col in analyte_cols}
            pdf_bytes = generate_pdf(
                email, institution, row_data, zscore_row,
                group_stats, analyte_cols, generated_at
            )
            report_list.append({
                "email": email,
                "institution": institution,
                "pdf_bytes": pdf_bytes,
            })

        with st.spinner(f"{len(report_list)}개 기관에 발송 중..."):
            result = send_all_reports(report_list)

        st.success(f"✅ 발송 완료: {len(result['success'])}개")
        if result["fail"]:
            st.error(f"❌ 발송 실패: {len(result['fail'])}개")
            for f in result["fail"]:
                st.text(f"  - {f['email']}: {f['error']}")
