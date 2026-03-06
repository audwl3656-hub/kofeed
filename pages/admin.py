import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

from utils.sheets import (
    get_all_data, BASE_FIELDS,
    SAMPLES, PROXIMATE, CATTLE_ONLY, AMINO_ACIDS, NIR_COMPONENTS,
    is_value_col, is_nir_col, get_component, get_sample, method_col,
)
from utils.zscore import (
    compute_zscores, compute_zscores_by_method, zscore_flag, zscore_color,
    robust_zscore,
)
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

# ── 메인 ──────────────────────────────────────────────────────
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

# ── 컬럼 분류 ─────────────────────────────────────────────────
all_cols = [c for c in df.columns if c not in BASE_FIELDS]

# 값 컬럼 (일반 + NIR)
value_cols     = [c for c in all_cols if is_value_col(c) and not is_nir_col(c)]
nir_value_cols = [c for c in all_cols if is_nir_col(c)]

# 숫자 변환
for col in value_cols + nir_value_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# ── 성분 그룹 분류 ────────────────────────────────────────────
def classify_cols(cols):
    """컬럼을 성분 그룹별로 분류: {group_label: [cols]}"""
    groups = {
        "일반성분": [],
        "ADF/NDF":  [],
        "아미노산": [],
    }
    for c in cols:
        comp = get_component(c)
        if comp in PROXIMATE:
            groups["일반성분"].append(c)
        elif comp in CATTLE_ONLY:
            groups["ADF/NDF"].append(c)
        elif comp in AMINO_ACIDS:
            groups["아미노산"].append(c)
    return groups


groups        = classify_cols(value_cols)
nir_groups    = {"NIR 일반성분": nir_value_cols}

# ── 탭 구성 ──────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊 제출 현황", "📈 Z-score 분석", "📧 보고서 발송"])

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab1:
    st.subheader("제출 현황")
    c1, c2, c3 = st.columns(3)
    c1.metric("제출 기관 수", len(df))
    c2.metric("분석 항목 수", len(value_cols))
    c3.metric("마지막 제출", df["제출일시"].max() if "제출일시" in df.columns else "-")

    st.dataframe(df, use_container_width=True)

    if st.button("🔄 데이터 새로고침"):
        st.cache_data.clear()
        st.rerun()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab2:
    st.subheader("Robust Z-score 분석")

    if len(df) < 3:
        st.warning("Z-score 계산을 위해 최소 3개 기관의 데이터가 필요합니다.")
    else:
        # ── 표시 옵션 ─────────────────────────────────────────
        all_group_labels = list(groups.keys()) + list(nir_groups.keys())
        sel_group = st.selectbox("성분 그룹 선택", all_group_labels)

        if sel_group in groups:
            target_cols = groups[sel_group]
        else:
            target_cols = nir_groups[sel_group]

        if not target_cols:
            st.info("해당 그룹에 제출된 데이터가 없습니다.")
        else:
            # 성분 목록 (중복 제거)
            comps_in_group = list(dict.fromkeys(
                get_component(c) for c in target_cols
            ))
            sel_comp = st.selectbox("성분 선택", comps_in_group)

            comp_cols = [c for c in target_cols if get_component(c) == sel_comp]

            # ── 그룹 통계 ─────────────────────────────────────
            st.markdown(f"#### 📐 {sel_comp} — 그룹 통계")
            stats_rows = []
            for col in comp_cols:
                vals = df[col].dropna()
                if len(vals) < 1:
                    continue
                med = np.median(vals)
                mad = np.median(np.abs(vals - med))
                stats_rows.append({
                    "컬럼": col,
                    "사료종류": get_sample(col) or "-",
                    "n": len(vals),
                    "중앙값": round(float(med), 4),
                    "MAD": round(float(mad), 4),
                    "최솟값": round(float(vals.min()), 4),
                    "최댓값": round(float(vals.max()), 4),
                })
            if stats_rows:
                st.dataframe(pd.DataFrame(stats_rows), use_container_width=True)

            # ── 전체 Z-score ──────────────────────────────────
            st.markdown(f"#### 📊 {sel_comp} — 기관별 전체 Z-score")
            z_df = compute_zscores(df, comp_cols)

            disp = pd.DataFrame({
                "기관명": df["기관명"].values,
            })
            for col in comp_cols:
                sample = get_sample(col) or col
                disp[f"{sample}_값"] = df[col].values
                disp[f"{sample}_Z전체"] = z_df[col].values.round(3)

            def color_z(val):
                try:
                    return f"background-color: {zscore_color(float(val))}"
                except Exception:
                    return ""

            z_cols = [c for c in disp.columns if c.endswith("_Z전체")]
            styled = disp.style.applymap(color_z, subset=z_cols)
            st.dataframe(styled, use_container_width=True)

            # ── 방법별 Z-score ────────────────────────────────
            # NIR 컬럼은 방법 컬럼이 없으므로 방법별 분석 생략
            if not sel_group.startswith("NIR"):
                m_col = method_col(sel_comp)
                if m_col in df.columns:
                    st.markdown(f"#### 🔬 {sel_comp} — 방법별 Z-score")

                    # 방법 분포
                    method_counts = df[m_col].value_counts().reset_index()
                    method_counts.columns = ["방법", "기관 수"]
                    st.dataframe(method_counts, use_container_width=True)

                    disp_m = pd.DataFrame({
                        "기관명": df["기관명"].values,
                        "방법":   df[m_col].values,
                    })
                    z_method_cols = []
                    for col in comp_cols:
                        sample = get_sample(col) or col
                        z_m = compute_zscores_by_method(df, col, m_col)
                        col_name = f"{sample}_Z방법별"
                        disp_m[col_name] = z_m.values.round(3)
                        z_method_cols.append(col_name)

                    styled_m = disp_m.style.applymap(color_z, subset=z_method_cols)
                    st.dataframe(styled_m, use_container_width=True)
                    st.caption(
                        "방법별 Z-score는 동일 방법 사용 기관이 3개 미만이면 N/A로 표시됩니다."
                    )


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab3:
    st.subheader("보고서 생성 및 일괄 발송")

    if len(df) < 3:
        st.warning("Z-score 계산을 위해 최소 3개 기관의 데이터가 필요합니다.")
        st.stop()

    # 전체 그룹 통계 계산 (값 컬럼 전체)
    z_all = compute_zscores(df, value_cols)

    # 방법별 Z-score 사전 계산
    z_method_all: dict[str, pd.Series] = {}
    for vcol in value_cols:
        comp = get_component(vcol)
        if comp:
            mc = method_col(comp)
            z_method_all[vcol] = compute_zscores_by_method(df, vcol, mc)

    group_stats: dict[str, dict] = {}
    for col in value_cols:
        vals = df[col].dropna()
        if len(vals) < 1:
            continue
        med = float(np.median(vals))
        mad = float(np.median(np.abs(vals - med)))
        group_stats[col] = {"median": med, "mad": mad, "n": len(vals)}

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── 개별 PDF 미리보기 ─────────────────────────────────────
    st.markdown("#### 개별 보고서 미리보기 / 다운로드")
    for idx, row in df.iterrows():
        email_to    = row.get("이메일", "")
        institution = row.get("기관명", "")

        row_data   = {col: row.get(col, "") for col in value_cols}
        zscore_row = {col: z_all.loc[idx, col] for col in value_cols}
        z_m_row    = {col: z_method_all[col].loc[idx] for col in value_cols}

        with st.expander(f"📄 {institution} ({email_to})"):
            summary = []
            for col in value_cols:
                z  = zscore_row.get(col, np.nan)
                zm = z_m_row.get(col, np.nan)
                try:
                    z_f  = float(z)
                    zm_f = float(zm) if not pd.isna(zm) else np.nan
                    summary.append({
                        "컬럼":       col,
                        "성분":       get_component(col) or col,
                        "사료종류":   get_sample(col) or "-",
                        "제출값":     row_data.get(col, "-"),
                        "Z전체":      round(z_f, 2),
                        "Z방법별":    round(zm_f, 2) if not np.isnan(zm_f) else "N/A",
                        "판정":       zscore_flag(z_f),
                    })
                except Exception:
                    pass
            if summary:
                st.dataframe(pd.DataFrame(summary), use_container_width=True)

            pdf_bytes = generate_pdf(
                email_to, institution, row_data, zscore_row, z_m_row,
                group_stats, value_cols, generated_at,
            )
            st.download_button(
                label="⬇️ PDF 다운로드",
                data=pdf_bytes,
                file_name=f"report_{institution}.pdf",
                mime="application/pdf",
                key=f"dl_{idx}",
            )

    st.divider()

    # ── 일괄 발송 ─────────────────────────────────────────────
    st.markdown("#### 📧 전체 기관 일괄 발송")
    st.warning("발송 후에는 취소할 수 없습니다. 모든 제출 데이터가 확정되었을 때만 실행하세요.")

    if st.button("🚀 전체 보고서 이메일 발송", type="primary"):
        report_list = []
        for idx, row in df.iterrows():
            email_to    = row.get("이메일", "")
            institution = row.get("기관명", "")
            if not email_to:
                continue
            row_data   = {col: row.get(col, "") for col in value_cols}
            zscore_row = {col: z_all.loc[idx, col] for col in value_cols}
            z_m_row    = {col: z_method_all[col].loc[idx] for col in value_cols}
            pdf_bytes  = generate_pdf(
                email_to, institution, row_data, zscore_row, z_m_row,
                group_stats, value_cols, generated_at,
            )
            report_list.append({
                "email": email_to, "institution": institution, "pdf_bytes": pdf_bytes,
            })

        with st.spinner(f"{len(report_list)}개 기관에 발송 중..."):
            result = send_all_reports(report_list)

        st.success(f"✅ 발송 완료: {len(result['success'])}개")
        if result["fail"]:
            st.error(f"❌ 발송 실패: {len(result['fail'])}개")
            for f in result["fail"]:
                st.text(f"  - {f['email']}: {f['error']}")
