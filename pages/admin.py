import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

from utils.sheets import get_all_data, BASE_FIELDS, submit_data
from utils.config import (
    get_config, save_config, get_samples, get_component_groups,
    get_nir_groups, get_all_value_columns, get_group_order,
    get_info_fields, get_method_options, get_questions,
    is_value_col, get_component_from_col, get_sample_from_col,
    DEFAULT_CONFIG, CONFIG_COLS,
)
from utils.zscore import (
    compute_zscores, compute_zscores_by_method,
    zscore_flag, zscore_color,
)
from utils.report import generate_pdf
from utils.email_sender import send_all_reports

st.set_page_config(page_title="관리자 페이지", page_icon=None, layout="wide")

# ── 인증 ──────────────────────────────────────────────────────
if "admin_auth" not in st.session_state:
    st.session_state.admin_auth = False

if not st.session_state.admin_auth:
    st.title("관리자 로그인")
    pw = st.text_input("비밀번호", type="password")
    if st.button("로그인"):
        if pw == st.secrets["admin"]["password"]:
            st.session_state.admin_auth = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    st.stop()

st.title("숙련도 시험 관리자")
st.caption(f"로그인됨 | {datetime.now().strftime('%Y-%m-%d %H:%M')}")
if st.button("로그아웃"):
    st.session_state.admin_auth = False
    st.rerun()
st.divider()

# ── 설정 로드 ─────────────────────────────────────────────────
cfg          = get_config()
SAMPLES      = get_samples(cfg)
INFO_FIELDS  = get_info_fields(cfg)
# 기관 정보 필드명 + 제출일시 = 메타 컬럼 (값 컬럼 아님)
META_COLS    = ["제출일시"] + [f["name"] for f in INFO_FIELDS]

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

# 값 컬럼 분류
value_cols = [c for c in df.columns if c not in META_COLS and is_value_col(c, SAMPLES)]
nir_cols   = [c for c in value_cols if c.startswith("NIR_")]
main_cols  = [c for c in value_cols if not c.startswith("NIR_")]

for col in value_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
tab1, tab2, tab3, tab4 = st.tabs([
    "제출 현황", "Z-score 분석", "보고서 발송", "설정"
])

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab1:
    st.subheader("제출 현황")
    if df.empty:
        st.info("제출된 데이터가 없습니다.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("제출 기관 수", len(df))
        c2.metric("분석 항목 수", len(main_cols))
        last = df.get("제출일시", pd.Series()).max()
        c3.metric("마지막 제출", last or "-")
        st.dataframe(df, use_container_width=True)

    if st.button("데이터 새로고침"):
        st.cache_data.clear()
        st.rerun()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab2:
    st.subheader("Robust Z-score 분석")

    if df.empty or len(df) < 3:
        st.warning("Z-score 계산을 위해 최소 3개 기관의 데이터가 필요합니다.")
    else:
        # 성분 그룹 분류
        GROUPS  = get_component_groups(cfg)
        all_comps = [item["name"] for items in GROUPS.values() for item in items]

        def cols_for_comp(comp, include_nir=False):
            prefix = "NIR_" if include_nir else ""
            return [c for c in value_cols if get_component_from_col(c, SAMPLES) == comp
                    and (c.startswith("NIR_") == include_nir)]

        # 분석 소스 선택
        src = st.radio("데이터 소스", ["일반 분석값", "NIR 측정값"], horizontal=True)
        use_nir = src == "NIR 측정값"

        sel_comp = st.selectbox("성분 선택", all_comps)
        comp_cols = cols_for_comp(sel_comp, include_nir=use_nir)

        if not comp_cols:
            st.info("해당 성분의 데이터가 없습니다.")
        else:
            # 그룹 통계
            st.markdown(f"#### 📐 {sel_comp} — 그룹 통계")
            stats_rows = []
            for col in comp_cols:
                vals = df[col].dropna()
                if vals.empty:
                    continue
                med = np.median(vals)
                mad = np.median(np.abs(vals - med))
                stats_rows.append({
                    "사료종류": get_sample_from_col(col, SAMPLES) or "-",
                    "n": len(vals),
                    "중앙값": round(float(med), 4),
                    "MAD":    round(float(mad), 4),
                    "최솟값": round(float(vals.min()), 4),
                    "최댓값": round(float(vals.max()), 4),
                })
            if stats_rows:
                st.dataframe(pd.DataFrame(stats_rows), use_container_width=True)

            # 전체 Z-score
            st.markdown(f"#### {sel_comp} — 기관별 전체 Z-score")
            z_df = compute_zscores(df, comp_cols)

            disp = pd.DataFrame({"기관명": df["기관명"].values})
            z_display_cols = []
            for col in comp_cols:
                s = get_sample_from_col(col, SAMPLES) or col
                disp[f"{s}_값"] = df[col].values
                disp[f"{s}_Z전체"] = z_df[col].round(3).values
                z_display_cols.append(f"{s}_Z전체")

            def color_z(val):
                try:
                    return f"background-color: {zscore_color(float(val))}"
                except Exception:
                    return ""

            st.dataframe(
                disp.style.applymap(color_z, subset=z_display_cols),
                use_container_width=True,
            )

            # 방법별 Z-score (NIR 제외)
            if not use_nir:
                mc = f"{sel_comp}_방법"
                if mc in df.columns:
                    st.markdown(f"#### {sel_comp} — 방법별 Z-score")
                    method_counts = df[mc].value_counts().reset_index()
                    method_counts.columns = ["방법", "기관 수"]
                    st.dataframe(method_counts, use_container_width=True)

                    disp_m = pd.DataFrame({
                        "기관명": df["기관명"].values,
                        "방법":   df[mc].values,
                    })
                    zm_cols = []
                    for col in comp_cols:
                        s = get_sample_from_col(col, SAMPLES) or col
                        zm = compute_zscores_by_method(df, col, mc)
                        disp_m[f"{s}_Z방법별"] = zm.round(3).values
                        zm_cols.append(f"{s}_Z방법별")

                    st.dataframe(
                        disp_m.style.applymap(color_z, subset=zm_cols),
                        use_container_width=True,
                    )
                    st.caption("방법별 Z-score: 동일 방법 사용 기관 3개 미만이면 N/A")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab3:
    st.subheader("보고서 생성 및 일괄 발송")

    if df.empty or len(df) < 3:
        st.warning("최소 3개 기관 데이터가 필요합니다.")
    else:
        z_all    = compute_zscores(df, main_cols)
        z_method = {
            col: compute_zscores_by_method(
                df, col, f"{get_component_from_col(col, SAMPLES)}_방법"
            )
            for col in main_cols
        }
        group_stats = {}
        for col in main_cols:
            vals = df[col].dropna()
            if vals.empty:
                continue
            med = float(np.median(vals))
            mad = float(np.median(np.abs(vals - med)))
            group_stats[col] = {"median": med, "mad": mad, "n": len(vals)}

        generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

        # 기관명/이메일 필드명 동적 추출
        inst_field  = next((f["name"] for f in INFO_FIELDS if "기관" in f["name"]), "기관명")
        email_field = next((f["name"] for f in INFO_FIELDS if f["email"]), "이메일")

        st.markdown("#### 개별 보고서")
        for idx, row in df.iterrows():
            inst     = row.get(inst_field, "")
            email_to = row.get(email_field, "")
            with st.expander(f"{inst} ({email_to})"):
                summary = []
                for col in main_cols:
                    z  = float(z_all.loc[idx, col])
                    zm = z_method[col].loc[idx]
                    summary.append({
                        "성분":     get_component_from_col(col, SAMPLES) or col,
                        "사료종류": get_sample_from_col(col, SAMPLES) or "-",
                        "제출값":   row.get(col, "-"),
                        "Z전체":    round(z, 2) if not np.isnan(z) else "N/A",
                        "Z방법별":  round(float(zm), 2) if not pd.isna(zm) else "N/A",
                        "판정":     zscore_flag(z),
                    })
                st.dataframe(pd.DataFrame(summary), use_container_width=True)

                row_data   = {col: row.get(col, "") for col in main_cols}
                zscore_row = {col: z_all.loc[idx, col] for col in main_cols}
                z_m_row    = {col: z_method[col].loc[idx] for col in main_cols}
                pdf_bytes  = generate_pdf(
                    email_to, inst, row_data, zscore_row, z_m_row,
                    group_stats, main_cols, generated_at,
                )
                st.download_button(
                    "PDF 다운로드", pdf_bytes,
                    f"report_{inst}.pdf", "application/pdf",
                    key=f"dl_{idx}",
                )

        st.divider()
        st.markdown("#### 전체 일괄 발송")
        st.warning("발송 후 취소 불가. 데이터 확정 후 실행하세요.")
        if st.button("전체 보고서 발송", type="primary"):
            report_list = []
            for idx, row in df.iterrows():
                email_to = row.get(email_field, "")
                inst     = row.get(inst_field, "")
                if not email_to:
                    continue
                row_data   = {col: row.get(col, "") for col in main_cols}
                zscore_row = {col: z_all.loc[idx, col] for col in main_cols}
                z_m_row    = {col: z_method[col].loc[idx] for col in main_cols}
                pdf_bytes  = generate_pdf(
                    email_to, inst, row_data, zscore_row, z_m_row,
                    group_stats, main_cols, generated_at,
                )
                report_list.append({
                    "email": email_to, "institution": inst, "pdf_bytes": pdf_bytes,
                })
            with st.spinner(f"{len(report_list)}개 기관 발송 중..."):
                result = send_all_reports(report_list)
            st.success(f"완료: {len(result['success'])}개")
            if result["fail"]:
                st.error(f"실패: {len(result['fail'])}개")
                for f in result["fail"]:
                    st.text(f"  {f['email']}: {f['error']}")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab4:
    st.subheader("설정 관리")
    st.caption("여기서 변경한 내용은 저장 즉시 제출 폼에 반영됩니다.")

    cfg_edit = cfg.copy()

    # ── 기관 정보 필드 ────────────────────────────────────────
    st.markdown("### 기관 정보 필드")
    st.caption(
        "제출 폼 상단의 기관 정보 입력 필드를 관리합니다.\n\n"
        "**placeholder**: 입력창에 표시되는 예시 텍스트 (group 칸에 입력).\n\n"
        "**flags**: `required` = 필수 입력 / `email` = 이메일 형식 검증 / 복수 시 `required,email`."
    )
    info_df = cfg_edit[cfg_edit["type"] == "info_field"][CONFIG_COLS].reset_index(drop=True)
    edited_info = st.data_editor(
        info_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "type":    st.column_config.TextColumn("type", disabled=True, default="info_field"),
            "group":   st.column_config.TextColumn("placeholder"),
            "name":    st.column_config.TextColumn("필드명 *"),
            "samples": st.column_config.TextColumn("flags", help="required / email / required,email"),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
        },
        key="edit_info",
        hide_index=True,
    )
    edited_info["type"] = "info_field"

    st.divider()

    # ── 사료 종류 ─────────────────────────────────────────────
    st.markdown("### 사료 종류")
    st.caption("사료명 추가/삭제/수정. 순서 숫자로 조정.")
    sample_df = cfg_edit[cfg_edit["type"] == "sample"][CONFIG_COLS].reset_index(drop=True)
    edited_samples = st.data_editor(
        sample_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "type":    st.column_config.TextColumn("type", disabled=True, default="sample"),
            "group":   None,
            "name":    st.column_config.TextColumn("사료명 *"),
            "samples": None,
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
        },
        key="edit_samples",
        hide_index=True,
    )
    edited_samples["type"] = "sample"
    edited_samples["group"] = ""
    edited_samples["samples"] = ""

    st.divider()

    # ── 섹션(그룹) 관리 ───────────────────────────────────────
    st.markdown("### 섹션 관리")
    st.caption(
        "섹션 이름 변경, 순서 조정, 추가/삭제, 활성화 여부 설정.\n\n"
        "**NIR 포함**: 체크 시 해당 섹션이 NIR 측정값 테이블에도 나타납니다.\n\n"
        "섹션명을 변경하면 아래 성분 종목의 '그룹' 칸도 동일하게 수정해 주세요."
    )
    group_df = cfg_edit[cfg_edit["type"] == "group"][CONFIG_COLS].reset_index(drop=True)
    group_df["NIR포함"] = group_df["samples"].str.lower().str.contains("nir").fillna(False)

    edited_groups_ui = st.data_editor(
        group_df[["name", "order", "enabled", "NIR포함"]],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "name":    st.column_config.TextColumn("섹션명 *"),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
            "NIR포함": st.column_config.CheckboxColumn("NIR 포함"),
        },
        key="edit_groups",
        hide_index=True,
    )
    # 원본 그룹명 → 새 그룹명 매핑 (rename 추적용)
    old_group_names = group_df["name"].tolist()
    new_group_names = edited_groups_ui["name"].tolist()
    rename_map = {}
    for old, new in zip(old_group_names, new_group_names):
        if old != new and pd.notna(new) and str(new).strip():
            rename_map[old] = str(new).strip()

    # edited_groups를 CONFIG_COLS 형식으로 변환
    edited_groups = edited_groups_ui.copy()
    edited_groups["type"] = "group"
    edited_groups["group"] = ""
    edited_groups["samples"] = edited_groups["NIR포함"].apply(lambda x: "nir" if x else "")
    edited_groups = edited_groups[CONFIG_COLS]

    st.divider()

    # ── 성분 종목 ─────────────────────────────────────────────
    st.markdown("### 성분 종목")
    st.caption(
        "그룹: 위 섹션 관리에서 정의한 섹션명과 일치해야 합니다.\n\n"
        "적용 사료: `all` = 전체 / `축우사료` 또는 `축우사료,양계사료` 처럼 쉼표로 구분."
    )
    comp_df = cfg_edit[cfg_edit["type"] == "component"][CONFIG_COLS].reset_index(drop=True)

    # 현재 유효한 섹션명 목록 (선택용)
    valid_groups = [
        str(n) for n in edited_groups_ui["name"].dropna()
        if str(n).strip()
    ]

    edited_comps = st.data_editor(
        comp_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "type":    st.column_config.TextColumn("type", disabled=True, default="component"),
            "group":   st.column_config.SelectboxColumn(
                "그룹", options=valid_groups,
            ),
            "name":    st.column_config.TextColumn("성분명 *"),
            "samples": st.column_config.TextColumn(
                "적용 사료",
                help="all = 전체 / 축우사료,양계사료 처럼 쉼표로 구분",
            ),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
        },
        key="edit_comps",
        hide_index=True,
    )
    edited_comps["type"] = "component"

    # 섹션명 변경 시 성분의 group 필드 자동 반영
    if rename_map:
        edited_comps["group"] = edited_comps["group"].replace(rename_map)
        st.info(f"섹션 이름 변경 자동 반영: {rename_map}")

    st.divider()

    # ── 방법 옵션 ─────────────────────────────────────────────
    st.markdown("### 방법 옵션")
    st.caption("분석 방법 드롭다운에 표시될 선택지를 관리합니다.")
    method_df = cfg_edit[cfg_edit["type"] == "method_option"][CONFIG_COLS].reset_index(drop=True)
    edited_methods = st.data_editor(
        method_df[["name", "order", "enabled"]],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "name":    st.column_config.TextColumn("방법명 *"),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
        },
        key="edit_methods",
        hide_index=True,
    )
    edited_methods["type"]    = "method_option"
    edited_methods["group"]   = ""
    edited_methods["samples"] = ""
    edited_methods = edited_methods[CONFIG_COLS]

    st.divider()

    # ── 추가 질문 ─────────────────────────────────────────────
    st.markdown("### 추가 질문")
    st.caption(
        "제출 폼 하단에 추가할 질문을 관리합니다.\n\n"
        "**질문 ID**: 데이터 저장 시 컬럼명으로 사용 (영문/숫자/밑줄 권장).\n\n"
        "**유형(samples 칸)**:\n"
        "- `text` 또는 `text:힌트` — 주관식\n"
        "- `choice:옵션1|옵션2|옵션3` — 단일 선택\n"
        "- `multicheck:옵션1|옵션2|옵션3` — 복수 선택"
    )
    question_df = cfg_edit[cfg_edit["type"] == "question"][CONFIG_COLS].reset_index(drop=True)
    edited_questions = st.data_editor(
        question_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "type":    st.column_config.TextColumn("type", disabled=True, default="question"),
            "group":   st.column_config.TextColumn("질문 ID", help="컬럼명으로 사용됨"),
            "name":    st.column_config.TextColumn("질문 내용 *"),
            "samples": st.column_config.TextColumn(
                "유형",
                help="text / text:힌트 / choice:옵션1|옵션2 / multicheck:옵션1|옵션2",
            ),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
        },
        key="edit_questions",
        hide_index=True,
    )
    edited_questions["type"] = "question"

    st.divider()

    # ── 저장 ──────────────────────────────────────────────────
    col_save, col_reset = st.columns([1, 1])
    with col_save:
        if st.button("설정 저장", type="primary", use_container_width=True):
            edited_info      = edited_info[edited_info["name"].astype(str).str.strip() != ""]
            edited_samples   = edited_samples[edited_samples["name"].astype(str).str.strip() != ""]
            edited_groups    = edited_groups[edited_groups["name"].astype(str).str.strip() != ""]
            edited_comps     = edited_comps[edited_comps["name"].astype(str).str.strip() != ""]
            edited_methods   = edited_methods[edited_methods["name"].astype(str).str.strip() != ""]
            edited_questions = edited_questions[edited_questions["name"].astype(str).str.strip() != ""]

            new_cfg = pd.concat(
                [edited_info, edited_samples, edited_groups,
                 edited_comps, edited_methods, edited_questions],
                ignore_index=True,
            )[CONFIG_COLS]

            with st.spinner("저장 중..."):
                try:
                    save_config(new_cfg)
                    st.success("설정이 저장되었습니다. 제출 폼이 즉시 업데이트됩니다.")
                    st.rerun()
                except Exception as e:
                    st.error(f"저장 실패: {e}")

    with col_reset:
        if st.button("기본값으로 초기화", use_container_width=True):
            if st.session_state.get("confirm_reset"):
                with st.spinner("초기화 중..."):
                    save_config(DEFAULT_CONFIG.copy())
                    st.success("기본값으로 초기화되었습니다.")
                    st.session_state.confirm_reset = False
                    st.rerun()
            else:
                st.session_state.confirm_reset = True
                st.warning("한 번 더 클릭하면 초기화됩니다.")
