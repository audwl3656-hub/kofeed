import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))

from utils.sheets import get_all_data, BASE_FIELDS, submit_data
from utils.config import (
    get_config, save_config, get_samples, get_component_groups,
    get_all_value_columns, get_group_order,
    get_info_fields, get_method_options, get_questions,
    is_value_col, get_component_from_col, get_sample_from_col,
    get_col_suffix, get_base_col,
    get_participant_map,
    DEFAULT_CONFIG, CONFIG_COLS,
)
from utils.zscore import (
    compute_zscores, compute_zscores_by_method, compute_zscores_by_method_multi,
    zscore_flag, zscore_color,
)
from utils.report import generate_pdf_overall, generate_pdf_by_method, generate_pdf_summary
from utils.email_sender import send_all_reports
from utils.config import get_history, append_history_rows, delete_history_rows
from utils.history_dashboard import generate_institution_html_bytes, generate_institution_email_html

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
st.caption(f"로그인됨 | {datetime.now(KST).strftime('%Y-%m-%d %H:%M')}")
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
value_cols = [c for c in df.columns if c not in META_COLS and is_value_col(c, SAMPLES)
              and not c.startswith("NIR_")]
main_cols  = value_cols

# 섹션/성분 설정 순서대로 정렬
_ordered = get_all_value_columns(cfg)
_main_set = set(main_cols)
main_cols = [c for c in _ordered if c in _main_set] + [c for c in main_cols if c not in set(_ordered)]

for col in value_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# 기관명 필드 (공통 사용)
inst_field  = next((f["name"] for f in INFO_FIELDS if any(kw in f["name"] for kw in ("기관", "회사", "기업", "업체"))), INFO_FIELDS[0]["name"] if INFO_FIELDS else "기관명")
email_field = next((f["name"] for f in INFO_FIELDS if "email" in f.get("samples", "").lower()), "이메일")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "제출 현황", "Z-score 분석", "보고서 발송", "설정", "연도별 히스토리"
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

        def cols_for_comp(comp):
            return [c for c in value_cols if get_component_from_col(c, SAMPLES) == comp]

        sel_comp = st.selectbox("성분 선택", all_comps)
        comp_cols = cols_for_comp(sel_comp)

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
                med  = float(np.median(vals))
                mean = float(vals.mean())
                std  = float(vals.std())
                cv   = round(std / mean * 100, 2) if mean != 0 else np.nan
                stats_rows.append({
                    "사료종류": get_sample_from_col(col, SAMPLES) or "-",
                    "n":      len(vals),
                    "평균":   round(mean, 4),
                    "중앙값": round(med, 4),
                    "CV(%)":  cv if not np.isnan(cv) else "N/A",
                    "최솟값": round(float(vals.min()), 4),
                    "최댓값": round(float(vals.max()), 4),
                })
            if stats_rows:
                st.dataframe(pd.DataFrame(stats_rows), use_container_width=True)

            # 전체 Z-score
            st.markdown(f"#### {sel_comp} — 기관별 전체 Z-score")
            z_df = compute_zscores(df, comp_cols)

            inst_col = next((f["name"] for f in INFO_FIELDS if any(kw in f["name"] for kw in ("기관", "회사", "기업", "업체"))), META_COLS[1] if len(META_COLS) > 1 else META_COLS[0])
            disp = pd.DataFrame({"기관명": df[inst_col].values if inst_col in df.columns else ""})
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
                disp.style.map(color_z, subset=z_display_cols),
                use_container_width=True,
            )

            # 방법별 Z-score
            mc = f"{sel_comp}_방법"
            if mc in df.columns:
                st.markdown(f"#### {sel_comp} — 방법별 Z-score")
                method_counts = df[mc].value_counts().reset_index()
                method_counts.columns = ["방법", "기관 수"]
                st.dataframe(method_counts, use_container_width=True)

                disp_m = pd.DataFrame({
                    "기관명": df[inst_col].values if inst_col in df.columns else "",
                    "방법":   df[mc].values,
                })
                zm_cols = []
                for col in comp_cols:
                    s = get_sample_from_col(col, SAMPLES) or col
                    zm = compute_zscores_by_method(df, col, mc)
                    disp_m[f"{s}_Z방법별"] = zm.round(3).values
                    zm_cols.append(f"{s}_Z방법별")

                st.dataframe(
                    disp_m.style.map(color_z, subset=zm_cols),
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

        # 방법별 Z-score: 동일 base 컬럼(_2, _3 포함)을 풀링하여 계산
        _base_to_vcols: dict = {}
        for col in main_cols:
            _base_to_vcols.setdefault(get_base_col(col), []).append(col)
        z_method: dict = {}
        for _cols in _base_to_vcols.values():
            _comp = get_component_from_col(_cols[0], SAMPLES)
            _mcols = [f"{_comp}_방법{get_col_suffix(c)}" for c in _cols]
            z_method.update(compute_zscores_by_method_multi(df, _cols, _mcols))
        # 전체 Z-score용 통계: 동일 base col(_2, _3 포함) 풀링하여 계산
        _base_pool: dict = {}
        for col in main_cols:
            base = get_base_col(col)
            vals = pd.to_numeric(df[col], errors="coerce").dropna()
            _base_pool.setdefault(base, []).extend(vals.tolist())

        group_stats = {}
        for col in main_cols:
            base = get_base_col(col)
            pool = np.array(_base_pool.get(base, []))
            if len(pool) == 0:
                continue
            med  = float(np.median(pool))
            q1, q3 = float(np.percentile(pool, 25)), float(np.percentile(pool, 75))
            niqr = (q3 - q1) * 0.7413
            mean = float(np.mean(pool))
            std  = float(np.std(pool, ddof=1)) if len(pool) > 1 else float("nan")
            cv   = (std / mean * 100) if mean != 0 and not np.isnan(std) else float("nan")
            group_stats[col] = {"median": med, "niqr": niqr, "n": len(pool),
                                "mean": mean, "std": std, "cv": cv}

        def _calc_method_group_stats(inst_method_dict):
            """기관의 방법을 기준으로 동일 방법 기관들만의 통계를 반환."""
            stats = {}
            for col in main_cols:
                comp = get_component_from_col(col, SAMPLES)
                sfx  = get_col_suffix(col)
                mc   = f"{comp}_방법{sfx}" if comp else None
                method = inst_method_dict.get(f"{comp}{sfx}", "").strip() if comp else ""
                if mc and mc in df.columns and method:
                    mask = df[mc].fillna("").astype(str).str.strip() == method
                    vals = df.loc[mask, col].dropna()
                    if not vals.empty:
                        med  = float(np.median(vals))
                        mean = float(vals.mean())
                        std  = float(vals.std())
                        cv   = (std / mean * 100) if mean != 0 else np.nan
                        stats[col] = {"median": med, "n": len(vals),
                                      "mean": mean, "std": std, "cv": cv}
                        continue
                stats[col] = {}
            return stats

        generated_at = datetime.now(KST).strftime("%Y-%m-%d %H:%M")

        # 기관명/이메일 필드명 동적 추출
        inst_field  = next((f["name"] for f in INFO_FIELDS if any(kw in f["name"] for kw in ("기관", "회사", "기업", "업체"))), INFO_FIELDS[0]["name"] if INFO_FIELDS else "기관명")
        email_field = next((f["name"] for f in INFO_FIELDS if f["email"]), "이메일")

        # ── 전체 요약 보고서
        st.markdown("#### 전체 요약 보고서")
        _c1, _c2 = st.columns(2)
        with _c1:
            _rpt_subtitle = st.text_input("보고서 부제", placeholder="예: 2025년 2차", key="rpt_subtitle")
            _rpt_p1 = st.text_input("① 시료 배부 기간", placeholder="예: 2025년 10월 7일 ~ 10월 10일", key="rpt_p1")
        with _c2:
            _rpt_p2 = st.text_input("② 분석 및 결과회신 기간", placeholder="예: 2025년 10월 10일 ~ 10월 31일", key="rpt_p2")
            _rpt_p3 = st.text_input("③ 결과처리 및 보고서 작성 기간", placeholder="예: 2025년 11월 1일 ~ 11월 6일", key="rpt_p3")
        _rpt_note = st.text_input("시료 주석 (선택)", placeholder="예: 시료분쇄 : 1.0 mm 입자", key="rpt_note")
        _rpt_summary = st.text_area(
            "나. 분석결과 요약 내용 (줄바꿈으로 구분)",
            placeholder="- 전반적으로 분석 결과가 양호하였음.\n- 조지방(산분해) 항목에서 일부 기관의 편차가 크게 나타남.",
            key="rpt_summary", height=120,
        )

        _pdf_kwargs = dict(
            df=df, z_all=z_all, z_method=z_method,
            group_stats=group_stats, value_cols=main_cols,
            inst_field=inst_field, generated_at=generated_at, samples=SAMPLES,
            participant_map=get_participant_map(cfg),
            subtitle=_rpt_subtitle,
            period_배부=_rpt_p1, period_회신=_rpt_p2, period_보고서=_rpt_p3,
            sample_note=_rpt_note, summary_text=_rpt_summary, cfg=cfg,
        )

        if st.button("📄 보고서 생성 및 미리보기", key="gen_summary", type="primary"):
            with st.spinner("보고서 생성 중..."):
                st.session_state["summary_pdf"] = generate_pdf_summary(**_pdf_kwargs)

        summary_pdf = st.session_state.get("summary_pdf")
        if summary_pdf:
            # ── 다운로드 버튼 ──
            st.download_button(
                "⬇ PDF 다운로드",
                summary_pdf, "회원사비교분석_전체요약.pdf", "application/pdf",
                key="dl_summary",
            )
            st.caption("입력 내용을 수정한 뒤 '보고서 생성 및 미리보기'를 다시 누르면 즉시 반영됩니다.")

            # ── PDF 미리보기 (JS Blob URL — Chrome data: URI 차단 우회) ──
            import base64 as _b64
            import streamlit.components.v1 as _components
            _pdf_b64 = _b64.b64encode(summary_pdf).decode("utf-8")
            _components.html(
                f"""
                <iframe id="pdf-preview" width="100%" height="880"
                  style="border:1px solid #ddd;border-radius:4px;"></iframe>
                <script>
                (function() {{
                    var b64 = "{_pdf_b64}";
                    var bin = atob(b64);
                    var arr = new Uint8Array(bin.length);
                    for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
                    var blob = new Blob([arr], {{type: "application/pdf"}});
                    var url  = URL.createObjectURL(blob);
                    var el   = document.getElementById("pdf-preview");
                    if (el) el.src = url;
                }})();
                </script>
                """,
                height=900,
                scrolling=False,
            )
        else:
            st.info("위에서 입력 후 '보고서 생성 및 미리보기' 버튼을 눌러주세요.")

        st.divider()

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

                row_data   = row.to_dict()
                zscore_row = {col: z_all.loc[idx, col] for col in main_cols}
                z_m_row    = {col: z_method[col].loc[idx] for col in main_cols}
                inst_method = {}
                for col in main_cols:
                    comp = get_component_from_col(col, SAMPLES)
                    sfx  = get_col_suffix(col)
                    if comp:
                        mc = f"{comp}_방법{sfx}"
                        if mc in df.columns:
                            inst_method[f"{comp}{sfx}"] = str(row.get(mc, "") or "")
                method_group_stats = _calc_method_group_stats(inst_method)
                pdf_overall = generate_pdf_overall(
                    email_to, inst, row_data, zscore_row,
                    group_stats, main_cols, generated_at, SAMPLES, inst_method,
                )
                pdf_method = generate_pdf_by_method(
                    email_to, inst, row_data, z_m_row,
                    method_group_stats, main_cols, generated_at, SAMPLES, inst_method,
                )
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(
                        "전체 Z-score PDF", pdf_overall,
                        f"회원사비교분석_{inst}_전체 Robust Z-score.pdf", "application/pdf",
                        key=f"dl_overall_{idx}",
                    )
                with c2:
                    st.download_button(
                        "방법별 Z-score PDF", pdf_method,
                        f"회원사비교분석_{inst}_방법별 Robust Z-score.pdf", "application/pdf",
                        key=f"dl_method_{idx}",
                    )

        st.divider()
        st.markdown("#### 전체 일괄 발송")
        st.warning("발송 후 취소 불가. 데이터 확정 후 실행하세요.")

        # 히스토리 HTML 첨부 여부
        hist_df = get_history()
        attach_html = st.checkbox(
            "연도별 Z-Score 대시보드(HTML) 첨부",
            value=not hist_df.empty,
            disabled=hist_df.empty,
            help="히스토리 탭에서 데이터를 먼저 저장해야 활성화됩니다.",
        )

        if st.button("전체 보고서 발송", type="primary"):
            report_list = []
            for idx, row in df.iterrows():
                email_to = row.get(email_field, "")
                inst     = row.get(inst_field, "")
                if not email_to:
                    continue
                row_data   = row.to_dict()
                zscore_row = {col: z_all.loc[idx, col] for col in main_cols}
                z_m_row    = {col: z_method[col].loc[idx] for col in main_cols}
                inst_method = {}
                for col in main_cols:
                    comp = get_component_from_col(col, SAMPLES)
                    sfx  = get_col_suffix(col)
                    if comp:
                        mc = f"{comp}_방법{sfx}"
                        if mc in df.columns:
                            inst_method[f"{comp}{sfx}"] = str(row.get(mc, "") or "")
                method_group_stats = _calc_method_group_stats(inst_method)
                pdf_overall = generate_pdf_overall(
                    email_to, inst, row_data, zscore_row,
                    group_stats, main_cols, generated_at, SAMPLES, inst_method,
                )
                pdf_method = generate_pdf_by_method(
                    email_to, inst, row_data, z_m_row,
                    method_group_stats, main_cols, generated_at, SAMPLES, inst_method,
                )
                # 기관별 연도별 대시보드 HTML (본문 삽입 + 첨부)
                html_bytes = None
                html_body  = None
                if attach_html and not hist_df.empty:
                    html_bytes = generate_institution_html_bytes(hist_df, inst)
                    html_body  = generate_institution_email_html(hist_df, inst)
                report_list.append({
                    "email": email_to, "institution": inst,
                    "pdf_overall": pdf_overall, "pdf_method": pdf_method,
                    "pdf_summary": summary_pdf,
                    "html_dashboard": html_bytes,
                    "html_body": html_body,
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
    if not info_df.empty:
        info_df["order"]   = pd.to_numeric(info_df["order"], errors="coerce").fillna(1).astype(int)
        info_df["enabled"] = info_df["enabled"].map(lambda x: str(x).strip().lower() in ("true", "1", "yes"))
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
    if not sample_df.empty:
        sample_df["order"]   = pd.to_numeric(sample_df["order"], errors="coerce").fillna(1).astype(int)
        sample_df["enabled"] = sample_df["enabled"].map(lambda x: str(x).strip().lower() in ("true", "1", "yes"))
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
    if not group_df.empty:
        group_df["order"]   = pd.to_numeric(group_df["order"], errors="coerce").fillna(1).astype(int)
        group_df["enabled"] = group_df["enabled"].map(lambda x: str(x).strip().lower() in ("true", "1", "yes"))
    edited_groups_ui = st.data_editor(
        group_df[["name", "order", "enabled"]],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "name":    st.column_config.TextColumn("섹션명 *"),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
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
    edited_groups["samples"] = ""
    edited_groups["use_equip"] = ""
    edited_groups["use_solvent"] = ""
    edited_groups["free_decimal"] = ""
    edited_groups["allow_multi"]  = ""
    edited_groups = edited_groups[CONFIG_COLS]

    st.divider()

    # ── 성분 종목 ─────────────────────────────────────────────
    st.markdown("### 성분 종목")
    st.caption(
        "그룹: 위 섹션 관리에서 정의한 섹션명과 일치해야 합니다.\n\n"
        "적용 사료: `all` = 전체 / `축우사료` 또는 `축우사료,양계사료` 처럼 쉼표로 구분."
    )
    comp_df = cfg_edit[cfg_edit["type"] == "component"].reset_index(drop=True)
    # 기존 시트에 컬럼 없을 때 기본값 추가
    for _col in ("use_equip", "use_solvent"):
        if _col not in comp_df.columns:
            comp_df[_col] = True
        else:
            raw = comp_df[_col].astype(str).str.strip().str.lower()
            comp_df[_col] = ~raw.isin(["false", "0", "no"])
    if "free_decimal" not in comp_df.columns:
        comp_df["free_decimal"] = False
    else:
        raw = comp_df["free_decimal"].astype(str).str.strip().str.lower()
        comp_df["free_decimal"] = raw.isin(["true", "1", "yes"])
    if "allow_multi" not in comp_df.columns:
        comp_df["allow_multi"] = False
    else:
        raw = comp_df["allow_multi"].astype(str).str.strip().str.lower()
        comp_df["allow_multi"] = raw.isin(["true", "1", "yes"])
    if not comp_df.empty:
        comp_df["order"]   = pd.to_numeric(comp_df["order"], errors="coerce").fillna(1).astype(int)
        comp_df["enabled"] = comp_df["enabled"].map(lambda x: str(x).strip().lower() in ("true", "1", "yes"))

    # 현재 유효한 섹션명 목록 (선택용)
    valid_groups = [
        str(n) for n in edited_groups_ui["name"].dropna()
        if str(n).strip()
    ]

    edited_comps = st.data_editor(
        comp_df[["type", "group", "name", "samples", "order", "enabled", "use_equip", "use_solvent", "free_decimal", "allow_multi"]],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "type":         None,
            "group":        st.column_config.SelectboxColumn("그룹", options=valid_groups),
            "name":         st.column_config.TextColumn("성분명 *"),
            "samples":      st.column_config.TextColumn(
                "적용 사료",
                help="all = 전체 / 축우사료,양계사료 처럼 쉼표로 구분",
            ),
            "order":        st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled":      st.column_config.CheckboxColumn("활성화"),
            "use_equip":    st.column_config.CheckboxColumn("기기명 사용"),
            "use_solvent":  st.column_config.CheckboxColumn("용매 사용"),
            "free_decimal": st.column_config.CheckboxColumn(
                "소수점 자유",
                help="체크 시 소수점 제한 없음(4자리). 기본은 소수점 2자리.",
            ),
            "allow_multi": st.column_config.CheckboxColumn(
                "방법 추가 허용",
                help="체크 시 제출 폼에서 + 버튼으로 방법을 추가할 수 있습니다.",
            ),
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
    st.caption(
        "분석 방법 드롭다운에 표시될 선택지를 관리합니다.\n\n"
        "**적용 성분**: 비워두면 모든 성분에 공통 적용. "
        "성분명을 입력하면 해당 성분에만 표시됩니다 (예: `수분`, `조단백질`).\n\n"
        "특정 성분에 전용 옵션이 1개 이상 있으면 공통 옵션은 해당 성분에 표시되지 않습니다."
    )
    _raw_method = cfg_edit[cfg_edit["type"] == "method_option"][CONFIG_COLS].reset_index(drop=True)
    if _raw_method.empty:
        method_df = pd.DataFrame(columns=["group", "name", "order", "enabled"]).astype(
            {"group": str, "name": str, "order": int, "enabled": bool}
        )
    else:
        method_df = pd.DataFrame({
            "group":   _raw_method["group"].astype(str).replace("nan", ""),
            "name":    _raw_method["name"].astype(str).replace("nan", ""),
            "order":   pd.to_numeric(_raw_method["order"], errors="coerce").fillna(1).astype(int),
            "enabled": _raw_method["enabled"].map(lambda x: str(x).strip().lower() in ("true", "1", "yes")),
        })
    edited_methods = st.data_editor(
        method_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "group":   st.column_config.TextColumn(
                "적용 성분",
                help="비워두면 전체 공통 / 성분명 입력 시 해당 성분에만 표시",
            ),
            "name":    st.column_config.TextColumn("방법명 *"),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
        },
        key="edit_methods",
        hide_index=True,
    )
    edited_methods["type"]         = "method_option"
    edited_methods["samples"]      = ""
    edited_methods["use_equip"]    = ""
    edited_methods["use_solvent"]  = ""
    edited_methods["free_decimal"] = ""
    edited_methods["allow_multi"]  = ""
    edited_methods = edited_methods[CONFIG_COLS]

    st.divider()

    # ── 용매 옵션 ─────────────────────────────────────────────
    st.markdown("### 용매 옵션")
    st.caption(
        "용매 드롭다운에 표시될 선택지를 관리합니다.\n\n"
        "**적용 성분**: 비워두면 모든 성분에 공통 적용. "
        "성분명을 입력하면 해당 성분에만 표시됩니다."
    )
    _raw_solvent = cfg_edit[cfg_edit["type"] == "solvent_option"][CONFIG_COLS].reset_index(drop=True)
    if _raw_solvent.empty:
        solvent_df = pd.DataFrame(columns=["group", "name", "order", "enabled"]).astype(
            {"group": str, "name": str, "order": int, "enabled": bool}
        )
    else:
        solvent_df = pd.DataFrame({
            "group":   _raw_solvent["group"].astype(str).replace("nan", ""),
            "name":    _raw_solvent["name"].astype(str).replace("nan", ""),
            "order":   pd.to_numeric(_raw_solvent["order"], errors="coerce").fillna(1).astype(int),
            "enabled": _raw_solvent["enabled"].map(lambda x: str(x).strip().lower() in ("true", "1", "yes")),
        })
    edited_solvents = st.data_editor(
        solvent_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "group":   st.column_config.TextColumn("적용 성분", help="비워두면 전체 공통"),
            "name":    st.column_config.TextColumn("용매명 *"),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
        },
        key="edit_solvents",
        hide_index=True,
    )
    edited_solvents["type"]         = "solvent_option"
    edited_solvents["samples"]      = ""
    edited_solvents["use_equip"]    = ""
    edited_solvents["use_solvent"]  = ""
    edited_solvents["free_decimal"] = ""
    edited_solvents["allow_multi"]  = ""
    edited_solvents = edited_solvents[CONFIG_COLS]

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
    _raw_q = cfg_edit[cfg_edit["type"] == "question"][CONFIG_COLS].reset_index(drop=True)
    if not _raw_q.empty:
        _raw_q["order"]   = pd.to_numeric(_raw_q["order"], errors="coerce").fillna(1).astype(int)
        _raw_q["enabled"] = _raw_q["enabled"].map(lambda x: str(x).strip().lower() in ("true", "1", "yes"))
    question_df = _raw_q
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
            "order":     st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled":   st.column_config.CheckboxColumn("활성화"),
            "use_equip": st.column_config.CheckboxColumn("필수", help="체크 시 제출 전 반드시 입력해야 합니다."),
        },
        key="edit_questions",
        hide_index=True,
    )
    edited_questions["type"] = "question"

    st.divider()

    # ── 참가자 코드 관리 ──────────────────────────────────────
    st.markdown("### 참가자 코드")
    st.caption(
        "코드 → 회사명 매핑. 참가코드 입력 시 회사명이 자동으로 결정됩니다.\n\n"
        "**코드**: 참가자에게 배포할 고유 코드 (예: 01, A001).\n\n"
        "**회사명**: 코드 입력 시 자동으로 채워질 회사명."
    )
    _raw_part = cfg_edit[cfg_edit["type"] == "participant"][CONFIG_COLS].reset_index(drop=True)
    if _raw_part.empty:
        part_df = pd.DataFrame(columns=["group", "name", "order", "enabled"]).astype(
            {"group": str, "name": str, "order": int, "enabled": bool}
        )
    else:
        part_df = pd.DataFrame({
            "group":   _raw_part["group"].astype(str).replace("nan", ""),
            "name":    _raw_part["name"].astype(str).replace("nan", ""),
            "order":   pd.to_numeric(_raw_part["order"], errors="coerce").fillna(1).astype(int),
            "enabled": _raw_part["enabled"].map(lambda x: str(x).strip().lower() in ("true", "1", "yes")),
        })
    edited_participants = st.data_editor(
        part_df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "group":   st.column_config.TextColumn("코드 *", help="참가자에게 배포할 고유 코드"),
            "name":    st.column_config.TextColumn("회사명 *"),
            "order":   st.column_config.NumberColumn("순서", min_value=1, step=1),
            "enabled": st.column_config.CheckboxColumn("활성화"),
        },
        key="edit_participants",
        hide_index=True,
    )
    edited_participants["type"]         = "participant"
    edited_participants["samples"]      = ""
    edited_participants["use_equip"]    = ""
    edited_participants["use_solvent"]  = ""
    edited_participants["free_decimal"] = ""
    edited_participants["allow_multi"]  = ""
    edited_participants = edited_participants[CONFIG_COLS]

    st.divider()

    # ── 저장 ──────────────────────────────────────────────────
    col_save, col_reset = st.columns([1, 1])
    with col_save:
        if st.button("설정 저장", type="primary", use_container_width=True):
            edited_info         = edited_info[edited_info["name"].astype(str).str.strip() != ""]
            edited_samples      = edited_samples[edited_samples["name"].astype(str).str.strip() != ""]
            edited_groups       = edited_groups[edited_groups["name"].astype(str).str.strip() != ""]
            edited_comps        = edited_comps[edited_comps["name"].astype(str).str.strip() != ""]
            edited_methods      = edited_methods[edited_methods["name"].astype(str).str.strip() != ""]
            edited_solvents     = edited_solvents[edited_solvents["name"].astype(str).str.strip() != ""]
            edited_questions    = edited_questions[edited_questions["name"].astype(str).str.strip() != ""]
            edited_participants = edited_participants[edited_participants["name"].astype(str).str.strip() != ""]

            # use_equip / use_solvent / free_decimal / allow_multi 컬럼이 없는 타입은 빈 문자열로 채움
            for _df in [edited_info, edited_samples, edited_groups, edited_methods, edited_solvents, edited_questions, edited_participants]:
                for _col in ("use_equip", "use_solvent", "free_decimal", "allow_multi"):
                    if _col not in _df.columns:
                        _df[_col] = ""
            new_cfg = pd.concat(
                [edited_info, edited_samples, edited_groups,
                 edited_comps, edited_methods, edited_solvents, edited_questions, edited_participants],
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


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab5:
    st.subheader("연도별 히스토리 관리")
    st.caption(
        "회차별 참가기관 중앙값을 연도별로 누적 저장합니다. "
        "저장된 데이터는 보고서 발송 시 **연도별 Z-Score 대시보드(HTML)**로 첨부됩니다."
    )

    hist_df = get_history()

    # ── 현재 히스토리 표시 ────────────────────────────────────
    if hist_df.empty:
        st.info("저장된 히스토리가 없습니다. 아래에서 이번 회차 데이터를 저장하세요.")
    else:
        st.markdown("#### 저장된 히스토리")
        st.dataframe(hist_df, use_container_width=True, hide_index=True)

        # 기관별 HTML 미리보기 다운로드
        insts_in_hist = sorted([str(v) for v in hist_df["institution"].dropna().unique()]) if "institution" in hist_df.columns else []
        if insts_in_hist:
            preview_inst = st.selectbox("미리보기 기관 선택", insts_in_hist, key="preview_inst")
            html_bytes = generate_institution_html_bytes(hist_df, preview_inst)
            st.download_button(
                f"{preview_inst} 대시보드 HTML 다운로드",
                data=html_bytes,
                file_name=f"사료_ZScore_{preview_inst}_연도별.html",
                mime="text/html",
            )

        # 행 삭제
        st.markdown("#### 히스토리 행 삭제")
        years_in_hist = sorted([str(v) for v in hist_df["year"].dropna().unique()])
        feeds_in_hist = sorted([str(v) for v in hist_df["feed"].dropna().unique()])
        del_col1, del_col2, del_col3 = st.columns([1, 1, 1])
        with del_col1:
            del_year = st.selectbox("삭제할 연도", years_in_hist, key="del_year")
        with del_col2:
            del_feed = st.selectbox("삭제할 사료", feeds_in_hist, key="del_feed")
        with del_col3:
            st.write("")
            st.write("")
            if st.button("선택 행 삭제", type="secondary"):
                delete_history_rows(del_year, del_feed)
                st.success(f"{del_year}년 {del_feed} 삭제 완료")
                st.rerun()

    st.divider()

    # ── 이번 회차 데이터 저장 ─────────────────────────────────
    st.markdown("#### 이번 회차 데이터를 히스토리에 저장")
    st.caption("현재 제출된 기관별 원값을 연도별로 누적 저장합니다. 각 기관이 자신의 연도별 추이를 확인할 수 있습니다.")

    if df.empty:
        st.warning("제출된 데이터가 없습니다.")
    else:
        save_year = st.number_input("저장 연도", min_value=2000, max_value=2100,
                                    value=datetime.now(KST).year, step=1)

        # 기관별·사료별 원값 수집
        preview_rows = []
        for _, row in df.iterrows():
            inst_name = row.get(inst_field, "")
            if not inst_name:
                continue
            for sample in SAMPLES:
                sample_cols = [c for c in main_cols if c.endswith(f"_{sample}")]
                if not sample_cols:
                    continue
                row_dict = {"year": int(save_year), "feed": sample, "institution": inst_name}
                has_val = False
                for col in sample_cols:
                    comp = get_component_from_col(col, SAMPLES)
                    if comp:
                        val = pd.to_numeric(row.get(col, None), errors="coerce")
                        row_dict[comp] = round(float(val), 4) if not pd.isna(val) else None
                        if not pd.isna(val):
                            has_val = True
                if has_val:
                    preview_rows.append(row_dict)

        if preview_rows:
            st.markdown(f"**저장 미리보기** — {len(preview_rows)}행 (기관 × 사료)")
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

            if st.button("히스토리에 저장", type="primary"):
                # 동일 연도+사료 기존 행 전체 삭제 후 새로 저장
                deleted_keys = set()
                for pr in preview_rows:
                    key = (pr["year"], pr["feed"])
                    if key not in deleted_keys:
                        delete_history_rows(pr["year"], pr["feed"])
                        deleted_keys.add(key)
                append_history_rows(preview_rows)
                st.success(f"{save_year}년 데이터 저장 완료 ({len(preview_rows)}행)")
                st.rerun()
