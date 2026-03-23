import io
from datetime import datetime
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, Line
from reportlab.graphics.shapes import String as GStr

from utils.config import get_component_from_col, get_sample_from_col

# 한글 TTF 폰트 등록 (Streamlit Cloud: packages.txt에 fonts-nanum 필요)
try:
    pdfmetrics.registerFont(TTFont(
        "NanumGothic",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    ))
    KO = "NanumGothic"
except Exception:
    KO = "Helvetica"  # 한글 깨질 수 있음 (로컬 환경 폴백)



def _make_table_style() -> TableStyle:
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTNAME",   (0, 0), (-1, -1), KO),
        ("FONTSIZE",   (0, 0), (-1, -1), 8),
        ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ])


def _build_sample_section(
    sample_name: str,
    cols_for_sample: list,
    row_data: dict,
    zscore_dict: dict,
    group_stats: dict,
    styles,
    samples: list,
    report_type: str,       # "overall" or "method"
    inst_method: dict = None,
) -> list:
    elements = []
    section_style = ParagraphStyle(
        "sec", parent=styles["Heading2"], fontSize=11, spaceAfter=3, spaceBefore=6,
        fontName=KO,
    )
    elements.append(Paragraph(sample_name, section_style))

    def fmt(v, d=2):
        try:
            f = float(v)
            return f"{f:.{d}f}" if not np.isnan(f) else "-"
        except Exception:
            return "-"

    z_cell_style = ParagraphStyle("zcell", fontName=KO, fontSize=8, alignment=TA_CENTER)

    if report_type == "overall":
        header = ["성분", "방법", "제출값", "중앙값", "CV(%)", "n", "Z전체"]
        cw = [30*mm, 35*mm, 22*mm, 22*mm, 16*mm, 12*mm, 23*mm]
    else:
        header = ["성분", "방법", "제출값", "중앙값", "CV(%)", "n", "Z방법별"]
        cw = [30*mm, 35*mm, 22*mm, 22*mm, 16*mm, 12*mm, 23*mm]

    def _z_cell(z_f: float, z_str: str):
        if z_str == "N/A":
            return z_str
        if abs(z_f) > 3:
            return Paragraph(f'<font color="red"><b><u>{z_str}</u></b></font>', z_cell_style)
        return z_str

    rows = [header]
    for col in cols_for_sample:
        comp  = get_component_from_col(col, samples) or col
        val   = row_data.get(col, "")

        # 제출값이 없는 항목은 제외 (빈 문자열, None, NaN, 0 포함)
        try:
            _empty = val is None or str(val).strip() in ("", "nan") or (isinstance(val, float) and np.isnan(val))
        except Exception:
            _empty = False
        if _empty:
            continue

        stats = group_stats.get(col, {})
        z     = zscore_dict.get(col, np.nan)

        try:
            z_f   = float(z)
            z_str = f"{z_f:.2f}" if not np.isnan(z_f) else "N/A"
        except Exception:
            z_f, z_str = np.nan, "N/A"

        cv_val = stats.get("cv")
        try:
            cv_str = f"{float(cv_val):.1f}" if cv_val is not None and not np.isnan(float(cv_val)) else "-"
        except Exception:
            cv_str = "-"

        # suffix-aware method lookup (e.g. col="조단백질_축우사료_2" → sfx="_2")
        sample_in_col = get_sample_from_col(col, samples) or ""
        sfx = col[len(f"{comp}_{sample_in_col}"):] if sample_in_col else ""
        method = row_data.get(f"{comp}_방법{sfx}", "") or (inst_method or {}).get(comp, "")
        rows.append([
            comp, method, fmt(val),
            fmt(stats.get("median", "")),
            cv_str,
            str(stats.get("n", "")),
            _z_cell(z_f, z_str),
        ])

    # 데이터 행이 없으면 섹션 전체 스킵
    if len(rows) == 1:
        return []

    t = Table(rows, colWidths=cw, repeatRows=1)
    t.setStyle(_make_table_style())
    elements.append(t)
    elements.append(Spacer(1, 4*mm))
    return elements


def _generate_zscore_pdf(
    title: str,
    email: str,
    institution: str,
    row_data: dict,
    zscore_dict: dict,
    group_stats: dict,
    value_cols: list,
    report_type: str,
    generated_at: str,
    samples: list,
    inst_method: dict = None,
) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title2", parent=styles["Title"], fontSize=16, spaceAfter=4,
        alignment=TA_CENTER, fontName=KO,
    )
    info_style = ParagraphStyle("info2", parent=styles["Normal"], fontSize=10,
                                spaceAfter=2, fontName=KO)
    note_style = ParagraphStyle("note2", parent=styles["Normal"], fontSize=8,
                                textColor=colors.grey, leftIndent=4, fontName=KO)

    elements = []
    elements.append(Paragraph(title, title_style))
    elements.append(Spacer(1, 2*mm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2c3e50")))
    elements.append(Spacer(1, 4*mm))
    elements.append(Paragraph(f"<b>기관명:</b> {institution}", info_style))
    elements.append(Paragraph(f"<b>이메일:</b> {email}", info_style))
    elements.append(Paragraph(f"<b>보고서 생성일:</b> {generated_at}", info_style))
    elements.append(Spacer(1, 5*mm))

    # NIR 제외 컬럼만 사용
    non_nir_cols = [c for c in value_cols if not c.startswith("NIR_")]

    # 사료 종류별로 그룹핑 (samples 순서 유지)
    sample_to_cols: dict[str, list] = {}
    for col in non_nir_cols:
        s = get_sample_from_col(col, samples) or "-"
        sample_to_cols.setdefault(s, []).append(col)

    # samples 리스트 순서대로 섹션 출력
    ordered_samples = [s for s in samples if s in sample_to_cols]
    for s in sample_to_cols:
        if s not in ordered_samples:
            ordered_samples.append(s)

    for sample_name in ordered_samples:
        scols = sample_to_cols[sample_name]
        elements += _build_sample_section(
            sample_name, scols, row_data, zscore_dict, group_stats,
            styles, samples, report_type, inst_method,
        )

    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
    elements.append(Spacer(1, 3*mm))
    elements.append(Paragraph("<b>판정 기준 (Robust Z-score)</b>", info_style))
    elements.append(Paragraph("적합: |Z| <= 2.0", note_style))
    elements.append(Paragraph("경고: 2.0 < |Z| <= 3.0", note_style))
    elements.append(Paragraph("부적합: |Z| > 3.0", note_style))
    elements.append(Spacer(1, 2*mm))
    if report_type == "overall":
        elements.append(Paragraph(
            "* Robust Z-score = (제출값 - 중앙값) / (1.4826 x MAD)  |  CV(%) = 표준편차 / 평균 x 100",
            note_style,
        ))
    else:
        elements.append(Paragraph(
            "* Z방법별: 동일 방법 사용 기관 5개 미만이면 N/A  |  CV(%) = 표준편차 / 평균 x 100",
            note_style,
        ))

    doc.build(elements)
    return buf.getvalue()


def generate_pdf_overall(
    email: str,
    institution: str,
    row_data: dict,
    zscore_row: dict,
    group_stats: dict,
    value_cols: list,
    generated_at: str = None,
    samples: list = None,
    inst_method: dict = None,
) -> bytes:
    if samples is None:
        from utils.config import get_samples
        samples = get_samples()
    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M")
    return _generate_zscore_pdf(
        title="한국사료협회 비교분석 시험 Robust Z-score 보고서 (전체)",
        email=email,
        institution=institution,
        row_data=row_data,
        zscore_dict=zscore_row,
        group_stats=group_stats,
        value_cols=value_cols,
        report_type="overall",
        generated_at=generated_at,
        samples=samples,
        inst_method=inst_method,
    )


def generate_pdf_by_method(
    email: str,
    institution: str,
    row_data: dict,
    z_method_row: dict,
    group_stats: dict,
    value_cols: list,
    generated_at: str = None,
    samples: list = None,
    inst_method: dict = None,
) -> bytes:
    if samples is None:
        from utils.config import get_samples
        samples = get_samples()
    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M")
    return _generate_zscore_pdf(
        title="한국사료협회 비교분석 시험 Robust Z-score 보고서 (방법별)",
        email=email,
        institution=institution,
        row_data=row_data,
        zscore_dict=z_method_row,
        group_stats=group_stats,
        value_cols=value_cols,
        report_type="method",
        generated_at=generated_at,
        samples=samples,
        inst_method=inst_method,
    )


def generate_pdf_summary(
    df,
    z_all,
    z_method: dict,
    group_stats: dict,
    value_cols: list,
    inst_field: str,
    generated_at: str,
    samples: list,
    participant_map: dict = None,
) -> bytes:
    """전체 요약 보고서: 통계 요약표 + 전체/방법별 Z-score 표."""
    import pandas as pd

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )
    styles = getSampleStyleSheet()
    title_style   = ParagraphStyle("ts", parent=styles["Title"],   fontSize=14, spaceAfter=4,  alignment=TA_CENTER, fontName=KO)
    info_style    = ParagraphStyle("is", parent=styles["Normal"],  fontSize=9,  spaceAfter=2,  fontName=KO)
    section_style = ParagraphStyle("ss", parent=styles["Heading2"],fontSize=11, spaceAfter=3,  spaceBefore=8,  fontName=KO)
    comp_style    = ParagraphStyle("cs", parent=styles["Normal"],  fontSize=9,  spaceAfter=2,  spaceBefore=5,  fontName=KO)
    note_style    = ParagraphStyle("ns", parent=styles["Normal"],  fontSize=7,  textColor=colors.grey, leftIndent=4, fontName=KO)

    def fmt(v, d=2):
        try:
            f = float(v)
            return f"{f:.{d}f}" if not np.isnan(f) else "-"
        except Exception:
            return "-"

    import pandas as _pd
    non_nir = [c for c in value_cols if not c.startswith("NIR_")]

    # 데이터가 없는 컬럼 제거 (통계 없음 = 제출 데이터 없음)
    non_nir = [c for c in non_nir if group_stats.get(c)]
    non_nir_set = set(non_nir)

    def _z_has_data(col, z_src):
        """해당 컬럼의 Z-score가 하나라도 유효한지 확인."""
        try:
            if isinstance(z_src, _pd.DataFrame):
                return z_src[col].notna().any() if col in z_src.columns else False
            ser = z_src.get(col)
            return ser.notna().any() if ser is not None else False
        except Exception:
            return False

    # 성분 목록 (순서 유지)
    seen_comps: list = []
    comp_to_cols: dict = {}
    for col in non_nir:
        comp = get_component_from_col(col, samples) or col
        if comp not in seen_comps:
            seen_comps.append(comp)
        comp_to_cols.setdefault(comp, []).append(col)

    avail_mm = 180
    comp_w   = 35

    elements = []

    # 제목
    _now = datetime.now()
    _half = "상반기" if _now.month <= 6 else "하반기"
    _report_title = f"{_now.year}년 {_half} 한국사료협회 비교분석 전체 보고서"
    elements.append(Paragraph(_report_title, title_style))
    elements.append(Spacer(1, 2*mm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2c3e50")))
    elements.append(Spacer(1, 5*mm))

    # ━━━ 1. 통계 요약 ━━━
    elements.append(Paragraph("1. 전체 통계 요약", section_style))

    valid_stat_samples = [s for s in samples if any(f"{c}_{s}" in non_nir_set for c in seen_comps)]

    # ── 열 너비 계산 ──
    comp_w   = 18   # 성분 (mm)
    method_w = 30   # 분석법 (mm)
    remain   = avail_mm - comp_w - method_w
    n_feeds  = max(len(valid_stat_samples), 1)
    per_feed = remain / n_feeds          # 각 사료 4열 합계 (mm)
    n_col = 4                            # N / 평균 / 표준편차 / 변이계수
    sub_w = per_feed / n_col
    cw_stat = ([comp_w*mm, method_w*mm]
               + [sub_w*mm] * n_col * len(valid_stat_samples))

    cell_s = ParagraphStyle("csp", fontName=KO, fontSize=7,
                            leading=9, alignment=TA_CENTER)
    def _p(txt):
        return Paragraph(str(txt), cell_s)

    # ── 헤더 2행 ──
    hdr0 = [_p("성분"), _p("분석법")]
    for s in valid_stat_samples:
        hdr0 += [_p(s), _p(""), _p(""), _p("")]
    hdr1 = [_p(""), _p("")]
    for _ in valid_stat_samples:
        hdr1 += [_p("N"), _p("평균"), _p("표준편차"), _p("변이계수")]
    stat_rows = [hdr0, hdr1]

    span_cmds = [
        ("SPAN", (0, 0), (0, 1)),  # 성분 헤더 병합
        ("SPAN", (1, 0), (1, 1)),  # 분석법 헤더 병합
    ]
    for i, _ in enumerate(valid_stat_samples):
        c0 = 2 + i * n_col
        span_cmds.append(("SPAN", (c0, 0), (c0 + n_col - 1, 0)))  # 사료명 병합

    whole_rows: set = set()   # "전체" 행 인덱스 추적

    for comp in seen_comps:
        method_col = f"{comp}_방법"
        methods: list[str] = []
        if method_col in df.columns:
            methods = sorted(df[method_col].dropna().astype(str).unique().tolist())

        comp_rows_data = []
        for meth in methods:
            row = [_p(""), _p(meth)]
            for s in valid_stat_samples:
                col = f"{comp}_{s}"
                if col not in df.columns:
                    row += [_p("")]*n_col; continue
                mask = (df[method_col].astype(str) == meth)
                vals = pd.to_numeric(df.loc[mask, col], errors="coerce").dropna()
                if len(vals) == 0:
                    row += [_p("")]*n_col; continue
                mean_ = vals.mean()
                std_  = vals.std(ddof=1) if len(vals) > 1 else float("nan")
                cv_   = (std_/mean_*100) if mean_ != 0 and not np.isnan(std_) else float("nan")
                row += [_p(len(vals)),
                        _p(fmt(mean_)),
                        _p(fmt(std_) if not np.isnan(std_) else "-"),
                        _p(fmt(cv_, 1) if not np.isnan(cv_) else "-")]
            comp_rows_data.append(row)

        # 전체 행
        whole_row = [_p(""), _p(f"{comp} 전체")]
        for s in valid_stat_samples:
            col = f"{comp}_{s}"
            if col not in df.columns:
                whole_row += [_p("")]*n_col; continue
            vals = pd.to_numeric(df[col], errors="coerce").dropna()
            if len(vals) == 0:
                whole_row += [_p("")]*n_col; continue
            mean_ = vals.mean()
            std_  = vals.std(ddof=1) if len(vals) > 1 else float("nan")
            cv_   = (std_/mean_*100) if mean_ != 0 and not np.isnan(std_) else float("nan")
            whole_row += [_p(len(vals)),
                          _p(fmt(mean_)),
                          _p(fmt(std_) if not np.isnan(std_) else "-"),
                          _p(fmt(cv_, 1) if not np.isnan(cv_) else "-")]
        comp_rows_data.append(whole_row)

        # 성분명은 첫 번째 행에만, 세로 병합
        if comp_rows_data:
            comp_rows_data[0][0] = _p(comp)
        base_idx = len(stat_rows)
        if len(comp_rows_data) > 1:
            span_cmds.append(("SPAN", (0, base_idx),
                              (0, base_idx + len(comp_rows_data) - 1)))
        whole_rows.add(base_idx + len(comp_rows_data) - 1)
        stat_rows.extend(comp_rows_data)

    stat_tbl = Table(stat_rows, colWidths=cw_stat, repeatRows=2)
    tbl_style_cmds = [
        ("BACKGROUND",    (0, 0), (-1, 1),  colors.HexColor("#2c3e50")),
        ("TEXTCOLOR",     (0, 0), (-1, 1),  colors.white),
        ("FONTNAME",      (0, 0), (-1, -1), KO),
        ("FONTSIZE",      (0, 0), (-1, -1), 7),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#dee2e6")),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING",   (0, 0), (-1, -1), 2),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS",(1, 2), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
        # 성분 열(0)은 흰색 고정 — ROWBACKGROUNDS가 SPAN 내부 행에 적용되어 색 갈림 방지
        ("BACKGROUND",    (0, 2), (0, -1),  colors.white),
    ] + span_cmds
    # 전체 행 강조
    for ri in whole_rows:
        tbl_style_cmds += [
            ("BACKGROUND",  (0, ri), (-1, ri), colors.HexColor("#e8edf2")),
            ("FONTNAME",    (0, ri), (-1, ri), KO),
        ]
    stat_tbl.setStyle(TableStyle(tbl_style_cmds))
    elements.append(stat_tbl)
    elements.append(Spacer(1, 6*mm))

    # ━━━ CV 세로 그룹 막대그래프 ━━━
    elements.append(Paragraph("각 성분별 변이계수(CV%)", section_style))

    # 데이터 수집: {comp: {feed: cv_value}}
    cv_data: dict[str, dict[str, float]] = {}
    for comp in seen_comps:
        for s in valid_stat_samples:
            col = f"{comp}_{s}"
            sd  = group_stats.get(col, {})
            try:
                cv_f = float(sd.get("cv", ""))
                if not np.isnan(cv_f):
                    cv_data.setdefault(comp, {})[s] = cv_f
            except Exception:
                pass

    if cv_data:
        comps_with_data  = [c for c in seen_comps if c in cv_data]
        feeds_with_data  = [s for s in valid_stat_samples
                            if any(s in cv_data.get(c, {}) for c in comps_with_data)]

        _PALETTE = ["#2563eb","#d97706","#16a34a","#dc2626","#7c3aed","#0891b2","#be185d"]
        feed_color = {s: _PALETTE[i % len(_PALETTE)] for i, s in enumerate(feeds_with_data)}

        # ── 치수 ──
        chart_w    = avail_mm * mm
        ml, mr, mb, mt = 28, 6, 32, 10   # margin left/right/bottom/top (pt)
        plot_w  = chart_w - ml - mr
        plot_h  = 110.0

        n_comps = len(comps_with_data)
        n_feeds = len(feeds_with_data)
        group_w = plot_w / n_comps
        bar_w   = min((group_w - 4) / max(n_feeds, 1), 14)
        group_bw = bar_w * n_feeds

        all_vals = [v for d in cv_data.values() for v in d.values()]
        max_cv   = max(all_vals) * 1.15 if all_vals else 10.0
        max_cv   = max(max_cv, 5.0)

        legend_h = 14 * ((n_feeds + 2) // 3)   # 3열 범례
        total_h  = mt + plot_h + mb + legend_h + 4

        d = Drawing(chart_w, total_h)
        base_y = legend_h + mb   # 플롯 영역 하단 y

        # Y축 눈금선 + 레이블
        n_ticks = 5
        for ti in range(n_ticks + 1):
            tv  = max_cv * ti / n_ticks
            ty  = base_y + (tv / max_cv) * plot_h
            d.add(GStr(ml - 3, ty - 3, f"{tv:.0f}", fontSize=6, fontName=KO,
                       textAnchor="end", fillColor=colors.HexColor("#666666")))
            d.add(Line(ml, ty, ml + plot_w, ty,
                       strokeColor=colors.HexColor("#e5e7eb"), strokeWidth=0.5))

        # 축선
        d.add(Line(ml, base_y, ml + plot_w, base_y,
                   strokeColor=colors.HexColor("#999999"), strokeWidth=0.7))
        d.add(Line(ml, base_y, ml, base_y + plot_h,
                   strokeColor=colors.HexColor("#999999"), strokeWidth=0.7))

        # 막대 + X축 레이블
        for ci, comp in enumerate(comps_with_data):
            gx = ml + ci * group_w + (group_w - group_bw) / 2
            for fi, feed in enumerate(feeds_with_data):
                cv_val = cv_data.get(comp, {}).get(feed)
                if cv_val is None:
                    continue
                bx = gx + fi * bar_w
                bh = (cv_val / max_cv) * plot_h
                d.add(Rect(bx, base_y, bar_w - 0.5, bh,
                           fillColor=colors.HexColor(feed_color[feed]),
                           strokeColor=None))
            # X축 성분 레이블 (긴 이름은 줄임)
            label = comp if len(comp) <= 6 else comp[:5] + "…"
            d.add(GStr(ml + ci * group_w + group_w / 2, base_y - 10,
                       label, fontSize=6, fontName=KO,
                       textAnchor="middle", fillColor=colors.black))

        # 범례 (3열)
        cols_per_row = 3
        for fi, feed in enumerate(feeds_with_data):
            row, col = divmod(fi, cols_per_row)
            lx = ml + col * (plot_w / cols_per_row)
            ly = legend_h - 12 - row * 13
            d.add(Rect(lx, ly, 9, 7,
                       fillColor=colors.HexColor(feed_color[feed]), strokeColor=None))
            d.add(GStr(lx + 12, ly, feed, fontSize=7, fontName=KO,
                       textAnchor="start", fillColor=colors.black))

        elements.append(d)

        # ── 하단 데이터 표 ──
        tbl_header = [""] + comps_with_data
        tbl_rows   = [tbl_header]
        for feed in feeds_with_data:
            row = [feed]
            for comp in comps_with_data:
                cv_val = cv_data.get(comp, {}).get(feed)
                row.append(f"{cv_val:.2f}" if cv_val is not None else "")
            tbl_rows.append(row)

        n_cols   = len(tbl_header)
        feed_col_w = 30 * mm
        comp_col_w = (avail_mm * mm - feed_col_w) / max(n_cols - 1, 1)
        tbl_cw   = [feed_col_w] + [comp_col_w] * (n_cols - 1)

        tbl = Table(tbl_rows, colWidths=tbl_cw)
        tbl_style = TableStyle([
            ("FONTNAME",    (0,0), (-1,-1), KO),
            ("FONTSIZE",    (0,0), (-1,-1), 7),
            ("ALIGN",       (1,0), (-1,-1), "CENTER"),
            ("ALIGN",       (0,0), (0,-1),  "LEFT"),
            ("BACKGROUND",  (0,0), (-1,0),  colors.HexColor("#dbeafe")),
            ("BACKGROUND",  (0,0), (0,-1),  colors.HexColor("#f1f5f9")),
            ("GRID",        (0,0), (-1,-1), 0.4, colors.HexColor("#cbd5e1")),
            ("TOPPADDING",  (0,0), (-1,-1), 2),
            ("BOTTOMPADDING",(0,0),(-1,-1), 2),
        ])
        # 사료별 행 색상 (범례 색상과 일치)
        for fi, feed in enumerate(feeds_with_data):
            fc = colors.HexColor(feed_color[feed])
            tbl_style.add("TEXTCOLOR", (0, fi+1), (0, fi+1), fc)
            tbl_style.add("FONTNAME",  (0, fi+1), (0, fi+1), KO)
        tbl.setStyle(tbl_style)
        elements.append(tbl)

    elements.append(Spacer(1, 8*mm))

    # ━━━ Z-score 공통 ━━━
    z_cell_p = ParagraphStyle("zcp", fontName=KO, fontSize=8, alignment=TA_CENTER)

    def z_cell(z_val):
        try:
            z_f = float(z_val)
            if np.isnan(z_f):
                return "N/A"
            z_str = f"{z_f:.2f}"
        except Exception:
            return "N/A"
        if abs(z_f) > 3:
            return Paragraph(f'<font color="red"><b><u>{z_str}</u></b></font>', z_cell_p)
        return z_str

    # 회사명 → 참가코드 역변환 (participant_map: {코드: 회사명})
    _name_to_code = {v: k for k, v in (participant_map or {}).items()}
    raw_names  = df[inst_field].fillna("").astype(str).tolist()
    inst_names = [_name_to_code.get(n, n) for n in raw_names]
    idx_list   = df.index.tolist()

    inst_w = 45

    def _get_zv(z_src, row_idx, col):
        try:
            if isinstance(z_src, _pd.DataFrame):
                return z_src.at[row_idx, col] if col in z_src.columns else np.nan
            ser = z_src.get(col)
            return ser.at[row_idx] if ser is not None else np.nan
        except Exception:
            return np.nan

    def _build_zscore_section(sec_title, z_src):
        elems = [Paragraph(sec_title, section_style)]
        for num, comp in enumerate(seen_comps, 1):
            # 이 성분의 모든 suffix 탐색 ("", "_2", "_3", ...)
            sfx_list = [""]
            for n in range(2, 10):
                if any(f"{comp}_{s}_{n}" in non_nir_set for s in samples):
                    sfx_list.append(f"_{n}")
                else:
                    break

            # suffix별로 Z-score 데이터 있는 사료 합산
            valid_samples = [s for s in samples
                             if any(f"{comp}_{s}{sfx}" in non_nir_set
                                    and _z_has_data(f"{comp}_{s}{sfx}", z_src)
                                    for sfx in sfx_list)]
            if not valid_samples:
                continue

            samp_w = (avail_mm - inst_w) / len(valid_samples)
            cw_z   = [inst_w*mm] + [samp_w*mm] * len(valid_samples)

            elems.append(Paragraph(f"{num}) {comp}", comp_style))
            z_rows = [["참가코드"] + valid_samples]
            for inst, row_idx in zip(inst_names, idx_list):
                for sfx in sfx_list:
                    # 이 기관에 해당 suffix 데이터가 하나라도 있어야 행 추가
                    row_vals = []
                    has_any = False
                    for s in valid_samples:
                        col = f"{comp}_{s}{sfx}"
                        zv = _get_zv(z_src, row_idx, col) if col in non_nir_set else np.nan
                        try:
                            if not np.isnan(float(zv)):
                                has_any = True
                        except Exception:
                            pass
                        row_vals.append(z_cell(zv))
                    if not has_any:
                        continue
                    sfx_num = sfx.lstrip("_")  # "" or "2"
                    inst_label = inst if sfx == "" else f"{inst}_{sfx_num}"
                    z_rows.append([inst_label] + row_vals)

            zt = Table(z_rows, colWidths=cw_z, repeatRows=1)
            zt.setStyle(_make_table_style())
            elems.append(zt)
            elems.append(Spacer(1, 4*mm))
        return elems

    elements += _build_zscore_section("2. 전체 Z-score", z_all)
    elements += _build_zscore_section("3. 방법별 Z-score", z_method)

    # 범례
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
    elements.append(Spacer(1, 3*mm))
    elements.append(Paragraph("<b>판정 기준 (Robust Z-score)</b>", info_style))
    elements.append(Paragraph("적합: |Z| ≤ 2.0   경고: 2.0 < |Z| ≤ 3.0   부적합: |Z| > 3.0", note_style))

    doc.build(elements)
    return buf.getvalue()


def generate_submission_pdf(
    row: dict,
    cfg,
    generated_at: str = None,
) -> bytes:
    """제출 데이터를 표 형식으로 보여주는 확인용 PDF (Z-score 없음)."""
    from utils.config import get_component_groups, get_nir_groups, get_info_fields, get_samples

    GROUPS      = get_component_groups(cfg)
    NIR_GROUPS  = get_nir_groups(cfg)
    INFO_FIELDS = get_info_fields(cfg)
    ALL_SAMPLES = get_samples(cfg)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )
    styles    = getSampleStyleSheet()
    title_sty = ParagraphStyle(
        "title", parent=styles["Title"], fontSize=16, spaceAfter=4,
        alignment=TA_CENTER, fontName=KO,
    )
    info_sty = ParagraphStyle("info", parent=styles["Normal"], fontSize=10,
                              spaceAfter=2, fontName=KO)
    sec_sty  = ParagraphStyle(
        "sec", parent=styles["Heading2"], fontSize=11, spaceAfter=3,
        spaceBefore=6, fontName=KO,
    )

    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M")
    elements = []

    elements.append(Paragraph("데이터 제출 확인서", title_sty))
    elements.append(Spacer(1, 5*mm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2c3e50")))
    elements.append(Spacer(1, 4*mm))

    for field in INFO_FIELDS:
        val = row.get(field["name"], "")
        if val:
            elements.append(Paragraph(f"<b>{field['name']}:</b> {val}", info_sty))
    elements.append(Paragraph(f"<b>제출일시:</b> {row.get('제출일시', generated_at)}", info_sty))
    elements.append(Spacer(1, 5*mm))

    def _fmt(v) -> str:
        """수치는 소수점 2자리, 빈 값은 빈 문자열로 반환."""
        if v is None or v == "":
            return ""
        try:
            return f"{float(v):.2f}"
        except (ValueError, TypeError):
            return str(v)

    for group_name, items in GROUPS.items():
        grp_samples: list[str] = []
        for item in items:
            for s in item["samples"]:
                if s not in grp_samples:
                    grp_samples.append(s)

        header = ["성분", "방법", "기기명", "용매"] + grp_samples
        tbl_rows = [header]
        for item in items:
            comp = item["name"]
            # 기본 행 + 추가 방법 행(_2, _3, ...) 모두 출력
            suffixes = [""]
            for n in range(2, 10):
                sfx = f"_{n}"
                if any(row.get(f"{comp}_{s}{sfx}") not in (None, "", "nan") for s in grp_samples):
                    suffixes.append(sfx)
                else:
                    break
            for sfx in suffixes:
                method  = str(row.get(f"{comp}_방법{sfx}", "") or "")
                equip   = str(row.get(f"{comp}_기기{sfx}",  "") or "")
                solvent = str(row.get(f"{comp}_용매{sfx}",  "") or "")
                vals    = [_fmt(row.get(f"{comp}_{s}{sfx}")) for s in grp_samples]
                comp_label = comp if sfx == "" else f"  ↳ {comp}"
                tbl_rows.append([comp_label, method, equip, solvent] + vals)

        elements.append(Paragraph(group_name, sec_sty))

        fixed_w    = [28*mm, 28*mm, 25*mm, 20*mm]
        remaining  = 180*mm - sum(fixed_w)
        sample_w   = [remaining / len(grp_samples)] * len(grp_samples) if grp_samples else []
        t = Table(tbl_rows, colWidths=fixed_w + sample_w, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
            ("TEXTCOLOR",     (0, 0), (-1, 0), colors.white),
            ("FONTNAME",      (0, 0), (-1, -1), KO),
            ("FONTSIZE",      (0, 0), (-1, -1), 8),
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
            ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
            ("TOPPADDING",    (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 4*mm))

    # NIR 섹션
    def _render_nir_table(nir_rows: list):
        nir_header = nir_rows[0]
        n_samp = len(nir_header) - 2
        fixed_w = [35*mm, 35*mm]
        remaining = 180*mm - sum(fixed_w)
        sample_w = [remaining / n_samp] * n_samp if n_samp > 0 else []
        t = Table(nir_rows, colWidths=fixed_w + sample_w, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
            ("TEXTCOLOR",     (0, 0), (-1, 0), colors.white),
            ("FONTNAME",      (0, 0), (-1, -1), KO),
            ("FONTSIZE",      (0, 0), (-1, -1), 8),
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
            ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
            ("TOPPADDING",    (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        return t

    if NIR_GROUPS:
        elements.append(Paragraph("NIR 측정값", sec_sty))
        for _, items in NIR_GROUPS.items():
            grp_samples: list[str] = []
            for item in items:
                for s in item["samples"]:
                    if s not in grp_samples:
                        grp_samples.append(s)
            nir_rows = [["성분", "기기명"] + grp_samples]
            for item in items:
                comp  = item["name"]
                equip = str(row.get(f"NIR_{comp}_기기", "") or "")
                vals  = [_fmt(row.get(f"NIR_{comp}_{s}")) for s in grp_samples]
                nir_rows.append([comp, equip] + vals)
            if len(nir_rows) > 1:
                elements.append(_render_nir_table(nir_rows))
                elements.append(Spacer(1, 4*mm))
    else:
        # config에 NIR 그룹이 없을 때: row 키에서 직접 추출
        nir_comps: dict[str, dict] = {}
        for key, val in row.items():
            if not key.startswith("NIR_"):
                continue
            rest = key[4:]
            if rest.endswith("_기기"):
                comp = rest[:-4]
                nir_comps.setdefault(comp, {})["_기기"] = str(val or "")
            else:
                for s in ALL_SAMPLES:
                    if rest.endswith(f"_{s}"):
                        comp = rest[: -len(s) - 1]
                        nir_comps.setdefault(comp, {})[s] = _fmt(val)
                        break

        if nir_comps:
            grp_samples = [s for s in ALL_SAMPLES
                           if any(s in d for d in nir_comps.values())]
            nir_rows = [["성분", "기기명"] + grp_samples]
            for comp, d in nir_comps.items():
                equip = d.get("_기기", "")
                vals  = [d.get(s, "") for s in grp_samples]
                nir_rows.append([comp, equip] + vals)
            if len(nir_rows) > 1:
                elements.append(Paragraph("NIR 측정값", sec_sty))
                elements.append(_render_nir_table(nir_rows))
                elements.append(Spacer(1, 4*mm))

    doc.build(elements)
    return buf.getvalue()
