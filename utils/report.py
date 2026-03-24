import io
from datetime import datetime
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, BaseDocTemplate, PageTemplate, Frame,
    TableOfContents, NextPageTemplate, PageBreak,
    Table, TableStyle, Paragraph, Spacer, HRFlowable,
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


class _SummaryDoc(BaseDocTemplate):
    """BaseDocTemplate subclass that feeds heading paragraphs into the TOC."""
    def afterFlowable(self, flowable):
        if hasattr(flowable, 'style'):
            sn = getattr(flowable.style, 'name', '')
            if sn == '_TOC1':
                self.notify('TOCEntry', (0, flowable.getPlainText(), self.page))
            elif sn == '_TOC2':
                self.notify('TOCEntry', (1, flowable.getPlainText(), self.page))
            elif sn == '_TOC3':
                self.notify('TOCEntry', (2, flowable.getPlainText(), self.page))


def _draw_summary_cover(canvas, doc, main_title: str, subtitle: str, date_str: str, org_str: str):
    """Cover page drawing: colored bars + title box + date/org."""
    w, h = A4
    canvas.saveState()

    # White background
    canvas.setFillColorRGB(1, 1, 1)
    canvas.rect(0, 0, w, h, fill=1, stroke=0)

    # Box dimensions (centered)
    bx = 55 * mm
    bw = w - 110 * mm
    by = h * 0.42
    bh = 65 * mm

    # Blue bar at top
    canvas.setFillColor(colors.HexColor("#1e3a8a"))
    canvas.rect(bx, by + bh - 9 * mm, bw, 9 * mm, fill=1, stroke=0)

    # Orange stripe
    canvas.setFillColor(colors.HexColor("#d97706"))
    canvas.rect(bx, by + bh - 12 * mm, bw, 3 * mm, fill=1, stroke=0)

    # Green stripe
    canvas.setFillColor(colors.HexColor("#16a34a"))
    canvas.rect(bx, by + bh - 14 * mm, bw, 2 * mm, fill=1, stroke=0)

    # Title text area with red dashed border
    ta_y = by
    ta_h = bh - 14 * mm
    canvas.setDash(4, 3)
    canvas.setStrokeColor(colors.HexColor("#ef4444"))
    canvas.setLineWidth(1.2)
    canvas.rect(bx, ta_y, bw, ta_h, fill=0, stroke=1)
    canvas.setDash()

    mid_y = ta_y + ta_h / 2
    canvas.setFillColor(colors.black)
    canvas.setFont(KO, 22)
    if subtitle:
        canvas.drawCentredString(w / 2, mid_y + 7 * mm, main_title)
        canvas.setFont(KO, 15)
        canvas.drawCentredString(w / 2, mid_y - 5 * mm, subtitle)
    else:
        canvas.drawCentredString(w / 2, mid_y - 2 * mm, main_title)

    # Date (blue)
    canvas.setFont(KO, 13)
    canvas.setFillColor(colors.HexColor("#2563eb"))
    canvas.drawCentredString(w / 2, by - 20 * mm, date_str)

    # Organization
    canvas.setFont(KO, 13)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(w / 2, by - 34 * mm, org_str)

    canvas.restoreState()


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
    subtitle: str = "",
    org_name: str = "한국사료협회 사료기술연구소",
    period: str = "",
) -> bytes:
    """전체 요약 보고서: 표지 + 목차 + 개요 + 통계요약 + CV차트 + Z-score 표."""
    import pandas as pd
    import pandas as _pd

    buf = io.BytesIO()
    styles = getSampleStyleSheet()

    # ── 스타일 ──
    h1_style   = ParagraphStyle('_TOC1', fontName=KO, fontSize=13, spaceBefore=10, spaceAfter=4,  leading=16, keepWithNext=1)
    h2_style   = ParagraphStyle('_TOC2', fontName=KO, fontSize=11, spaceBefore=6,  spaceAfter=3,  leading=14, keepWithNext=1)
    h3_style   = ParagraphStyle('_TOC3', fontName=KO, fontSize=9,  spaceBefore=4,  spaceAfter=2,  leading=12, keepWithNext=1, leftIndent=6)
    info_style = ParagraphStyle("is",    fontName=KO, fontSize=9,  spaceAfter=2,   leading=13)
    comp_style = ParagraphStyle("cs",    fontName=KO, fontSize=9,  spaceAfter=2,   spaceBefore=4, leading=12)
    note_style = ParagraphStyle("ns",    fontName=KO, fontSize=7,  spaceAfter=1,   textColor=colors.grey, leftIndent=4)

    def fmt(v, d=2):
        try:
            f = float(v)
            return f"{f:.{d}f}" if not np.isnan(f) else "-"
        except Exception:
            return "-"

    # ── 날짜 파싱 ──
    try:
        _dt = datetime.strptime(generated_at[:10], "%Y-%m-%d")
        date_display = _dt.strftime("%Y.  %m.  %d.")
        _year = _dt.year
        _half = "상반기" if _dt.month <= 6 else "하반기"
    except Exception:
        date_display = generated_at[:10] if generated_at else ""
        _year, _half = datetime.now().year, ""

    # ── 데이터 준비 ──
    avail_mm = 180
    non_nir = [c for c in value_cols if not c.startswith("NIR_")]
    non_nir = [c for c in non_nir if group_stats.get(c)]
    non_nir_set = set(non_nir)

    seen_comps: list = []
    comp_to_cols: dict = {}
    for col in non_nir:
        comp = get_component_from_col(col, samples) or col
        if comp not in seen_comps:
            seen_comps.append(comp)
        comp_to_cols.setdefault(comp, []).append(col)

    valid_stat_samples = [s for s in samples
                          if any(f"{c}_{s}" in non_nir_set for c in seen_comps)]

    _name_to_code = {v: k for k, v in (participant_map or {}).items()}
    raw_inst_names = df[inst_field].fillna("").astype(str).tolist()
    inst_names = [_name_to_code.get(n, n) for n in raw_inst_names]
    idx_list = df.index.tolist()

    # ── 문서 설정 ──
    lm, rm, tm, bm = 15*mm, 15*mm, 18*mm, 18*mm
    pw, ph = A4
    fw, fh = pw - lm - rm, ph - tm - bm

    cover_title = "회원사 비교분석 결과"

    def _cover_bg(canvas, doc):
        _draw_summary_cover(canvas, doc, cover_title, subtitle, date_display, org_name)

    def _content_page(canvas, doc):
        canvas.saveState()
        canvas.setFont(KO, 8)
        canvas.setFillColor(colors.HexColor("#94a3b8"))
        canvas.drawRightString(pw - rm, bm - 8*mm, f"- {doc.page} -")
        canvas.restoreState()

    doc = _SummaryDoc(
        buf, pagesize=A4,
        leftMargin=lm, rightMargin=rm, topMargin=tm, bottomMargin=bm,
    )
    doc.addPageTemplates([
        PageTemplate(id='Cover', frames=[Frame(0, 0, pw, ph, id='cover',
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)],
            onPage=_cover_bg),
        PageTemplate(id='Content', frames=[Frame(lm, bm, fw, fh, id='normal')],
            onPage=_content_page),
    ])

    elements = []

    # ─── 표지 ───
    elements.append(NextPageTemplate('Content'))
    elements.append(PageBreak())

    # ─── 목차 ───
    toc_title_sty = ParagraphStyle("toc_title", fontName=KO, fontSize=16,
                                    alignment=TA_CENTER, spaceAfter=14)
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle('TOCLv1', fontName=KO, fontSize=11, leftIndent=0,  spaceAfter=3, leading=14),
        ParagraphStyle('TOCLv2', fontName=KO, fontSize=10, leftIndent=14, spaceAfter=2, leading=13),
        ParagraphStyle('TOCLv3', fontName=KO, fontSize=9,  leftIndent=28, spaceAfter=1, leading=12),
    ]
    toc.dotsMinLevel = 0
    elements.append(Paragraph("[ 목  차 ]", toc_title_sty))
    elements.append(Spacer(1, 4*mm))
    elements.append(toc)
    elements.append(PageBreak())

    # ─── 1. 비교분석 개요 ───
    elements.append(Paragraph("1. 비교분석 개요", h1_style))

    _period_str = period or f"{_year}년"
    elements.append(Paragraph("가. 기간", h2_style))
    elements.append(Paragraph(f"비교분석 시험 기간: {_period_str}", info_style))
    elements.append(Spacer(1, 2*mm))

    elements.append(Paragraph("나. 사료 및 분석항목", h2_style))
    elements.append(Paragraph(f"대상 사료: {', '.join(samples)}", info_style))
    elements.append(Paragraph(f"분석 항목 ({len(seen_comps)}종): {', '.join(seen_comps)}", info_style))
    elements.append(Spacer(1, 2*mm))

    elements.append(Paragraph("다. 통계처리방법", h2_style))
    elements.append(Paragraph(
        "이상치에 강건한 Robust Z-score (로버스트 Z-점수) 방법을 적용하였습니다.", info_style))
    elements.append(Paragraph(
        "Z = (제출값 − 중앙값) / (1.4826 × MAD),  MAD = 중앙절대편차", note_style))
    elements.append(Paragraph(
        "|Z| ≤ 2.0: 적합  |  2.0 < |Z| ≤ 3.0: 경고  |  |Z| > 3.0: 부적합", note_style))
    elements.append(Spacer(1, 2*mm))

    elements.append(Paragraph("라. 참가회사", h2_style))
    elements.append(Paragraph(f"참가 기관 수: {len(df)}개", info_style))
    _inst_list = sorted(set(inst_names))
    elements.append(Paragraph(f"참가 기관: {', '.join(_inst_list)}", info_style))
    elements.append(PageBreak())

    # ─── 2. 비교분석 결과 ───
    elements.append(Paragraph("2. 비교분석 결과", h1_style))

    # ━━ 가. 분석결과와 통계 ━━
    elements.append(Paragraph("가. 분석결과와 통계", h2_style))

    comp_w_stat = 18
    method_w_stat = 30
    remain_stat = avail_mm - comp_w_stat - method_w_stat
    n_feeds_stat = max(len(valid_stat_samples), 1)
    per_feed_stat = remain_stat / n_feeds_stat
    n_col = 4
    sub_w = per_feed_stat / n_col
    cw_stat = ([comp_w_stat*mm, method_w_stat*mm]
               + [sub_w*mm] * n_col * len(valid_stat_samples))

    cell_s = ParagraphStyle("csp", fontName=KO, fontSize=7, leading=9, alignment=TA_CENTER)
    def _p(txt):
        return Paragraph(str(txt), cell_s)

    hdr0 = [_p("성분"), _p("분석법")]
    for s in valid_stat_samples:
        hdr0 += [_p(s), _p(""), _p(""), _p("")]
    hdr1 = [_p(""), _p("")]
    for _ in valid_stat_samples:
        hdr1 += [_p("N"), _p("평균"), _p("표준편차"), _p("변이계수")]
    stat_rows = [hdr0, hdr1]

    span_cmds = [
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (1, 0), (1, 1)),
    ]
    for i, _ in enumerate(valid_stat_samples):
        c0 = 2 + i * n_col
        span_cmds.append(("SPAN", (c0, 0), (c0 + n_col - 1, 0)))

    whole_rows: set = set()

    for comp in seen_comps:
        method_col = f"{comp}_방법"
        methods: list = []
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
                row += [_p(len(vals)), _p(fmt(mean_)),
                        _p(fmt(std_) if not np.isnan(std_) else "-"),
                        _p(fmt(cv_, 1) if not np.isnan(cv_) else "-")]
            comp_rows_data.append(row)

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
            whole_row += [_p(len(vals)), _p(fmt(mean_)),
                          _p(fmt(std_) if not np.isnan(std_) else "-"),
                          _p(fmt(cv_, 1) if not np.isnan(cv_) else "-")]
        comp_rows_data.append(whole_row)

        if comp_rows_data:
            comp_rows_data[0][0] = _p(comp)
        base_idx = len(stat_rows)
        if len(comp_rows_data) > 1:
            span_cmds.append(("SPAN", (0, base_idx), (0, base_idx + len(comp_rows_data) - 1)))
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
        ("BACKGROUND",    (0, 2), (0, -1),  colors.white),
    ] + span_cmds
    for ri in whole_rows:
        tbl_style_cmds += [("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#e8edf2"))]
    stat_tbl.setStyle(TableStyle(tbl_style_cmds))
    elements.append(stat_tbl)
    elements.append(Spacer(1, 6*mm))

    # ━━ 나. 분석결과 요약 (CV 가로 막대 차트) ━━
    elements.append(Paragraph("나. 분석결과 요약", h2_style))
    elements.append(Paragraph("각 성분별 변이계수(CV%)", info_style))

    cv_data: dict = {}
    for comp in seen_comps:
        for s in valid_stat_samples:
            col = f"{comp}_{s}"
            sd = group_stats.get(col, {})
            try:
                cv_f = float(sd.get("cv", ""))
                if not np.isnan(cv_f):
                    cv_data.setdefault(comp, {})[s] = cv_f
            except Exception:
                pass

    if cv_data:
        _PALETTE_CV = ["#2563eb","#d97706","#16a34a","#dc2626","#7c3aed","#0891b2","#be185d"]
        comps_with_data = [c for c in seen_comps if c in cv_data]
        feeds_with_data = [s for s in valid_stat_samples
                           if any(s in cv_data.get(c, {}) for c in comps_with_data)]
        feed_color = {s: _PALETTE_CV[i % len(_PALETTE_CV)] for i, s in enumerate(feeds_with_data)}

        chart_w_cv = avail_mm * mm
        ml_cv, mr_cv, mt_cv, mb_cv = 62, 8, 10, 22   # pt margins
        plot_w_cv = chart_w_cv - ml_cv - mr_cv

        all_cv_vals = [v for d in cv_data.values() for v in d.values()]
        max_cv_val  = max(all_cv_vals) * 1.15 if all_cv_vals else 10.0
        max_cv_val  = max(max_cv_val, 5.0)

        n_comps_cv = len(comps_with_data)
        n_feeds_cv = len(feeds_with_data)
        bar_h_cv   = 6.5    # pt per single feed bar
        group_h_cv = bar_h_cv * n_feeds_cv + 4   # pt per component group
        plot_h_cv  = group_h_cv * n_comps_cv + 2

        total_h_cv = mt_cv + plot_h_cv + mb_cv
        d_cv = Drawing(chart_w_cv, total_h_cv)
        base_x = ml_cv
        base_y = mb_cv

        # X 눈금선 + 레이블
        n_ticks_cv = 5
        for ti in range(n_ticks_cv + 1):
            tv = max_cv_val * ti / n_ticks_cv
            tx = base_x + (tv / max_cv_val) * plot_w_cv
            d_cv.add(Line(tx, base_y, tx, base_y + plot_h_cv,
                          strokeColor=colors.HexColor("#e5e7eb"), strokeWidth=0.4))
            d_cv.add(GStr(tx, base_y - 10, f"{tv:.0f}",
                          fontSize=5.5, fontName=KO, textAnchor="middle",
                          fillColor=colors.HexColor("#666666")))

        # 축선
        d_cv.add(Line(base_x, base_y, base_x, base_y + plot_h_cv,
                      strokeColor=colors.HexColor("#999"), strokeWidth=0.7))
        d_cv.add(Line(base_x, base_y, base_x + plot_w_cv, base_y,
                      strokeColor=colors.HexColor("#999"), strokeWidth=0.7))

        # 막대 + Y축 레이블 (성분명)
        for ci, comp in enumerate(comps_with_data):
            gy = base_y + ci * group_h_cv + 2
            for fi, feed in enumerate(feeds_with_data):
                cv_val = cv_data.get(comp, {}).get(feed)
                if cv_val is None:
                    continue
                bw = (cv_val / max_cv_val) * plot_w_cv
                by2 = gy + fi * bar_h_cv
                d_cv.add(Rect(base_x, by2, bw, bar_h_cv - 0.8,
                              fillColor=colors.HexColor(feed_color[feed]),
                              strokeColor=None))
            label = comp if len(comp) <= 9 else comp[:8] + "…"
            mid_g = base_y + ci * group_h_cv + group_h_cv / 2 - 2
            d_cv.add(GStr(ml_cv - 3, mid_g, label, fontSize=5.5, fontName=KO,
                          textAnchor="end", fillColor=colors.black))

        # 범례 (차트 내부 우상단)
        legend_box_w = 65
        legend_x = base_x + plot_w_cv - legend_box_w - 2
        legend_y  = base_y + plot_h_cv - 4
        for fi, feed in enumerate(feeds_with_data):
            ly = legend_y - fi * 11
            d_cv.add(Rect(legend_x, ly, 8, 6,
                          fillColor=colors.HexColor(feed_color[feed]),
                          strokeColor=None))
            d_cv.add(GStr(legend_x + 10, ly + 1, feed, fontSize=6.5, fontName=KO,
                          textAnchor="start", fillColor=colors.black))

        elements.append(d_cv)
        elements.append(Spacer(1, 3*mm))

        # 하단 데이터 표
        tbl_hdr = [""] + comps_with_data
        tbl_rows_cv = [tbl_hdr]
        for feed in feeds_with_data:
            row_cv = [feed]
            for comp in comps_with_data:
                cv_val = cv_data.get(comp, {}).get(feed)
                row_cv.append(f"{cv_val:.2f}" if cv_val is not None else "")
            tbl_rows_cv.append(row_cv)
        n_cols_cv  = len(tbl_hdr)
        feed_col_w = 30 * mm
        comp_col_w = (avail_mm * mm - feed_col_w) / max(n_cols_cv - 1, 1)
        tbl_cv = Table(tbl_rows_cv, colWidths=[feed_col_w] + [comp_col_w] * (n_cols_cv - 1))
        tbl_cv_style = TableStyle([
            ("FONTNAME",      (0,0), (-1,-1), KO),
            ("FONTSIZE",      (0,0), (-1,-1), 7),
            ("ALIGN",         (1,0), (-1,-1), "CENTER"),
            ("ALIGN",         (0,0), (0,-1),  "LEFT"),
            ("BACKGROUND",    (0,0), (-1,0),  colors.HexColor("#dbeafe")),
            ("BACKGROUND",    (0,0), (0,-1),  colors.HexColor("#f1f5f9")),
            ("GRID",          (0,0), (-1,-1), 0.4, colors.HexColor("#cbd5e1")),
            ("TOPPADDING",    (0,0), (-1,-1), 2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ])
        for fi, feed in enumerate(feeds_with_data):
            fc = colors.HexColor(feed_color[feed])
            tbl_cv_style.add("TEXTCOLOR", (0, fi+1), (0, fi+1), fc)
        tbl_cv.setStyle(tbl_cv_style)
        elements.append(tbl_cv)

    elements.append(Spacer(1, 8*mm))

    # ━━ Z-score 공통 ━━
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

    def _z_has_data(col, z_src):
        try:
            if isinstance(z_src, _pd.DataFrame):
                return z_src[col].notna().any() if col in z_src.columns else False
            ser = z_src.get(col)
            return ser.notna().any() if ser is not None else False
        except Exception:
            return False

    def _get_zv(z_src, row_idx, col):
        try:
            if isinstance(z_src, _pd.DataFrame):
                return z_src.at[row_idx, col] if col in z_src.columns else np.nan
            ser = z_src.get(col)
            return ser.at[row_idx] if ser is not None else np.nan
        except Exception:
            return np.nan

    inst_w = 45

    def _build_zscore_section(sec_title, sec_heading_style, z_src, heading_prefix):
        elems = []
        elems.append(Paragraph(sec_title, sec_heading_style))
        for num, comp in enumerate(seen_comps, 1):
            sfx_list = [""]
            for n in range(2, 10):
                if any(f"{comp}_{s}_{n}" in non_nir_set for s in samples):
                    sfx_list.append(f"_{n}")
                else:
                    break
            valid_samples = [s for s in samples
                             if any(f"{comp}_{s}{sfx}" in non_nir_set
                                    and _z_has_data(f"{comp}_{s}{sfx}", z_src)
                                    for sfx in sfx_list)]
            if not valid_samples:
                continue
            samp_w = (avail_mm - inst_w) / len(valid_samples)
            cw_z = [inst_w*mm] + [samp_w*mm] * len(valid_samples)
            elems.append(Paragraph(f"{num}) {comp}", h3_style))
            z_rows = [["참가코드"] + valid_samples]
            for inst, row_idx in zip(inst_names, idx_list):
                for sfx in sfx_list:
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
                    sfx_num = sfx.lstrip("_")
                    inst_label = inst if sfx == "" else f"{inst}_{sfx_num}"
                    z_rows.append([inst_label] + row_vals)
            zt = Table(z_rows, colWidths=cw_z, repeatRows=1)
            zt.setStyle(_make_table_style())
            elems.append(zt)
            elems.append(Spacer(1, 4*mm))
        return elems

    elements += _build_zscore_section("다. 시료, 성분별 Robust Z-score", h2_style, z_all, "다")
    elements += _build_zscore_section("라. 방법별 Robust Z-score",        h2_style, z_method, "라")

    # 판정 기준
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
    elements.append(Spacer(1, 3*mm))
    elements.append(Paragraph("<b>판정 기준 (Robust Z-score)</b>", info_style))
    elements.append(Paragraph("적합: |Z| ≤ 2.0   경고: 2.0 < |Z| ≤ 3.0   부적합: |Z| > 3.0", note_style))

    doc.multiBuild(elements)
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
