import io
from datetime import datetime
import numpy as np
from reportlab.lib.pagesizes import A4, landscape as _landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, BaseDocTemplate, PageTemplate, Frame,
    NextPageTemplate, PageBreak,
    Table, TableStyle, Paragraph, Spacer, HRFlowable,
)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Rect, Line
from reportlab.graphics.shapes import String as GStr

from utils.config import get_component_from_col, get_sample_from_col, get_method_options

# 한글 TTF 폰트 등록
import os as _os
_FONT_CANDIDATES = [
    # 로컬/배포: 리포 내 fonts 폴더
    _os.path.join(_os.path.dirname(__file__), "..", "fonts", "KoPub Batang Medium.ttf"),
    # Streamlit Cloud fallback
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
]
KO = "Helvetica"
for _fp in _FONT_CANDIDATES:
    try:
        pdfmetrics.registerFont(TTFont("KoPubBatangM", _fp))
        KO = "KoPubBatangM"
        break
    except Exception:
        continue



def _make_table_style() -> TableStyle:
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTNAME",   (0, 0), (-1, -1), KO),
        ("FONTSIZE",   (0, 0), (-1, -1), 10),
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
        "sec", parent=styles["Heading2"], fontSize=13, spaceAfter=3, spaceBefore=6,
        fontName=KO,
    )
    elements.append(Paragraph(sample_name, section_style))

    def fmt(v, d=2):
        try:
            f = float(v)
            return f"{f:.{d}f}" if not np.isnan(f) else "-"
        except Exception:
            return "-"

    z_cell_style = ParagraphStyle("zcell", fontName=KO, fontSize=10, alignment=TA_CENTER)

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
        method = row_data.get(f"{comp}_방법{sfx}", "") or (inst_method or {}).get(f"{comp}{sfx}", "") or (inst_method or {}).get(comp, "")
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
        "title2", parent=styles["Title"], fontSize=18, spaceAfter=4,
        alignment=TA_CENTER, fontName=KO,
    )
    info_style = ParagraphStyle("info2", parent=styles["Normal"], fontSize=12,
                                spaceAfter=2, fontName=KO)
    note_style = ParagraphStyle("note2", parent=styles["Normal"], fontSize=10,
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

    # 사료 종류별로 그룹핑 (samples 순서 유지)
    sample_to_cols: dict[str, list] = {}
    for col in value_cols:
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
    """Cover page: blue(75%) + green(25%) side-by-side bars top & bottom, title between."""
    w, h = A4
    canvas.saveState()

    canvas.setFillColorRGB(1, 1, 1)
    canvas.rect(0, 0, w, h, fill=1, stroke=0)

    # Title area dimensions
    bx   = 40 * mm
    bw   = w - 80 * mm
    bar_h = 8 * mm
    title_h = 45 * mm          # height between top/bottom bar pairs
    box_top = h * 0.80         # top of top-bar
    box_mid_top = box_top - bar_h
    box_mid_bot = box_mid_top - title_h

    blue_w  = bw * 0.75
    green_w = bw * 0.25

    def _draw_bar_pair(y):
        canvas.setFillColor(colors.HexColor("#4472C4"))
        canvas.rect(bx, y, blue_w, bar_h, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor("#70AD47"))
        canvas.rect(bx + blue_w, y, green_w, bar_h, fill=1, stroke=0)

    _draw_bar_pair(box_mid_top)   # 위쪽 막대 쌍
    _draw_bar_pair(box_mid_bot)   # 아래쪽 막대 쌍

    # Title text (between bars)
    mid_y = box_mid_bot + title_h / 2
    canvas.setFillColor(colors.black)
    if subtitle:
        canvas.setFont(KO, 22)
        canvas.drawCentredString(w / 2, mid_y + 6 * mm, main_title)
        canvas.setFont(KO, 15)
        canvas.drawCentredString(w / 2, mid_y - 6 * mm, subtitle)
    else:
        canvas.setFont(KO, 22)
        canvas.drawCentredString(w / 2, mid_y - 2 * mm, main_title)

    # Date (blue, below bottom bar)
    canvas.setFont(KO, 15)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(w / 2, box_mid_bot - 80 * mm, date_str)

    # Organization
    canvas.setFont(KO, 20)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(w / 2, box_mid_bot - 100 * mm, org_str)

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
    period_배부: str = "",
    period_회신: str = "",
    period_보고서: str = "",
    sample_note: str = "",
    summary_text: str = "",
    cfg=None,
) -> bytes:
    """전체 요약 보고서: 표지 + 목차 + 개요 + 통계요약 + CV차트 + Z-score 표."""
    import pandas as pd
    import pandas as _pd

    buf = io.BytesIO()
    styles = getSampleStyleSheet()

    # ── 스타일 ──
    h1_style   = ParagraphStyle('_TOC1', fontName=KO, fontSize=15, spaceBefore=10, spaceAfter=10, leading=18, keepWithNext=1)
    h2_style   = ParagraphStyle('_TOC2', fontName=KO, fontSize=13, spaceBefore=8,  spaceAfter=10, leading=16, keepWithNext=1)
    h3_style   = ParagraphStyle('_TOC3', fontName=KO, fontSize=11, spaceBefore=6,  spaceAfter=10, leading=14, keepWithNext=1, leftIndent=6)
    info_style = ParagraphStyle("is",    fontName=KO, fontSize=11, spaceAfter=2,   leading=15)
    comp_style = ParagraphStyle("cs",    fontName=KO, fontSize=11, spaceAfter=2,   spaceBefore=4, leading=14)
    note_style = ParagraphStyle("ns",    fontName=KO, fontSize=9,  spaceAfter=1,   textColor=colors.grey, leftIndent=4)

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
    except Exception:
        date_display = generated_at[:10] if generated_at else ""

    # ── 데이터 준비 ──
    avail_mm = 180
    non_nir = [c for c in value_cols if group_stats.get(c)]
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
    lm, rm, tm, bm = 20*mm, 20*mm, 25*mm, 20*mm
    pw, ph = A4
    fw, fh = pw - lm - rm, ph - tm - bm

    # 가로 페이지 (landscape) 설정
    pw_land, ph_land = _landscape(A4)
    fw_land = pw_land - lm - rm
    fh_land = ph_land - tm - bm
    avail_land_mm = fw_land / mm   # ~267mm

    cover_title = "한국사료협회 비교분석 결과"

    def _cover_bg(canvas, doc):
        _draw_summary_cover(canvas, doc, cover_title, subtitle, date_display, org_name)

    def _content_page(canvas, doc):
        canvas.saveState()
        canvas.setFont(KO, 8)
        canvas.setFillColor(colors.HexColor("#94a3b8"))
        canvas.drawRightString(pw - rm, bm - 8*mm, f"- {doc.page} -")
        canvas.restoreState()

    def _land_page(canvas, doc):
        canvas.saveState()
        canvas.setFont(KO, 8)
        canvas.setFillColor(colors.HexColor("#94a3b8"))
        canvas.drawRightString(pw_land - rm, bm - 8*mm, f"- {doc.page} -")
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
        PageTemplate(id='Landscape', frames=[Frame(lm, bm, fw_land, fh_land, id='land')],
            onPage=_land_page, pagesize=_landscape(A4)),
    ])

    elements = []

    # ─── 표지 ───
    elements.append(NextPageTemplate('Content'))
    elements.append(PageBreak())

    # ─── 목차 ───
    toc_title_sty = ParagraphStyle("toc_title", fontName=KO, fontSize=18,
                                    alignment=TA_CENTER, spaceAfter=14)
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle('TOCLv1', fontName=KO, fontSize=13, leftIndent=0,  spaceAfter=3, leading=16),
        ParagraphStyle('TOCLv2', fontName=KO, fontSize=12, leftIndent=14, spaceAfter=2, leading=15),
        ParagraphStyle('TOCLv3', fontName=KO, fontSize=11, leftIndent=28, spaceAfter=1, leading=14),
    ]
    toc.dotsMinLevel = 0
    elements.append(Paragraph("[ 목  차 ]", toc_title_sty))
    elements.append(Spacer(1, 4*mm))
    elements.append(toc)
    elements.append(PageBreak())

    # ─── 1. 비교분석 개요 ───
    elements.append(Paragraph("1. 비교분석 개요", h1_style))

    # 가. 기간
    elements.append(Paragraph("가. 기간", h2_style))
    elements.append(Paragraph(f"1) 시료 배부 : {period_배부}", info_style))
    elements.append(Paragraph(f"2) 분석 및 결과회신 : {period_회신}", info_style))
    elements.append(Paragraph(f"3) 결과 통계처리 및 보고서 작성 : {period_보고서}", info_style))
    elements.append(Spacer(1, 8*mm))

    # 나. 시료 및 분석항목 (표 형식)
    elements.append(Paragraph("나. 시료 및 분석항목", h2_style))
    # 사료별 성분 매핑
    sample_to_comps: dict = {}
    for comp in seen_comps:
        for col in comp_to_cols.get(comp, []):
            s = get_sample_from_col(col, samples)
            if s and s in samples:
                sample_to_comps.setdefault(s, [])
                if comp not in sample_to_comps[s]:
                    sample_to_comps[s].append(comp)
    cell_ov = ParagraphStyle("cov", fontName=KO, fontSize=10, leading=13)
    ov_rows = [[Paragraph("<b>시료</b>", cell_ov), Paragraph("<b>분석항목</b>", cell_ov)]]
    for i, s in enumerate(samples):
        comps_s = sample_to_comps.get(s, [])
        if not comps_s:
            continue
        ov_rows.append([
            Paragraph(f"{s}(샘플{i+1})", cell_ov),
            Paragraph(", ".join(comps_s), cell_ov),
        ])
    ov_tbl = Table(ov_rows, colWidths=[45*mm, 135*mm])
    ov_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  colors.HexColor("#83aaff")),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.white),
        ("FONTNAME",      (0, 0), (-1, -1), KO),
        ("FONTSIZE",      (0, 0), (-1, -1), 10),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#000000")),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#FFFFFF")]),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]))
    elements.append(ov_tbl)
    if sample_note:
        elements.append(Paragraph(f"* {sample_note}", note_style))
    elements.append(Spacer(1, 8*mm))

    # 다. 통계처리방법
    elements.append(Paragraph("다. 통계처리방법", h2_style))
    elements.append(Paragraph(
        "1) 평균(x̄), 표준편차(σ), 변이계수(CV) 등을 산출함.", info_style))
    elements.append(Paragraph(
        "2) 시험소 간 비교숙련도 시험용 Robust Z-score, Outlier : KS Q ISO 13528(통계적방법)에 따라 다음과 같이 계산하여 해석함.",
        info_style))
    elements.append(Spacer(1, 2*mm))
    # 수식 표 (분수 형태)
    formula_cell = ParagraphStyle("fc", fontName=KO, fontSize=10.5, alignment=TA_CENTER, leading=14)
    f_tbl = Table([
        [Paragraph("Robust Z-score  =", formula_cell),
         Paragraph("결과값(Result) − 중위수(Median)", formula_cell),
         Paragraph("", formula_cell)],
        [Paragraph("", formula_cell),
         Paragraph("정규화된 사분위범위(Normalized IQR)", formula_cell),
         Paragraph("", formula_cell)],
    ], colWidths=[55*mm, 100*mm, 25*mm])
    f_tbl.setStyle(TableStyle([
        ("FONTNAME",      (0,0), (-1,-1), KO),
        ("FONTSIZE",      (0,0), (-1,-1), 13),
        ("ALIGN",         (0,0), (-1,-1), "CENTER"),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("LINEBELOW",     (1,0), (1,0),   0.8, colors.black),  # 분수선
        ("SPAN",          (0,0), (0,1)),  # "Robust Z-score =" 세로 병합
        ("TOPPADDING",    (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    elements.append(f_tbl)
    elements.append(Spacer(1, 2*mm))
    elements.append(Paragraph("- Median : 결과를 크기순으로 나열했을 때 중간값", note_style))
    elements.append(Paragraph("- Normalized IQR : 제3사분위수(Q3)에서 제1사분위수(Q1)를 뺀 값 × 0.7413", note_style))
    elements.append(Spacer(1, 2*mm))
    crit_data = [
        ["|Z| ≤ 2", ": 만족(Satisfactory)"],
        ["2 < |Z| < 3", ": 의심(Doubt)"],
        ["3 ≤ |Z|", ": 불만족(Outlier)"],
    ]
    crit_tbl = Table(crit_data, colWidths=[35*mm, 60*mm], hAlign="CENTER")
    crit_tbl.setStyle(TableStyle([
        ("FONTNAME",  (0,0), (-1,-1), KO),
        ("FONTSIZE",  (0,0), (-1,-1), 10.5),
        ("ALIGN",     (0,0), (-1,-1), "CENTER"),
        ("TOPPADDING",(0,0), (-1,-1), 1),
        ("BOTTOMPADDING",(0,0), (-1,-1), 1),
    ]))
    elements.append(crit_tbl)
    elements.append(Spacer(1, 8*mm))

    # 라. 참가회원사 (실제 기관명 사용)
    elements.append(Paragraph("라. 참가회원사", h2_style))
    _inst_real = sorted(set(raw_inst_names))
    elements.append(Paragraph(", ".join(_inst_real), info_style))

    # ─── 2. 비교분석 결과 (가로 페이지 시작) ───
    elements.append(NextPageTemplate('Landscape'))
    elements.append(PageBreak())
    elements.append(Paragraph("2. 비교분석 결과", h1_style))

    # ━━ 가. 분석결과와 통계 ━━
    elements.append(Paragraph("가. 분석결과와 통계", h2_style))

    comp_w_stat = 20
    method_w_stat = 35
    remain_stat = avail_land_mm - comp_w_stat - method_w_stat
    n_feeds_stat = max(len(valid_stat_samples), 1)
    per_feed_stat = remain_stat / n_feeds_stat
    n_col = 4
    sub_w = per_feed_stat / n_col
    cw_stat = ([comp_w_stat*mm, method_w_stat*mm]
               + [sub_w*mm] * n_col * len(valid_stat_samples))

    cell_s = ParagraphStyle("csp", fontName=KO, fontSize=9, leading=11, alignment=TA_CENTER)
    def _p(txt):
        s = str(txt)
        return Paragraph(s, cell_s) if s.strip() else ""

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
    dead_style_cmds: list = []

    # 데이터가 전혀 없는 (성분, 사료) 조합 사전 탐지 → 대각선 처리용
    def _is_dead(comp, s):
        for n in range(1, 10):
            sfx = "" if n == 1 else f"_{n}"
            if f"{comp}_{s}{sfx}" in non_nir_set:
                return False
        return True

    dead_cell_s = ParagraphStyle("dcp", fontName=KO, fontSize=11, alignment=TA_CENTER,
                                 textColor=colors.HexColor("#bbbbbb"))

    row_cursor = len(stat_rows)  # 헤더 2행 다음부터

    for comp in seen_comps:
        method_entries = []
        for n in range(1, 10):
            sfx = "" if n == 1 else f"_{n}"
            mc  = f"{comp}_방법{sfx}"
            if mc not in df.columns:
                if n > 1:
                    break
                continue
            for m in df[mc].dropna().astype(str).unique():
                if m.strip():
                    method_entries.append((mc, sfx, m.strip()))

        comp_rows_data = []

        def _fill_sample_cells(row, sfx, mc=None, meth=None):
            for si, s in enumerate(valid_stat_samples):
                c0 = 2 + si * n_col
                abs_row = row_cursor + len(comp_rows_data)
                if _is_dead(comp, s):
                    row += [Paragraph("╱", dead_cell_s), _p(""), _p(""), _p("")]
                    dead_style_cmds.append(("SPAN", (c0, abs_row), (c0+3, abs_row)))
                    dead_style_cmds.append(("BACKGROUND", (c0, abs_row), (c0+3, abs_row),
                                            colors.HexColor("#e8e8e8")))
                    continue
                col = f"{comp}_{s}{sfx}"
                if col not in df.columns:
                    row += [_p("")]*n_col; continue
                if mc and meth:
                    mask = (df[mc].fillna("").astype(str).str.strip() == meth)
                    vals = pd.to_numeric(df.loc[mask, col], errors="coerce").dropna()
                else:
                    # 전체 행: 모든 suffix 풀링
                    all_vals = []
                    for n2 in range(1, 10):
                        sx2 = "" if n2 == 1 else f"_{n2}"
                        c2  = f"{comp}_{s}{sx2}"
                        if c2 not in df.columns:
                            if n2 > 1: break
                            continue
                        all_vals.extend(pd.to_numeric(df[c2], errors="coerce").dropna().tolist())
                    vals = pd.Series(all_vals)
                if len(vals) == 0:
                    row += [_p("")]*n_col; continue
                mean_ = vals.mean()
                std_  = vals.std(ddof=1) if len(vals) > 1 else float("nan")
                cv_   = (std_/mean_*100) if mean_ != 0 and not np.isnan(std_) else float("nan")
                row += [_p(len(vals)), _p(fmt(mean_)),
                        _p(fmt(std_) if not np.isnan(std_) else "-"),
                        _p(fmt(cv_, 1) if not np.isnan(cv_) else "-")]

        for mc, sfx, meth in method_entries:
            row = [_p(""), _p(meth)]
            _fill_sample_cells(row, sfx, mc, meth)
            comp_rows_data.append(row)

        # 전체 행
        whole_row = [_p(""), _p(f"{comp} 전체")]
        _fill_sample_cells(whole_row, "")
        comp_rows_data.append(whole_row)

        if comp_rows_data:
            comp_rows_data[0][0] = _p(comp)
        base_idx = len(stat_rows)
        if len(comp_rows_data) > 1:
            span_cmds.append(("SPAN", (0, base_idx), (0, base_idx + len(comp_rows_data) - 1)))
        whole_rows.add(base_idx + len(comp_rows_data) - 1)
        row_cursor += len(comp_rows_data)
        stat_rows.extend(comp_rows_data)

    stat_tbl = Table(stat_rows, colWidths=cw_stat, repeatRows=2)
    tbl_style_cmds = [
        ("BACKGROUND",    (0, 0), (-1, 1),  colors.HexColor("#729dc9")),
        ("TEXTCOLOR",     (0, 0), (-1, 1),  colors.white),
        ("FONTNAME",      (0, 0), (-1, -1), KO),
        ("FONTSIZE",      (0, 0), (-1, -1), 10),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",          (0, 0), (-1, -1), 0.4, colors.black),
        ("TOPPADDING",    (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING",   (0, 0), (-1, -1), 2),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS",(1, 2), (-1, -1), [colors.white, colors.HexColor("#f0f0f0")]),
        ("BACKGROUND",    (0, 2), (0, -1),  colors.white),
    ] + span_cmds + dead_style_cmds
    for ri in whole_rows:
        tbl_style_cmds += [("BACKGROUND", (0, ri), (-1, ri), colors.HexColor("#ffffff"))]
    stat_tbl.setStyle(TableStyle(tbl_style_cmds))
    elements.append(stat_tbl)

    # 가로 페이지 종료 → 세로 페이지로 복귀
    elements.append(NextPageTemplate('Content'))
    elements.append(PageBreak())

    # ━━ 나. 분석결과 요약 (CV 가로 막대 차트) ━━
    elements.append(Paragraph("나. 분석결과 요약", h2_style))

    cv_data: dict = {}
    for comp in seen_comps:
        for s in valid_stat_samples:
            col = f"{comp}_{s}"
            sd = group_stats.get(col, {})
            if not sd:
                continue
            try:
                cv_f = float(sd.get("cv", float("nan")))
            except Exception:
                cv_f = float("nan")
            cv_data.setdefault(comp, {})[s] = cv_f

    if cv_data:
        _PALETTE_CV = ["#2563eb","#d97706","#16a34a","#dc2626","#7c3aed","#0891b2","#be185d"]
        comps_with_data = [c for c in seen_comps if c in cv_data]
        feeds_with_data = [s for s in valid_stat_samples
                           if any(s in cv_data.get(c, {}) for c in comps_with_data)]
        feed_color = {s: _PALETTE_CV[i % len(_PALETTE_CV)] for i, s in enumerate(feeds_with_data)}

        chart_w_cv = avail_mm * mm
        ml_cv, mr_cv, mt_cv, mb_cv = 36, 12, 28, 20
        plot_w_cv = chart_w_cv - ml_cv - mr_cv
        plot_h_cv = 120

        all_cv_vals = [v for d in cv_data.values() for v in d.values()
                       if not (isinstance(v, float) and np.isnan(v))]
        raw_max = max(all_cv_vals) if all_cv_vals else 10.0
        # Y축 max를 10 단위로 올림
        import math as _math
        n_ticks_cv = 9
        tick_step = _math.ceil(raw_max * 1.1 / n_ticks_cv / 10) * 10 or 10
        max_cv_val = tick_step * n_ticks_cv

        n_comps_cv = len(comps_with_data)
        n_feeds_cv = len(feeds_with_data)
        group_w    = plot_w_cv / max(n_comps_cv, 1)
        bar_w_cv   = (group_w - 4) / max(n_feeds_cv, 1)

        total_h_cv = mt_cv + plot_h_cv + mb_cv
        d_cv = Drawing(chart_w_cv, total_h_cv)
        base_x = ml_cv
        base_y = mb_cv

        # 차트 배경 + 테두리
        d_cv.add(Rect(base_x, base_y, plot_w_cv, plot_h_cv,
                      fillColor=colors.white,
                      strokeColor=colors.HexColor("#aaaaaa"), strokeWidth=0.5))

        # Y 눈금선 + 레이블 (.2f 형식)
        for ti in range(n_ticks_cv + 1):
            tv = tick_step * ti
            ty = base_y + (tv / max_cv_val) * plot_h_cv
            d_cv.add(Line(base_x, ty, base_x + plot_w_cv, ty,
                          strokeColor=colors.HexColor("#dddddd"), strokeWidth=0.4))
            d_cv.add(GStr(base_x - 3, ty - 3, f"{tv:.2f}",
                          fontSize=7, fontName=KO, textAnchor="end",
                          fillColor=colors.HexColor("#555555")))

        # 세로 막대 + X축 레이블 (직선)
        for ci, comp in enumerate(comps_with_data):
            gx = base_x + ci * group_w + 2
            for fi, feed in enumerate(feeds_with_data):
                cv_val = cv_data.get(comp, {}).get(feed)
                if cv_val is None or (isinstance(cv_val, float) and np.isnan(cv_val)):
                    continue
                bx = gx + fi * bar_w_cv
                bh = max((cv_val / max_cv_val) * plot_h_cv, 0)
                d_cv.add(Rect(bx, base_y, max(bar_w_cv - 1, 1), bh,
                              fillColor=colors.HexColor(feed_color[feed]),
                              strokeColor=None))
            label = comp if len(comp) <= 8 else comp[:7] + "…"
            lx = base_x + ci * group_w + group_w / 2
            d_cv.add(GStr(lx, base_y - 12, label,
                          fontSize=7.5, fontName=KO, textAnchor="middle",
                          fillColor=colors.black))

        # 제목
        d_cv.add(GStr(base_x + plot_w_cv / 2, base_y + plot_h_cv + 14,
                      "각 성분별 변이계수(CV%)",
                      fontSize=11, fontName=KO, textAnchor="middle",
                      fillColor=colors.black))

        # 범례 (차트 좌상단 안쪽)
        leg_x = base_x + 6
        leg_y = base_y + plot_h_cv - 10
        for fi, feed in enumerate(feeds_with_data):
            ly = leg_y - fi * 11
            d_cv.add(Rect(leg_x, ly, 8, 6,
                          fillColor=colors.HexColor(feed_color[feed]),
                          strokeColor=None))
            d_cv.add(GStr(leg_x + 10, ly + 0.5, feed, fontSize=7.5, fontName=KO,
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
                if cv_val is not None and not (isinstance(cv_val, float) and np.isnan(cv_val)):
                    row_cv.append(f"{cv_val:.2f}")
                else:
                    row_cv.append("")
            tbl_rows_cv.append(row_cv)
        n_cols_cv  = len(tbl_hdr)
        feed_col_w = 30 * mm
        comp_col_w = (avail_mm * mm - feed_col_w) / max(n_cols_cv - 1, 1)
        tbl_cv = Table(tbl_rows_cv, colWidths=[feed_col_w] + [comp_col_w] * (n_cols_cv - 1))
        tbl_cv_style = TableStyle([
            ("FONTNAME",      (0,0), (-1,-1), KO),
            ("FONTSIZE",      (0,0), (-1,-1), 10),
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

    if summary_text:
        elements.append(Spacer(1, 4*mm))
        for line in summary_text.splitlines():
            line = line.strip()
            if line:
                elements.append(Paragraph(line, info_style))

    elements.append(Spacer(1, 8*mm))


    def _get_zv(z_src, row_idx, col):
        try:
            if isinstance(z_src, _pd.DataFrame):
                return z_src.at[row_idx, col] if col in z_src.columns else np.nan
            ser = z_src.get(col)
            return ser.at[row_idx] if ser is not None else np.nan
        except Exception:
            return np.nan

    def _build_zscore_section(sec_title, sec_heading_style, z_src, heading_prefix, min_n=1, split_at=0):
        elems = []
        elems.append(Paragraph(sec_title, sec_heading_style))

        hdr_s  = ParagraphStyle("zh", fontName=KO, fontSize=8, alignment=TA_CENTER,
                                 textColor=colors.white, leading=10)
        cell_s2 = ParagraphStyle("zc", fontName=KO, fontSize=8, alignment=TA_CENTER, leading=10)
        red_s   = ParagraphStyle("zr", fontName=KO, fontSize=8, alignment=TA_CENTER, leading=10,
                                 textColor=colors.HexColor("#dc2626"))
        org_s   = ParagraphStyle("zo", fontName=KO, fontSize=8, alignment=TA_CENTER, leading=10,
                                 textColor=colors.HexColor("#d97706"))

        def _hp(txt): return Paragraph(str(txt), hdr_s)
        def _cp(txt): return Paragraph(str(txt), cell_s2)

        def _zc(z_val):
            try:
                z_f = float(z_val)
                if np.isnan(z_f): return _cp("-")
            except Exception:
                return _cp("-")
            z_str = f"{z_f:.2f}"
            if abs(z_f) > 3:
                return Paragraph(f"<b><u>{z_str}</u></b>", red_s)
            elif abs(z_f) > 2:
                return Paragraph(f"<b>{z_str}</b>", org_s)
            return _cp(z_str)

        for num, comp in enumerate(seen_comps, 1):
            sfx_list = [""]
            for n in range(2, 10):
                if any(f"{comp}_{s}_{n}" in non_nir_set for s in samples):
                    sfx_list.append(f"_{n}")
                else:
                    break

            valid_samples = [s for s in samples
                             if any(f"{comp}_{s}{sfx}" in non_nir_set for sfx in sfx_list)]
            if not valid_samples:
                continue

            # 컬럼 폭: 분석방법 | Lab | (결과, Z-score) × n_samples
            meth_w = 30
            lab_w  = 14
            n_samp = len(valid_samples)
            per_s  = (avail_mm - meth_w - lab_w) / n_samp
            res_w  = per_s * 0.42
            z_w    = per_s * 0.58
            cw_z   = [meth_w*mm, lab_w*mm] + [res_w*mm, z_w*mm] * n_samp

            # 헤더 1행: 분석방법 | Lab | 사료명(중간값) [2열 span] ...
            hdr1 = [_hp("분석방법"), _hp("Lab")]
            for s in valid_samples:
                col0 = f"{comp}_{s}"
                gs   = group_stats.get(col0, {})
                try:
                    med = float(gs.get("median", float("nan")))
                    med_str = f"{med:.2f}" if not np.isnan(med) else "-"
                except Exception:
                    med_str = "-"
                hdr1 += [_hp(f"{s}  (중간값: {med_str})"), _hp("")]

            # 헤더 2행: '' | '' | 결과 | Z-score | ...
            hdr2 = [_hp(""), _hp("")]
            for _ in valid_samples:
                hdr2 += [_hp("결과"), _hp("Z-score")]

            z_rows   = [hdr1, hdr2]
            span_z   = [("SPAN", (0, 0), (0, 1)), ("SPAN", (1, 0), (1, 1))]
            for si in range(n_samp):
                span_z.append(("SPAN", (2 + si*2, 0), (3 + si*2, 0)))

            # 방법별 그룹핑: {method_name: [(inst_label, row_vals), ...]}
            method_to_rows: dict = {}
            for sfx in sfx_list:
                mc = f"{comp}_방법{sfx}"
                for inst, row_idx in zip(inst_names, idx_list):
                    meth = str(df.at[row_idx, mc] if mc in df.columns else "").strip()
                    if not meth:
                        meth = "(방법 미기재)"
                    row_vals = []
                    has_any  = False
                    for s in valid_samples:
                        col = f"{comp}_{s}{sfx}"
                        if col not in non_nir_set:
                            row_vals += [_cp("-"), _cp("-")]; continue
                        raw = df.at[row_idx, col] if col in df.columns else ""
                        try:
                            rv = float(raw)
                            res_str = f"{rv:.2f}" if not np.isnan(rv) else "-"
                            if not np.isnan(rv): has_any = True
                        except Exception:
                            res_str = "-"
                        zv = _get_zv(z_src, row_idx, col)
                        row_vals += [_cp(res_str), _zc(zv)]
                    if not has_any:
                        continue
                    lab_label = inst if sfx == "" else f"{inst}{sfx}"
                    method_to_rows.setdefault(meth, []).append(
                        [_cp(""), _cp(lab_label)] + row_vals
                    )

            def _lab_sort_key(row):
                try: return (0, int(row[1].text))
                except Exception: return (1, str(row[1].text))

            ordered_methods = get_method_options(cfg, comp=comp) if cfg is not None else []
            def _meth_order(item):
                m = item[0]
                try: return ordered_methods.index(m)
                except ValueError: return len(ordered_methods)

            # n 부족 방법 제거
            if min_n > 1:
                method_to_rows = {m: r for m, r in method_to_rows.items() if len(r) >= min_n}

            if not method_to_rows:
                continue

            sorted_methods = sorted(method_to_rows.items(), key=_meth_order)
            total_data_rows = sum(len(r) for _, r in sorted_methods)
            do_split = split_at > 0 and total_data_rows > split_at

            def _make_zt(rows, spans):
                t = Table(rows, colWidths=cw_z, repeatRows=2)
                t.setStyle(TableStyle([
                    ("BACKGROUND",    (0, 0), (-1, 1),  colors.HexColor("#4472C4")),
                    ("TEXTCOLOR",     (0, 0), (-1, 1),  colors.white),
                    ("FONTNAME",      (0, 0), (-1, -1), KO),
                    ("FONTSIZE",      (0, 0), (-1, -1), 8),
                    ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
                    ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
                    ("TOPPADDING",    (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LEFTPADDING",   (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
                    ("ROWBACKGROUNDS",(0, 2), (-1, -1), [colors.white, colors.HexColor("#f0f4fa")]),
                    ("BACKGROUND",    (0, 2), (0, -1),  colors.white),
                    ("BACKGROUND",    (1, 2), (1, -1),  colors.HexColor("#f8fafc")),
                ] + spans))
                return t

            elems.append(Paragraph(f"{num}) {comp}", h3_style))

            if do_split:
                # 방법별 개별 표
                meth_sub_style = ParagraphStyle("ms", fontName=KO, fontSize=9,
                                                spaceBefore=4, spaceAfter=3,
                                                textColor=colors.HexColor("#1e3a8a"))
                for meth, mrows in sorted_methods:
                    mrows.sort(key=_lab_sort_key)
                    sub_rows  = [hdr1, hdr2]
                    sub_spans = [("SPAN", (0, 0), (0, 1)), ("SPAN", (1, 0), (1, 1))]
                    for si in range(n_samp):
                        sub_spans.append(("SPAN", (2 + si*2, 0), (3 + si*2, 0)))
                    mrows[0][0] = _cp(meth)
                    sub_rows.extend(mrows)
                    if len(mrows) > 1:
                        sub_spans.append(("SPAN", (0, 2), (0, 1 + len(mrows))))
                    elems.append(Paragraph(f"▶ {meth}", meth_sub_style))
                    elems.append(_make_zt(sub_rows, sub_spans))
                    elems.append(Spacer(1, 3*mm))
            else:
                # 기존: 하나의 통합 표
                for meth, mrows in sorted_methods:
                    mrows.sort(key=_lab_sort_key)
                    mrows[0][0] = _cp(meth)
                    mstart = len(z_rows)
                    z_rows.extend(mrows)
                    if len(mrows) > 1:
                        span_z.append(("SPAN", (0, mstart), (0, mstart + len(mrows) - 1)))

                if len(z_rows) <= 2:
                    continue
                elems.append(_make_zt(z_rows, span_z))
                elems.append(Spacer(1, 4*mm))
        return elems

    def _build_zscore_charts(z_src, group_by_method=False, min_n=1):
        """Z-score 정렬 막대그래프 생성 (다/라 공용)"""
        import math as _m
        elems = []
        chart_w = avail_mm * mm
        ml, mr  = 40, 10          # pt  y축 레이블 / 우측 여백
        mt_pad  = 24              # pt  제목 공간
        mb_pad  = 38              # pt  x 레이블 + z값 공간
        plot_w  = chart_w - ml - mr
        plot_h  = 88              # pt

        COL_BLUE = colors.HexColor("#4472C4")
        COL_RED  = colors.HexColor("#dc2626")
        COL_GRID = colors.HexColor("#dddddd")
        COL_ZERO = colors.HexColor("#888888")
        COL_TXT  = colors.HexColor("#444444")

        def _sol_abbr(meth):
            m = str(meth).upper()
            if "PETROLEUM" in m: return "P"
            if "DIETHYL"   in m: return "DE"
            if "ETHER"     in m: return "E"
            if "에테르"    in m: return "E"
            return ""

        is_fat = lambda c: "조지방" in str(c)

        for comp in seen_comps:
            sfx_list = [""]
            for n in range(2, 10):
                if any(f"{comp}_{s}_{n}" in non_nir_set for s in samples):
                    sfx_list.append(f"_{n}")
                else:
                    break

            valid_samps = [s for s in samples
                           if any(f"{comp}_{s}{sfx}" in non_nir_set for sfx in sfx_list)]
            if not valid_samps:
                continue

            for s in valid_samps:
                grouped: dict = {}

                for sfx in sfx_list:
                    col = f"{comp}_{s}{sfx}"
                    if col not in non_nir_set:
                        continue
                    mc = f"{comp}_방법{sfx}"
                    for inst, row_idx in zip(inst_names, idx_list):
                        zv = _get_zv(z_src, row_idx, col)
                        try:
                            zv = float(zv)
                        except Exception:
                            continue
                        if np.isnan(zv):
                            continue

                        meth_str = ""
                        if mc in df.columns:
                            try:
                                meth_str = str(df.at[row_idx, mc]).strip()
                            except Exception:
                                pass
                        if not meth_str or meth_str == "nan":
                            meth_str = "(방법 미기재)"

                        grp_key = meth_str if group_by_method else "__ALL__"

                        if is_fat(comp):
                            abbr  = _sol_abbr(meth_str)
                            label = f"{inst}-{abbr}" if abbr else str(inst)
                        else:
                            label = str(inst)

                        grouped.setdefault(grp_key, []).append((label, zv))

                grouped = {k: v for k, v in grouped.items() if len(v) >= min_n}
                if not grouped:
                    continue

                ordered_methods = get_method_options(cfg, comp=comp) if cfg is not None else []

                def _gk_order(k, _om=ordered_methods):
                    try:    return _om.index(k)
                    except: return len(_om)

                for grp_key, bars in sorted(grouped.items(), key=lambda x: _gk_order(x[0])):
                    bars.sort(key=lambda x: x[1])
                    n_bars = len(bars)

                    title = (f"{s}({comp}-{grp_key})"
                             if group_by_method and grp_key not in ("__ALL__", "(방법 미기재)")
                             else f"{s}({comp})")

                    z_vals  = [z for _, z in bars]
                    abs_max = max(abs(min(z_vals)), abs(max(z_vals)), 2.5)
                    y_hi    = float(_m.ceil(abs_max * 1.15))
                    y_lo    = -y_hi
                    y_span  = y_hi - y_lo

                    gap   = max(0.8, plot_w * 0.008)
                    bar_w = max((plot_w - gap * (n_bars + 1)) / n_bars, 3.0)
                    lbl_fs = max(5.0, 7.5 - max(0, n_bars - 20) * 0.1)
                    z_fs   = max(4.5, 6.5 - max(0, n_bars - 20) * 0.1)

                    total_h = mt_pad + plot_h + mb_pad
                    d = Drawing(chart_w, total_h)

                    # 플롯 배경
                    d.add(Rect(ml, mb_pad, plot_w, plot_h,
                               fillColor=colors.white,
                               strokeColor=colors.HexColor("#aaaaaa"), strokeWidth=0.5))

                    # Y축 눈금 (1 단위)
                    for tv in range(int(y_lo), int(y_hi) + 1):
                        ty = mb_pad + (tv - y_lo) / y_span * plot_h
                        lw = 0.7 if tv == 0 else 0.3
                        lc = COL_ZERO if tv == 0 else COL_GRID
                        d.add(Line(ml, ty, ml + plot_w, ty,
                                   strokeColor=lc, strokeWidth=lw))
                        d.add(GStr(ml - 3, ty - 3.5, f"{tv:.2f}",
                                   fontSize=6.5, fontName=KO, textAnchor="end",
                                   fillColor=COL_TXT))

                    # ±3 기준선 (빨간 점선)
                    for ref in [3.0, -3.0]:
                        if y_lo <= ref <= y_hi:
                            ry = mb_pad + (ref - y_lo) / y_span * plot_h
                            d.add(Line(ml, ry, ml + plot_w, ry,
                                       strokeColor=COL_RED, strokeWidth=0.6,
                                       strokeDashArray=[3, 2]))

                    # 막대
                    zero_y = mb_pad + (0 - y_lo) / y_span * plot_h
                    for i, (lbl, zv) in enumerate(bars):
                        bx = ml + gap + i * (bar_w + gap)
                        bc = COL_RED if abs(zv) >= 3 else COL_BLUE
                        bh = abs(zv) / y_span * plot_h
                        by = zero_y if zv >= 0 else zero_y - bh
                        d.add(Rect(bx, by, bar_w, max(bh, 0.5),
                                   fillColor=bc, strokeColor=None))
                        cx = bx + bar_w / 2
                        # x축 레이블 (lab 코드)
                        d.add(GStr(cx, mb_pad - 11, str(lbl),
                                   fontSize=lbl_fs, fontName=KO, textAnchor="middle",
                                   fillColor=colors.black))
                        # z-score 값
                        d.add(GStr(cx, mb_pad - 23, f"{zv:.2f}",
                                   fontSize=z_fs, fontName=KO, textAnchor="middle",
                                   fillColor=COL_RED if abs(zv) >= 3 else COL_TXT))

                    # "z score" 레이블 (왼쪽)
                    d.add(GStr(ml - 3, mb_pad - 23, "z score",
                               fontSize=6, fontName=KO, textAnchor="end",
                               fillColor=COL_TXT))

                    # 차트 제목
                    d.add(GStr(ml + plot_w / 2, mb_pad + plot_h + 10,
                               title, fontSize=10, fontName=KO, textAnchor="middle",
                               fillColor=colors.black))

                    elems.append(d)
                    elems.append(Spacer(1, 5*mm))

        return elems

    elements += _build_zscore_section("다. 시료, 성분별 Robust Z-score", h2_style, z_all, "다")
    elements += _build_zscore_charts(z_all, group_by_method=False)
    elements += _build_zscore_section("라. 방법별 Robust Z-score",        h2_style, z_method, "라", min_n=5, split_at=5)
    elements += _build_zscore_charts(z_method, group_by_method=True, min_n=5)

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
    from utils.config import get_component_groups, get_info_fields

    GROUPS      = get_component_groups(cfg)
    INFO_FIELDS = get_info_fields(cfg)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )
    styles    = getSampleStyleSheet()
    title_sty = ParagraphStyle(
        "title", parent=styles["Title"], fontSize=18, spaceAfter=4,
        alignment=TA_CENTER, fontName=KO,
    )
    info_sty = ParagraphStyle("info", parent=styles["Normal"], fontSize=12,
                              spaceAfter=2, fontName=KO)
    sec_sty  = ParagraphStyle(
        "sec", parent=styles["Heading2"], fontSize=13, spaceAfter=3,
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

        _cell_sty = ParagraphStyle("sub_cell", fontName=KO, fontSize=8, leading=10, alignment=TA_CENTER)
        _cell_hdr = ParagraphStyle("sub_hdr",  fontName=KO, fontSize=8, leading=10, alignment=TA_CENTER,
                                   textColor=colors.white)
        def _pc(txt, hdr=False):
            return Paragraph(str(txt), _cell_hdr if hdr else _cell_sty)

        header = [_pc(c, hdr=True) for c in ["성분", "방법", "기기명", "용매"] + grp_samples]
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
                # 제출값이 하나도 없는 행은 제외
                if all(v == "" for v in vals):
                    continue
                tbl_rows.append([_pc(comp), _pc(method), _pc(equip), _pc(solvent)]
                                 + [_pc(v) for v in vals])

        if len(tbl_rows) <= 1:  # 헤더만 있으면 제출 데이터 없는 그룹 → 제외
            continue
        elements.append(Paragraph(group_name, sec_sty))

        fixed_w    = [28*mm, 28*mm, 25*mm, 20*mm]
        remaining  = 180*mm - sum(fixed_w)
        sample_w   = [remaining / len(grp_samples)] * len(grp_samples) if grp_samples else []
        t = Table(tbl_rows, colWidths=fixed_w + sample_w, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
            ("FONTNAME",      (0, 0), (-1, -1), KO),
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
            ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
            ("TOPPADDING",    (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 4*mm))

    doc.build(elements)
    return buf.getvalue()
