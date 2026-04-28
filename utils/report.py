import io
from datetime import datetime
import numpy as np
from reportlab.lib.pagesizes import A4, landscape as _landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, BaseDocTemplate, PageTemplate, Frame,
    NextPageTemplate, PageBreak, KeepTogether,
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
    _os.path.join(_os.path.dirname(__file__), "..", "fonts", "KoPub Batang Medium.ttf"),
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
]
_FONT_BOLD_CANDIDATES = [
    _os.path.join(_os.path.dirname(__file__), "..", "fonts", "KoPub Batang Bold.ttf"),
    "/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf",
]
KO = "Helvetica"
KO_B = "Helvetica-Bold"
for _fp in _FONT_CANDIDATES:
    try:
        pdfmetrics.registerFont(TTFont("KoPubBatangM", _fp))
        KO = "KoPubBatangM"
        break
    except Exception:
        continue
for _fp in _FONT_BOLD_CANDIDATES:
    try:
        pdfmetrics.registerFont(TTFont("KoPubBatangB", _fp))
        KO_B = "KoPubBatangB"
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
    meth_s    = ParagraphStyle("mcs",  fontName=KO, fontSize=10, alignment=TA_CENTER, leading=12)
    meth_s_sm = ParagraphStyle("mcss", fontName=KO, fontSize=7,  alignment=TA_CENTER, leading=9)

    def _mcp(txt):
        s = str(txt)
        return Paragraph(s, meth_s_sm if len(s) > 18 else meth_s)

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

    # 성분 순서 인덱스 (cols_for_sample 등장 순서 기준)
    comp_order: dict = {}
    for _ci, _col in enumerate(cols_for_sample):
        _c = get_component_from_col(_col, samples) or _col
        if _c not in comp_order:
            comp_order[_c] = _ci

    raw_rows = []  # (comp, method_str, val, stats, z_f, z_str, cv_str)
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
            cv_str = f"{float(cv_val):.2f}" if cv_val is not None and not np.isnan(float(cv_val)) else "-"
        except Exception:
            cv_str = "-"

        # suffix-aware method lookup (e.g. col="조단백질_축우사료_2" → sfx="_2")
        sample_in_col = get_sample_from_col(col, samples) or ""
        sfx = col[len(f"{comp}_{sample_in_col}"):] if sample_in_col else ""
        method = row_data.get(f"{comp}_방법{sfx}", "") or (inst_method or {}).get(f"{comp}{sfx}", "") or (inst_method or {}).get(comp, "") or "-"
        raw_rows.append((comp, method, val, stats, z_f, z_str, cv_str))

    # 성분 순서 → 방법명 순서 정렬 (overall/method 모두 적용)
    raw_rows.sort(key=lambda r: (comp_order.get(r[0], 999), str(r[1]).lower()))

    rows = [header]
    for comp, method, val, stats, z_f, z_str, cv_str in raw_rows:
        rows.append([
            comp, _mcp(method), fmt(val),
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


    if report_type == "overall":
        elements.append(Paragraph(
            "* Robust Z-score = (결과값 - 중위수) / (정규화된 사분위범위)  |  CV(%) = 표준편차 / 평균 x 100",
            note_style,
        ))
    else:
        elements.append(Paragraph(
            "* Robust Z-score = (결과값 - 중위수) / (정규화된 사분위범위)  |  CV(%) = 표준편차 / 평균 x 100",
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
    sample_comp_text: dict = None,
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

    def _fmt_comp(name):
        """'조지방(산분해)' -> '조지방<br/>(산분해)' for Paragraph"""
        if "(" in name:
            parts = name.split("(", 1)
            return f"{parts[0]}<br/>({parts[1]}"
        return name


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
        canvas.setFillColor(colors.black)
        canvas.drawCentredString(pw / 2, bm - 8*mm, f"- {doc.page} -")
        canvas.restoreState()

    def _land_page(canvas, doc):
        canvas.saveState()
        canvas.setFont(KO, 8)
        canvas.setFillColor(colors.black)
        canvas.drawCentredString(pw_land / 2, bm - 8*mm, f"- {doc.page} -")
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
    # 그룹 정보: 아미노산/NIR은 섹션명으로 표기
    _SECTION_GROUPS = {"아미노산", "NIR"}
    from utils.config import get_group_order as _get_go, get_component_groups as _get_cg
    _group_order = _get_go(cfg) if cfg is not None else []
    _enabled_section_groups = [g["name"] for g in _group_order if g["enabled"] and g["name"] in _SECTION_GROUPS]
    _comp_to_group: dict = {}
    if cfg is not None:
        _cg = _get_cg(cfg)
        for _gname, _items in _cg.items():
            for _it in _items:
                _comp_to_group[_it["name"]] = _gname
    # 사료별 성분 매핑 (섹션 그룹 제외한 개별 성분)
    sample_to_comps: dict = {}
    for comp in seen_comps:
        if _comp_to_group.get(comp) in _SECTION_GROUPS:
            continue
        for col in comp_to_cols.get(comp, []):
            s = get_sample_from_col(col, samples)
            if s and s in samples:
                sample_to_comps.setdefault(s, [])
                if comp not in sample_to_comps[s]:
                    sample_to_comps[s].append(comp)
    cell_ov = ParagraphStyle("cov", fontName=KO, fontSize=10, leading=13)
    ov_rows = [[Paragraph("<b>시료</b>", cell_ov), Paragraph("<b>분석항목</b>", cell_ov)]]
    for i, s in enumerate(samples):
        comps_s = list(sample_to_comps.get(s, []))
        for _sg in _enabled_section_groups:
            comps_s.append(_sg)
        if not comps_s:
            continue
        comp_text = (sample_comp_text or {}).get(s, "").strip() or ", ".join(comps_s)
        ov_rows.append([
            Paragraph(f"{s}(샘플{i+1})", cell_ov),
            Paragraph(comp_text, cell_ov),
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
        "1) 평균(x-bar), 표준편차(σ), 변이계수(CV) 등을 산출함.", info_style))
    elements.append(Paragraph(
        "2) 시험소 간 비교숙련도 시험용 Robust Z-score, Outlier : KS Q ISO 13528(통계적방법)에 따라",
        info_style))
    elements.append(Paragraph(
        "다음과 같이 계산하여 해석함.", info_style))
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
        ["|Z| ≤ 2", "      : 만족(Satisfactory)"],
        ["2 < |Z| < 3", ": 의심(Doubt)"],
        ["3 ≤ |Z|", "    : 불만족(Outlier)"],
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
    _inst_escaped = ", ".join(n.replace("&", "&amp;") for n in _inst_real)
    _inst_style = ParagraphStyle("inst", parent=info_style, leading=round(info_style.fontSize * 1.6))
    elements.append(Paragraph(f"{_inst_escaped} - {len(_inst_real)}개 회원사", _inst_style))

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

    cell_s      = ParagraphStyle("csp",  fontName=KO, fontSize=8,   leading=10, alignment=TA_CENTER)
    cell_s_hdr  = ParagraphStyle("csph", fontName=KO, fontSize=8,   leading=10, alignment=TA_CENTER, textColor=colors.white)
    cell_s_sm   = ParagraphStyle("csps", fontName=KO, fontSize=6.5, leading=8,  alignment=TA_CENTER)
    cell_s_bold = ParagraphStyle("cspb", fontName=KO_B, fontSize=8,   leading=10, alignment=TA_CENTER)
    cell_s_sm_b = ParagraphStyle("cspsb",fontName=KO_B, fontSize=6.5, leading=8,  alignment=TA_CENTER)
    _STAT_ROW_H = 16
    def _p(txt):
        return Paragraph(str(txt), cell_s)
    def _pm(txt):
        s = str(txt)
        return Paragraph(s, cell_s_sm if len(s) > 20 else cell_s)
    def _pmb(txt):
        """전체 행용 — 굵게 (KO_B 폰트 직접 사용)"""
        s = str(txt)
        return Paragraph(s, cell_s_sm_b if len(s) > 20 else cell_s_bold)
    def _pb(txt):
        return Paragraph(str(txt), cell_s_bold)
    def _hp(txt):
        return Paragraph(str(txt), cell_s_hdr)

    hdr0 = [_hp("성분"), _hp("분석법")]
    for s in valid_stat_samples:
        hdr0 += [_hp(s), _hp(""), _hp(""), _hp("")]
    hdr1 = [_hp(""), _hp("")]
    for _ in valid_stat_samples:
        hdr1 += [_hp("N"), _hp("평균"), _hp("표준편차"), _hp("변이계수")]
    stat_rows = [hdr0, hdr1]

    span_cmds = [
        ("SPAN", (0, 0), (0, 1)),
        ("SPAN", (1, 0), (1, 1)),
    ]
    for i, _ in enumerate(valid_stat_samples):
        c0 = 2 + i * n_col
        span_cmds.append(("SPAN", (c0, 0), (c0 + n_col - 1, 0)))

    whole_rows: set = set()
    empty_ranges: list = []   # (abs_row, col_start, col_end)
    comp_span_ranges: list = []

    def _is_dead(comp, s):
        for n in range(1, 10):
            sfx = "" if n == 1 else f"_{n}"
            if f"{comp}_{s}{sfx}" in non_nir_set:
                return False
        return True

    def _calc_vals(comp, s, sfx, mc, meth):
        col = f"{comp}_{s}{sfx}"
        if col not in df.columns:
            return None
        if mc and meth:
            mask = (df[mc].fillna("").astype(str).str.strip() == meth)
            vals = pd.to_numeric(df.loc[mask, col], errors="coerce").dropna()
        else:
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
            return None
        mean_ = vals.mean()
        std_  = vals.std(ddof=1) if len(vals) > 1 else float("nan")
        cv_   = (std_/mean_*100) if mean_ != 0 and not np.isnan(std_) else float("nan")
        return (len(vals), fmt(mean_),
                fmt(std_) if not np.isnan(std_) else "-",
                fmt(cv_, 2) if not np.isnan(cv_) else "-")

    row_cursor = len(stat_rows)

    for comp in seen_comps:
        method_entries = []
        for n in range(1, 10):
            sfx = "" if n == 1 else f"_{n}"
            mc  = f"{comp}_방법{sfx}"
            if mc not in df.columns:
                if n > 1: break
                continue
            for m in df[mc].dropna().astype(str).unique():
                if m.strip():
                    method_entries.append((mc, sfx, m.strip()))

        comp_rows_data = []

        def _build_row(label_cell, sfx, mc=None, meth=None, bold=False):
            abs_row = row_cursor + len(comp_rows_data)
            row = [_p(""), label_cell]
            _dat = _pb if bold else _p
            run_start = None
            for si, s in enumerate(valid_stat_samples):
                c0 = 2 + si * n_col
                is_last = (si == len(valid_stat_samples) - 1)
                cv = None if _is_dead(comp, s) else _calc_vals(comp, s, sfx, mc, meth)
                if cv is None:
                    row += ["", "", "", ""]
                    if run_start is None:
                        run_start = c0
                    if is_last:
                        empty_ranges.append((abs_row, run_start, c0 + 3))
                else:
                    if run_start is not None:
                        empty_ranges.append((abs_row, run_start, c0 - 1))
                        run_start = None
                    row += [_dat(cv[0]), _dat(cv[1]), _dat(cv[2]), _dat(cv[3])]
            return row

        for mc, sfx, meth in method_entries:
            comp_rows_data.append(_build_row(_pm(meth), sfx, mc, meth))
        comp_rows_data.append(_build_row(_pmb(f"{comp} 전체"), "", bold=True))

        if comp_rows_data:
            comp_rows_data[0][0] = _p(_fmt_comp(comp))
        base_idx = len(stat_rows)
        end_idx  = base_idx + len(comp_rows_data) - 1
        if len(comp_rows_data) > 1:
            span_cmds.append(("SPAN", (0, base_idx), (0, end_idx)))
        comp_span_ranges.append((base_idx, end_idx))
        whole_rows.add(end_idx)
        row_cursor += len(comp_rows_data)
        stat_rows.extend(comp_rows_data)

    _ALT_A = colors.white
    _ALT_B = colors.HexColor("#f0f4fa")
    _COL1  = colors.HexColor("#f8fafc")
    _DEAD_BG = colors.HexColor("#e8e8e8")
    _DEAD_LINE = colors.HexColor("#aaaaaa")

    from reportlab.graphics.shapes import Line as _Line

    # ── 연속 행 그룹화 → 같은 성분(comp) 블록 안에서만 세로 병합 ──
    # comp_span_ranges: [(base_idx, end_idx), ...] — 성분별 행 범위
    def _comp_of_row(row_idx):
        for ci, (b, e) in enumerate(comp_span_ranges):
            if b <= row_idx <= e:
                return ci
        return -1

    sorted_empty = sorted(empty_ranges, key=lambda x: (x[1], x[2], x[0]))
    final_spans = []   # (r_start, r_end, c_start, c_end)
    i = 0
    while i < len(sorted_empty):
        r, cs, ce = sorted_empty[i]
        group_rows = [r]
        j = i + 1
        while j < len(sorted_empty):
            rj, csj, cej = sorted_empty[j]
            if (csj == cs and cej == ce
                    and rj == group_rows[-1] + 1
                    and _comp_of_row(rj) == _comp_of_row(group_rows[-1])):
                group_rows.append(rj)
                j += 1
            else:
                break
        final_spans.append((group_rows[0], group_rows[-1], cs, ce))
        i = j

    # ── 대각선 Drawing 삽입 ──
    for (r_start, r_end, c_start, c_end) in final_spans:
        merged_w = sum(cw_stat[c_start : c_end + 1])
        merged_h = _STAT_ROW_H * (r_end - r_start + 1)
        d = Drawing(merged_w, merged_h)
        d.add(_Line(0, merged_h, merged_w, 0,
                    strokeColor=_DEAD_LINE, strokeWidth=0.7))
        stat_rows[r_start][c_start] = d
        for r in range(r_start, r_end + 1):
            for c in range(c_start, c_end + 1):
                if r != r_start or c != c_start:
                    stat_rows[r][c] = ""

    stat_tbl = Table(stat_rows, colWidths=cw_stat, repeatRows=2,
                     rowHeights=[_STAT_ROW_H] * len(stat_rows))

    # ── SPAN 명령 맨 앞 (ReportLab 필수) ──
    all_spans = list(span_cmds)
    for (r_start, r_end, c_start, c_end) in final_spans:
        all_spans.append(("SPAN", (c_start, r_start), (c_end, r_end)))

    tbl_style_cmds = all_spans + [
        ("BACKGROUND",     (0, 0), (-1, 1),  colors.HexColor("#4472C4")),
        ("TEXTCOLOR",      (0, 0), (-1, 1),  colors.white),
        ("FONTNAME",       (0, 0), (-1, -1), KO),
        ("FONTSIZE",       (0, 0), (-1, -1), 8),
        ("ALIGN",          (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",         (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",           (0, 0), (-1, -1), 0.4, colors.HexColor("#000000")),
        ("TOPPADDING",     (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING",  (0, 0), (-1, -1), 2),
        ("LEFTPADDING",    (0, 0), (-1, -1), 2),
        ("RIGHTPADDING",   (0, 0), (-1, -1), 2),
        ("ROWBACKGROUNDS", (0, 2), (-1, -1), [_ALT_A, _ALT_B]),
        ("BACKGROUND",     (0, 2), (0, -1),  _ALT_A),
        ("BACKGROUND",     (1, 2), (1, -1),  _COL1),
    ]
    for (r_start, r_end, c_start, c_end) in final_spans:
        tbl_style_cmds.append(("BACKGROUND", (c_start, r_start), (c_end, r_end), _DEAD_BG))

    stat_tbl.setStyle(TableStyle(tbl_style_cmds))
    elements.append(stat_tbl)

    # 가로 페이지 종료 → 세로 페이지로 복귀
    elements.append(NextPageTemplate('Content'))
    elements.append(PageBreak())

    # ━━ 나. 분석결과 요약 (CV 가로 막대 차트) ━━
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
        ml_cv, mr_cv, mt_cv = 36, 12, 28
        _leg_h_cv  = 22                        # 범례 영역 높이
        mb_cv      = 20 + 22 + _leg_h_cv      # x축 레이블(22) + 범례 영역
        plot_w_cv  = chart_w_cv - ml_cv - mr_cv
        plot_h_cv  = 120

        all_cv_vals = [v for d in cv_data.values() for v in d.values()
                       if not (isinstance(v, float) and np.isnan(v))]
        raw_max = max(all_cv_vals) if all_cv_vals else 10.0
        import math as _math
        # tick_step: raw_max 크기에 따라 자동 결정
        # 예) max≤5→1, ≤10→2, ≤20→5, ≤50→10, ≤100→20, >100→50
        _step_candidates = [1, 2, 5, 10, 20, 50, 100]
        tick_step = next(
            (s for s in _step_candidates if _math.ceil(raw_max * 1.15 / s) <= 10),
            50
        )
        n_ticks_cv = _math.ceil(raw_max * 1.15 / tick_step)
        max_cv_val = tick_step * n_ticks_cv

        n_comps_cv = len(comps_with_data)
        n_feeds_cv = len(feeds_with_data)
        group_w    = plot_w_cv / max(n_comps_cv, 1)
        bar_w_cv   = (group_w - 4) / max(n_feeds_cv, 1) * 0.5

        total_h_cv = mt_cv + plot_h_cv + mb_cv
        d_cv = Drawing(chart_w_cv, total_h_cv)
        d_cv.transform = (1, 0, 0, 1, -30, 0)
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
            _tv_lbl = str(int(tv)) if tv == int(tv) else f"{tv:.1f}"
            d_cv.add(GStr(base_x - 3, ty - 3, _tv_lbl,
                          fontSize=7, fontName=KO, textAnchor="end",
                          fillColor=colors.HexColor("#555555")))

        # 세로 막대 + X축 레이블 (직선) — 그룹 내 중앙 정렬
        _total_bar_w = n_feeds_cv * bar_w_cv
        for ci, comp in enumerate(comps_with_data):
            _bar_start = base_x + ci * group_w + (group_w - _total_bar_w) / 2
            for fi, feed in enumerate(feeds_with_data):
                cv_val = cv_data.get(comp, {}).get(feed)
                if cv_val is None or (isinstance(cv_val, float) and np.isnan(cv_val)):
                    continue
                bx = _bar_start + fi * bar_w_cv
                bh = max((cv_val / max_cv_val) * plot_h_cv, 0)
                d_cv.add(Rect(bx, base_y, max(bar_w_cv - 1, 1), bh,
                              fillColor=colors.HexColor(feed_color[feed]),
                              strokeColor=None))
            lx = base_x + ci * group_w + group_w / 2
            if "(" in comp:
                _parts = comp.split("(", 1)
                d_cv.add(GStr(lx, base_y - 9, _parts[0],
                              fontSize=7.5, fontName=KO, textAnchor="middle",
                              fillColor=colors.black))
                d_cv.add(GStr(lx, base_y - 19, "(" + _parts[1],
                              fontSize=7.5, fontName=KO, textAnchor="middle",
                              fillColor=colors.black))
            else:
                label = comp if len(comp) <= 8 else comp[:7] + "…"
                d_cv.add(GStr(lx, base_y - 12, label,
                              fontSize=7.5, fontName=KO, textAnchor="middle",
                              fillColor=colors.black))

        # 제목
        d_cv.add(GStr(base_x + plot_w_cv / 2, base_y + plot_h_cv + 14,
                      "각 성분별 변이계수(CV%)",
                      fontSize=11, fontName=KO, textAnchor="middle",
                      fillColor=colors.black))

        # 범례 (X축 레이블 아래, 가로 배열)
        _leg_item_w = 55   # 범례 항목당 폭(색상박스+텍스트)
        _leg_total_w = len(feeds_with_data) * _leg_item_w
        leg_x0 = base_x + (plot_w_cv - _leg_total_w) / 2
        leg_y0 = base_y - 22 - _leg_h_cv + 2
        for fi, feed in enumerate(feeds_with_data):
            lx = leg_x0 + fi * _leg_item_w
            d_cv.add(Rect(lx, leg_y0, 8, 6,
                          fillColor=colors.HexColor(feed_color[feed]),
                          strokeColor=None))
            d_cv.add(GStr(lx + 10, leg_y0 + 0.5, feed, fontSize=7.5, fontName=KO,
                          textAnchor="start", fillColor=colors.black))

        elements.append(d_cv)
        elements.append(Spacer(1, 3*mm))

        # 하단 데이터 표
        # ── CV 하단 표 (z-score 표 스타일) ──
        _cv_hdr_s  = ParagraphStyle("cvh", fontName=KO, fontSize=8, alignment=TA_CENTER,
                                     textColor=colors.white, leading=10)
        _cv_cell_s = ParagraphStyle("cvc", fontName=KO, fontSize=8, alignment=TA_CENTER, leading=10)
        def _cvhp(t): return Paragraph(str(t), _cv_hdr_s)
        def _cvcp(t): return Paragraph(str(t), _cv_cell_s)

        tbl_hdr = [_cvhp("사료종류")] + [_cvhp(_fmt_comp(c)) for c in comps_with_data]
        tbl_rows_cv = [tbl_hdr]
        for feed in feeds_with_data:
            row_cv = [_cvcp(feed)]
            for comp in comps_with_data:
                cv_val = cv_data.get(comp, {}).get(feed)
                if cv_val is not None and not (isinstance(cv_val, float) and np.isnan(cv_val)):
                    row_cv.append(_cvcp(f"{cv_val:.2f}"))
                else:
                    row_cv.append(_cvcp("-"))
            tbl_rows_cv.append(row_cv)
        n_cols_cv  = len(tbl_hdr)
        feed_col_w = 32 * mm
        comp_col_w = (avail_mm * mm - feed_col_w) / max(n_cols_cv - 1, 1)
        tbl_cv = Table(tbl_rows_cv, colWidths=[feed_col_w] + [comp_col_w] * (n_cols_cv - 1))
        tbl_cv_style = TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  colors.HexColor("#4472C4")),
            ("TEXTCOLOR",     (0, 0), (-1, 0),  colors.black),
            ("BACKGROUND",    (0, 1), (0, -1),  colors.HexColor("#f8fafc")),
            ("FONTNAME",      (0, 0), (-1, -1), KO),
            ("FONTSIZE",      (0, 0), (-1, -1), 8),
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#111111")),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f4fa")]),
            ("TOPPADDING",    (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("LEFTPADDING",   (0, 0), (-1, -1), 2),
            ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
        ])
        tbl_cv.setStyle(tbl_cv_style)
        # 표 중앙 정렬: 외부 단일 셀 Table로 감싸기
        tbl_cv_wrap = Table([[tbl_cv]], colWidths=[avail_mm * mm])
        tbl_cv_wrap.setStyle(TableStyle([
            ("ALIGN",   (0,0), (-1,-1), "CENTER"),
            ("VALIGN",  (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING",  (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING",   (0,0), (-1,-1), 0),
            ("BOTTOMPADDING",(0,0), (-1,-1), 0),
        ]))
        elements.append(tbl_cv_wrap)

    elements.append(Spacer(1, 4*mm))
    elements.append(Paragraph("나. 분석결과 요약", h2_style))
    
    elements.append(Spacer(1, 4*mm))
    if summary_text:
        for line in summary_text.splitlines():
            line = line.strip()
            if line:
                elements.append(Paragraph(line, info_style))
        elements.append(Spacer(1, 4*mm))


    def _get_zv(z_src, row_idx, col):
        try:
            if isinstance(z_src, _pd.DataFrame):
                return z_src.at[row_idx, col] if col in z_src.columns else np.nan
            ser = z_src.get(col)
            return ser.at[row_idx] if ser is not None else np.nan
        except Exception:
            return np.nan

    def _charts_for_comp(comp, z_src, group_by_method=False, min_n=1):
        """단일 성분 z-score 막대그래프 (z-score 오름차순) elements 반환"""
        import math as _m, re as _re
        chart_elems = []
        chart_w = avail_mm * mm
        ml, mr  = 40, 40
        mt_pad, mb_pad = 24, 38
        plot_w  = chart_w - ml - mr
        plot_h  = 88
        COL_BLUE = colors.HexColor("#4472C4")
        COL_RED  = colors.HexColor("#dc2626")
        COL_GRID = colors.HexColor("#dddddd")
        COL_ZERO = colors.HexColor("#888888")
        COL_TXT  = colors.HexColor("#444444")

        def _sol_abbr(val):
            m = str(val).strip().upper()
            if not m or m in ("NAN", "해당없음", "-"): return ""
            if "P.ETHER" in m or "PETROLEUM" in m: return "P"
            if "(D)E.ETHER" in m or "DIETHYL" in m: return "E"
            if "헥산" in val or "HEXAN" in m: return "H"
            if "에탄올" in val or "ETHANOL" in m: return "EtOH"
            if "아세톤" in val or "ACETON" in m: return "Ac"
            if "ETHER" in m or "에테르" in val: return "E"
            return ""

        sfx_list = [""]
        for n in range(2, 10):
            if any(f"{comp}_{s}_{n}" in non_nir_set for s in samples):
                sfx_list.append(f"_{n}")
            else:
                break

        valid_samps = [s for s in samples
                       if any(f"{comp}_{s}{sfx}" in non_nir_set for sfx in sfx_list)]
        if not valid_samps:
            return chart_elems

        is_fat = "조지방" in str(comp)

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
                    if is_fat:
                        # 해당 sfx 용매 컬럼 읽기
                        def _read_sol(s_col):
                            if s_col not in df.columns:
                                return ""
                            try:
                                v = str(df.at[row_idx, s_col]).strip()
                                return "" if v in ("nan", "해당없음", "-", "") else v
                            except Exception:
                                return ""
                        sol_val = _read_sol(f"{comp}_용매{sfx}")
                        # sfx="_2" 등 추가 행에서 비어있으면 기본(sfx="") 행 용매로 fallback
                        if not sol_val and sfx:
                            sol_val = _read_sol(f"{comp}_용매")
                        abbr = _sol_abbr(sol_val) if sol_val else _sol_abbr(meth_str)
                        label = f"{inst}{abbr}" if abbr else str(inst)
                    else:
                        label = str(inst)
                    grouped.setdefault(grp_key, []).append((label, zv))

            grouped = {k: v for k, v in grouped.items() if len(v) >= min_n}
            if not grouped:
                continue

            ordered_methods = get_method_options(cfg, comp=comp) if cfg is not None else []

            def _gk2(k, _om=ordered_methods):
                try:    return _om.index(k)
                except: return len(_om)

            for grp_key, bars in sorted(grouped.items(), key=lambda x: _gk2(x[0])):
                # 그래프: z-score 오름차순 정렬
                bars.sort(key=lambda item: item[1])
                n_bars = len(bars)
                title = (f"{s}({comp}-{grp_key})"
                         if group_by_method and grp_key not in ("__ALL__", "(방법 미기재)")
                         else f"{s}({comp})")

                z_vals  = [z for _, z in bars]
                abs_max = max(abs(min(z_vals)), abs(max(z_vals)), 2.5)
                y_hi    = float(_m.ceil(abs_max * 1.15))
                y_lo    = -y_hi
                y_span  = y_hi * 2

                # y축 tick step: y_hi <= 5 → 1, 5 < y_hi <= 10 → 2, > 10 → 5
                if y_hi <= 5:
                    tick_step = 1
                elif y_hi <= 10:
                    tick_step = 2
                else:
                    tick_step = 5

                slot_w  = plot_w / n_bars
                bar_w   = max(slot_w * 0.4, 2.0)
                lbl_fs  = max(5.0, 7.5 - max(0, n_bars - 20) * 0.1)
                z_fs    = max(4.5, 6.5 - max(0, n_bars - 20) * 0.1)

                d = Drawing(chart_w, mt_pad + plot_h + mb_pad)
                d.transform = (1, 0, 0, 1, -5, 0)
                d.add(Rect(ml, mb_pad, plot_w, plot_h,
                           fillColor=colors.white,
                           strokeColor=colors.HexColor("#aaaaaa"), strokeWidth=0.5))

                # tick 범위: y_lo ~ y_hi, step 단위 정수만
                tick_lo = int(_m.ceil(y_lo / tick_step)) * tick_step
                tick_hi = int(_m.floor(y_hi / tick_step)) * tick_step
                tv = tick_lo
                while tv <= tick_hi + 1e-9:
                    ty = mb_pad + (tv - y_lo) / y_span * plot_h
                    d.add(Line(ml, ty, ml + plot_w, ty,
                               strokeColor=COL_ZERO if tv == 0 else COL_GRID,
                               strokeWidth=0.7 if tv == 0 else 0.3))
                    label_val = int(tv) if tv == int(tv) else tv
                    d.add(GStr(ml - 3, ty - 3.5, str(label_val),
                               fontSize=6.5, fontName=KO, textAnchor="end",
                               fillColor=COL_TXT))
                    tv = round(tv + tick_step, 9)

                zero_y = mb_pad + (0 - y_lo) / y_span * plot_h
                for i, (lbl, zv) in enumerate(bars):
                    cx  = ml + i * slot_w + slot_w / 2
                    bx  = cx - bar_w / 2
                    bc  = COL_RED if abs(zv) >= 3 else COL_BLUE
                    bh  = abs(zv) / y_span * plot_h
                    by  = zero_y if zv >= 0 else zero_y - bh
                    d.add(Rect(bx, by, bar_w, max(bh, 0.5), fillColor=bc, strokeColor=None))
                    d.add(GStr(cx, mb_pad - 11, str(lbl),
                               fontSize=lbl_fs, fontName=KO, textAnchor="middle",
                               fillColor=colors.black))
                    d.add(GStr(cx, mb_pad - 23, f"{zv:.2f}",
                               fontSize=z_fs, fontName=KO, textAnchor="middle",
                               fillColor=COL_RED if abs(zv) >= 3 else COL_TXT))

                d.add(GStr(ml - 3, mb_pad - 23, "z score",
                           fontSize=6, fontName=KO, textAnchor="end",
                           fillColor=COL_TXT))
                d.add(GStr(ml + plot_w / 2, mb_pad + plot_h + 10, title,
                           fontSize=10, fontName=KO, textAnchor="middle",
                           fillColor=colors.black))
                chart_elems.append(d)
                chart_elems.append(Spacer(1, 4*mm))

        return chart_elems

    # 페이지 본문 높이 (pt) — 표+그래프 함께 들어가는지 판단
    _PAGE_H = fh

    def _sol_abbr_fn(val):
        m = str(val).strip().upper()
        if not m or m in ("NAN", "해당없음", "-"): return ""
        if "P.ETHER" in m or "PETROLEUM" in m: return "P"
        if "(D)E.ETHER" in m or "DIETHYL" in m: return "E"
        if "헥산" in val or "HEXAN" in m: return "H"
        if "에탄올" in val or "ETHANOL" in m: return "EtOH"
        if "아세톤" in val or "ACETON" in m: return "Ac"
        if "ETHER" in m or "에테르" in val: return "E"
        return ""

    def _build_zscore_section(sec_title, sec_heading_style, z_src, heading_prefix, min_n=1, split_at=0):
        elems = []
        elems.append(Paragraph(sec_title, sec_heading_style))

        hdr_s  = ParagraphStyle("zh", fontName=KO, fontSize=9, alignment=TA_CENTER,
                                 textColor=colors.white, leading=11)
        cell_s2 = ParagraphStyle("zc", fontName=KO, fontSize=8, alignment=TA_CENTER, leading=10)
        red_s   = ParagraphStyle("zr", fontName=KO, fontSize=8, alignment=TA_CENTER, leading=10,
                                 textColor=colors.HexColor("#dc2626"))
        org_s   = ParagraphStyle("zo", fontName=KO, fontSize=8, alignment=TA_CENTER, leading=10,
                                 textColor=colors.HexColor("#000000"))

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

        def _lab_sort_key(row):
            """표: Lab 오름차순"""
            import re as _re2
            txt = str(row[1].text) if hasattr(row[1], "text") else str(row[1])
            m = _re2.match(r'[A-Za-z]*(\d+)', txt)
            return (0, int(m.group(1))) if m else (1, txt)

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

            # 컬럼 폭: 분석방법 | Lab | [용매] | (결과, Z-score) × n_samples
            _sol_col = f"{comp}_용매"
            def _sol_has_data(sc):
                return (sc in df.columns and
                        df[sc].dropna().astype(str).str.strip()
                        .replace({"nan":"","해당없음":"","-":""}).ne("").any())
            _has_solvent = _sol_has_data(_sol_col)
            if not _has_solvent:
                for _sfx2 in sfx_list[1:]:
                    if _sol_has_data(f"{comp}_용매{_sfx2}"):
                        _has_solvent = True; break

            meth_w = 40
            lab_w  = 14
            sol_w  = 12
            n_samp = len(valid_samples)
            _fixed_w = meth_w + lab_w + (sol_w if _has_solvent else 0)
            per_s  = (avail_mm - _fixed_w) / n_samp
            res_w  = per_s * 0.5
            z_w    = per_s * 0.5
            if _has_solvent:
                cw_z    = [meth_w*mm, lab_w*mm, sol_w*mm] + [res_w*mm, z_w*mm] * n_samp
                _dat_col = 3
            else:
                cw_z    = [meth_w*mm, lab_w*mm] + [res_w*mm, z_w*mm] * n_samp
                _dat_col = 2

            # 헤더 1행
            hdr1 = [_hp("분석방법"), _hp("Lab")]
            if _has_solvent:
                hdr1.append(_hp("용매"))
            for s in valid_samples:
                col0 = f"{comp}_{s}"
                gs   = group_stats.get(col0, {})
                try:
                    med = float(gs.get("median", float("nan")))
                    med_str = f"{med:.2f}" if not np.isnan(med) else "-"
                except Exception:
                    med_str = "-"
                hdr1 += [_hp(f"{s}  (중간값: {med_str})"), _hp("")]

            # 헤더 2행
            hdr2 = [_hp(""), _hp("")]
            if _has_solvent:
                hdr2.append(_hp(""))
            for _ in valid_samples:
                hdr2 += [_hp("결과"), _hp("Z-score")]

            z_rows = [hdr1, hdr2]
            span_z = [("SPAN", (0, 0), (0, 1)), ("SPAN", (1, 0), (1, 1))]
            if _has_solvent:
                span_z.append(("SPAN", (2, 0), (2, 1)))
            for si in range(n_samp):
                span_z.append(("SPAN", (_dat_col + si*2, 0), (_dat_col + si*2 + 1, 0)))

            # 방법별 그룹핑 + 방법별 raw 값 수집 (중간값 계산용)
            method_to_rows: dict = {}
            method_raw_vals: dict = {}   # {meth: {s: [float, ...]}}
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
                            if not np.isnan(rv):
                                has_any = True
                                method_raw_vals.setdefault(meth, {}).setdefault(s, []).append(rv)
                        except Exception:
                            res_str = "-"
                        zv = _get_zv(z_src, row_idx, col)
                        row_vals += [_cp(res_str), _zc(zv)]
                    if not has_any:
                        continue
                    lab_label = inst if sfx == "" else f"{inst}{sfx}"
                    # 용매 셀 (약어 적용)
                    if _has_solvent:
                        def _rsol(sc):
                            if sc not in df.columns: return ""
                            try:
                                v = str(df.at[row_idx, sc]).strip()
                                return "" if v in ("nan","해당없음","-","") else v
                            except: return ""
                        sol_v = _rsol(f"{comp}_용매{sfx}") or _rsol(f"{comp}_용매")
                        _abbr_v = _sol_abbr_fn(sol_v) if sol_v else ""
                        sol_cell = [_cp(_abbr_v if _abbr_v else sol_v)]
                    else:
                        sol_cell = []
                    method_to_rows.setdefault(meth, []).append(
                        [_cp(""), _cp(lab_label)] + sol_cell + row_vals
                    )

            ordered_methods = get_method_options(cfg, comp=comp) if cfg is not None else []
            def _meth_order(item):
                try: return ordered_methods.index(item[0])
                except ValueError: return len(ordered_methods)

            if min_n > 1:
                method_to_rows = {m: r for m, r in method_to_rows.items() if len(r) >= min_n}

            if not method_to_rows:
                continue

            sorted_methods = sorted(method_to_rows.items(), key=_meth_order)
            total_data_rows = sum(len(r) for _, r in sorted_methods)
            do_split = split_at > 0 and total_data_rows > split_at

            def _make_zt(rows, spans):
                _n_data = len(rows) - 2
                _row_h  = [20, 16] + [14] * _n_data
                t = Table(rows, colWidths=cw_z, repeatRows=2, rowHeights=_row_h)
                _zt_style = spans + [
                    ("BACKGROUND",    (0, 0), (-1, 1),  colors.HexColor("#4472C4")),
                    ("TEXTCOLOR",     (0, 0), (-1, 1),  colors.white),
                    ("FONTNAME",      (0, 0), (-1, -1), KO),
                    ("FONTSIZE",      (0, 2), (-1, -1), 8),
                    ("FONTSIZE",      (0, 0), (-1, 1),  9),
                    ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
                    ("GRID",          (0, 0), (-1, -1), 0.4, colors.HexColor("#000000")),
                    ("TOPPADDING",    (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                    ("LEFTPADDING",   (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING",  (0, 0), (-1, -1), 2),
                    ("ROWBACKGROUNDS",(0, 2), (-1, -1), [colors.white, colors.HexColor("#f0f4fa")]),
                    ("BACKGROUND",    (0, 2), (0, -1),  colors.white),
                    ("BACKGROUND",    (1, 2), (_dat_col - 1, -1),  colors.HexColor("#f8fafc")),
                ]
                t.setStyle(TableStyle(_zt_style))
                return t

            comp_heading = Paragraph(f"{num}) {comp}", h3_style)

            # ── 표 elements 구성 ──
            tbl_elems = []
            meth_sub_style = ParagraphStyle("ms", fontName=KO, fontSize=9,
                                            spaceBefore=4, spaceAfter=3,
                                            keepWithNext=1,
                                            textColor=colors.black)
            if do_split:
                for meth, mrows in sorted_methods:
                    mrows.sort(key=_lab_sort_key)
                    # 방법별 중간값 헤더 생성
                    meth_hdr1 = [_hp("분석방법"), _hp("Lab")]
                    if _has_solvent:
                        meth_hdr1.append(_hp("용매"))
                    for s in valid_samples:
                        m_vals = method_raw_vals.get(meth, {}).get(s, [])
                        try:
                            med_m = float(np.median(m_vals)) if m_vals else float("nan")
                            med_str_m = f"{med_m:.2f}" if not np.isnan(med_m) else "-"
                        except Exception:
                            med_str_m = "-"
                        meth_hdr1 += [_hp(f"{s}  (중간값: {med_str_m})"), _hp("")]
                    sub_rows  = [meth_hdr1, hdr2]
                    sub_spans = [("SPAN", (0, 0), (0, 1)), ("SPAN", (1, 0), (1, 1))]
                    if _has_solvent:
                        sub_spans.append(("SPAN", (2, 0), (2, 1)))
                    for si in range(n_samp):
                        sub_spans.append(("SPAN", (_dat_col + si*2, 0), (_dat_col + si*2 + 1, 0)))
                    mrows[0][0] = _cp(meth)
                    sub_rows.extend(mrows)
                    if len(mrows) > 1:
                        sub_spans.append(("SPAN", (0, 2), (0, 1 + len(mrows))))
                    tbl_elems.append(Paragraph(f"▶ {meth}", meth_sub_style))
                    tbl_elems.append(_make_zt(sub_rows, sub_spans))
                    tbl_elems.append(Spacer(1, 3*mm))
            else:
                for meth, mrows in sorted_methods:
                    mrows.sort(key=_lab_sort_key)
                    mrows[0][0] = _cp(meth)
                    mstart = len(z_rows)
                    z_rows.extend(mrows)
                    if len(mrows) > 1:
                        span_z.append(("SPAN", (0, mstart), (0, mstart + len(mrows) - 1)))
                if len(z_rows) <= 2:
                    continue
                tbl_elems.append(_make_zt(z_rows, span_z))
                tbl_elems.append(Spacer(1, 4*mm))

            # ── 그래프 elements ──
            chart_elems = _charts_for_comp(
                comp, z_src,
                group_by_method=(heading_prefix == "라"),
                min_n=min_n,
            )

            # ── 표+그래프 한 페이지 가능 여부 추정 ──
            # 실측 기반: 행 14pt(8pt폰트+패딩+leading), 차트 165pt(Drawing150+Spacer11+여유4)
            ROW_H    = 14
            CHART_H  = 165
            OVERHEAD = 50   # h3 제목 30pt + 여백

            est_tbl_h = (2 + total_data_rows) * ROW_H
            if do_split:
                # ▶ sub-heading 15pt + 헤더 2행
                est_tbl_h += len(sorted_methods) * (15 + 2 * ROW_H)
            n_charts = len([e for e in chart_elems if isinstance(e, Drawing)])
            est_chart_h = n_charts * CHART_H

            # 임계값: 실제 페이지 높이의 95% — 진짜 안 들어갈 때만 분리
            fits = (OVERHEAD + est_tbl_h + est_chart_h) <= _PAGE_H * 0.95

            if fits and chart_elems:
                # KeepTogether: 블록 전체가 한 페이지에 맞을 때
                block = [comp_heading] + tbl_elems + chart_elems
                elems.append(KeepTogether(block))
            elif chart_elems:
                # 표 페이지 / 그래프 페이지 엄격 분리
                # 제목+표는 KeepTogether로 묶어 제목이 고아가 되지 않게
                elems.append(KeepTogether([comp_heading] + tbl_elems))
                elems.append(PageBreak())
                elems.extend(chart_elems)
                elems.append(PageBreak())
            else:
                # 그래프 없음 — 표만
                elems.append(KeepTogether([comp_heading] + tbl_elems))

        return elems

    elements.append(PageBreak())
    elements += _build_zscore_section("다. 시료, 성분별 Robust Z-score", h2_style, z_all,    "다")
    elements += _build_zscore_section("라. 방법별 Robust Z-score",        h2_style, z_method, "라", min_n=5, split_at=5)

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

    elements.append(Paragraph("한국사료협회 비교분석 데이터 제출 확인서", title_sty))
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
