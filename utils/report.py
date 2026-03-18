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
from reportlab.graphics.shapes import Drawing, Rect
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
        abs_z = abs(z_f)
        if abs_z > 3:
            return Paragraph(f'<font color="red"><b><u>{z_str}</u></b></font>', z_cell_style)
        elif abs_z > 2:
            return Paragraph(f'<font color="#27ae60">{z_str}</font>', z_cell_style)
        else:
            return Paragraph(f'<font color="#2980b9">{z_str}</font>', z_cell_style)

    rows = [header]
    for col in cols_for_sample:
        comp  = get_component_from_col(col, samples) or col
        val   = row_data.get(col, "")
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

        method = (inst_method or {}).get(comp, "")
        rows.append([
            comp, method, fmt(val),
            fmt(stats.get("median", "")),
            cv_str,
            str(stats.get("n", "")),
            _z_cell(z_f, z_str),
        ])

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
        title="회원사 비교분석 시험 Robust Z-score 보고서 (전체)",
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
        title="회원사 비교분석 시험 Robust Z-score 보고서 (방법별)",
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

    def _cv_bar(cv_val, draw_w, draw_h=10, max_cv=30.0):
        try:
            cv = float(cv_val)
            if np.isnan(cv):
                return "-"
        except Exception:
            return "-"
        d = Drawing(draw_w, draw_h)
        fill_w = min(cv / max_cv, 1.0) * draw_w
        bar_color = (colors.HexColor("#27ae60") if cv < 5 else
                     colors.HexColor("#f39c12") if cv < 10 else
                     colors.HexColor("#e74c3c"))
        d.add(Rect(0, 0, draw_w, draw_h,
                   fillColor=colors.HexColor("#ecf0f1"),
                   strokeColor=colors.HexColor("#cccccc"), strokeWidth=0.5))
        if fill_w > 0:
            d.add(Rect(0, 0, fill_w, draw_h, fillColor=bar_color, strokeColor=None))
        d.add(GStr(draw_w / 2, 1.5, f"{cv:.1f}%", fontSize=7, fontName=KO,
                   fillColor=colors.HexColor("#2c3e50"), textAnchor="middle"))
        return d

    non_nir = [c for c in value_cols if not c.startswith("NIR_")]
    non_nir_set = set(non_nir)

    # 성분 목록 (순서 유지)
    seen_comps: list = []
    comp_to_cols: dict = {}
    for col in non_nir:
        comp = get_component_from_col(col, samples) or col
        if comp not in seen_comps:
            seen_comps.append(comp)
        comp_to_cols.setdefault(comp, []).append(col)

    # 가용폭 계산
    avail_mm  = 180
    comp_w    = 35
    per_s     = (avail_mm - comp_w) / max(len(samples), 1)
    mean_w    = per_s * 0.40
    cv_w      = per_s * 0.60
    bar_pts   = cv_w * mm  # Drawing width (points)

    elements = []

    # 제목
    elements.append(Paragraph("회원사 비교분석 전체 요약 보고서", title_style))
    elements.append(Spacer(1, 2*mm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2c3e50")))
    elements.append(Spacer(1, 4*mm))
    elements.append(Paragraph(f"<b>보고서 생성일:</b> {generated_at}", info_style))
    elements.append(Spacer(1, 5*mm))

    # ━━━ 1. 통계 요약 ━━━
    elements.append(Paragraph("1. 전체 통계 요약", section_style))

    cw_stats = [comp_w*mm] + [mean_w*mm, cv_w*mm] * len(samples)
    hdr0 = ["성분"] + [s for s in samples for _ in range(2)]
    hdr1 = [""]     + ["평균", "CV(%)"] * len(samples)
    stat_rows = [hdr0, hdr1]
    for comp in seen_comps:
        row = [comp]
        for s in samples:
            col = f"{comp}_{s}"
            sd = group_stats.get(col, {})
            row.append(fmt(sd.get("mean", "")))
            row.append(_cv_bar(sd.get("cv"), bar_pts))
        stat_rows.append(row)

    stat_tbl = Table(stat_rows, colWidths=cw_stats, repeatRows=2)
    span_cmds = [("SPAN", (0, 0), (0, 1))]
    for i in range(len(samples)):
        c0 = 1 + i * 2
        span_cmds.append(("SPAN", (c0, 0), (c0 + 1, 0)))
    stat_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 1),  colors.HexColor("#2c3e50")),
        ("TEXTCOLOR",     (0, 0), (-1, 1),  colors.white),
        ("FONTNAME",      (0, 0), (-1, -1), KO),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0, 2), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ] + span_cmds))
    elements.append(stat_tbl)
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
        elif abs(z_f) > 2:
            return Paragraph(f'<font color="#27ae60">{z_str}</font>', z_cell_p)
        return Paragraph(f'<font color="#2980b9">{z_str}</font>', z_cell_p)

    inst_names = df[inst_field].fillna("").astype(str).tolist()
    idx_list   = df.index.tolist()

    inst_w = 45
    samp_w = (avail_mm - inst_w) / max(len(samples), 1)
    cw_z   = [inst_w*mm] + [samp_w*mm] * len(samples)

    def _build_zscore_section(sec_title, z_src):
        elems = [Paragraph(sec_title, section_style)]
        for num, comp in enumerate(seen_comps, 1):
            elems.append(Paragraph(f"{num}) {comp}", comp_style))
            z_rows = [["기관명"] + list(samples)]
            for inst, row_idx in zip(inst_names, idx_list):
                row = [inst]
                for s in samples:
                    col = f"{comp}_{s}"
                    if col not in non_nir_set:
                        row.append("-")
                        continue
                    try:
                        if isinstance(z_src, pd.DataFrame):
                            zv = z_src.at[row_idx, col] if col in z_src.columns else np.nan
                        else:
                            ser = z_src.get(col)
                            zv = ser.at[row_idx] if ser is not None else np.nan
                    except Exception:
                        zv = np.nan
                    row.append(z_cell(zv))
                z_rows.append(row)
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
    elements.append(Paragraph("CV(%) 막대: 녹색 &lt; 5%, 주황 5~10%, 빨강 &gt; 10%", note_style))

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
            comp    = item["name"]
            method  = str(row.get(f"{comp}_방법", "") or "")
            equip   = str(row.get(f"{comp}_기기",  "") or "")
            solvent = str(row.get(f"{comp}_용매",  "") or "")
            vals    = [_fmt(row.get(f"{comp}_{s}")) for s in grp_samples]
            tbl_rows.append([comp, method, equip, solvent] + vals)

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
