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
from utils.zscore import zscore_flag, zscore_color
from utils.config import get_component_from_col, get_sample_from_col


def _color_from_hex(hex_str: str):
    hex_str = hex_str.lstrip("#")
    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    return colors.Color(r / 255, g / 255, b / 255)


def _build_section(
    title: str,
    cols: list,
    row_data: dict,
    zscore_row: dict,
    z_method_row: dict,
    group_stats: dict,
    styles,
    samples: list,
) -> list:
    elements = []
    section_style = ParagraphStyle(
        "sec", parent=styles["Heading2"], fontSize=11, spaceAfter=3, spaceBefore=6,
    )
    elements.append(Paragraph(title, section_style))

    header = ["성분", "사료", "제출값", "중앙값", "MAD", "n", "Z전체", "Z방법별", "판정"]
    rows = [header]

    for col in cols:
        comp   = get_component_from_col(col, samples) or col
        sample = get_sample_from_col(col, samples) or "-"
        val    = row_data.get(col, "")
        z      = zscore_row.get(col, np.nan)
        zm     = z_method_row.get(col, np.nan)
        stats  = group_stats.get(col, {})

        def fmt(v, d=4):
            try:
                return f"{float(v):.{d}f}"
            except Exception:
                return "-"

        try:
            z_f  = float(z)
            z_str = f"{z_f:.2f}" if not np.isnan(z_f) else "-"
        except Exception:
            z_f, z_str = np.nan, "-"

        try:
            zm_f  = float(zm)
            zm_str = f"{zm_f:.2f}" if not np.isnan(zm_f) else "N/A"
        except Exception:
            zm_str = "N/A"

        rows.append([
            comp, sample,
            fmt(val), fmt(stats.get("median", "")), fmt(stats.get("mad", "")),
            str(stats.get("n", "")),
            z_str, zm_str,
            zscore_flag(z_f) if z_str != "-" else "N/A",
        ])

    cw = [20*mm, 20*mm, 20*mm, 20*mm, 16*mm, 10*mm, 18*mm, 18*mm, 18*mm]
    t = Table(rows, colWidths=cw, repeatRows=1)
    ts = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",   (0, 0), (-1, -1), 8),
        ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
        ("GRID",       (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ])
    for i, col in enumerate(cols, start=1):
        try:
            z_f = float(zscore_row.get(col, np.nan))
            if not np.isnan(z_f):
                ts.add("BACKGROUND", (6, i), (6, i), _color_from_hex(zscore_color(z_f)))
        except Exception:
            pass
    t.setStyle(ts)
    elements.append(t)
    elements.append(Spacer(1, 4*mm))
    return elements


def generate_pdf(
    email: str,
    institution: str,
    row_data: dict,
    zscore_row: dict,
    z_method_row: dict,
    group_stats: dict,
    value_cols: list,
    generated_at: str = None,
    samples: list = None,
) -> bytes:
    if samples is None:
        from utils.config import get_samples
        samples = get_samples()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title", parent=styles["Title"], fontSize=16, spaceAfter=4, alignment=TA_CENTER,
    )
    sub_style = ParagraphStyle(
        "sub", parent=styles["Normal"], fontSize=10, alignment=TA_CENTER, textColor=colors.grey,
    )
    info_style  = ParagraphStyle("info",  parent=styles["Normal"], fontSize=10, spaceAfter=2)
    note_style  = ParagraphStyle("note",  parent=styles["Normal"], fontSize=8,
                                 textColor=colors.grey, leftIndent=4)

    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M")
    elements = []

    elements.append(Paragraph("사료 숙련도 시험 결과 보고서", title_style))
    elements.append(Paragraph("Feed Proficiency Testing Report", sub_style))
    elements.append(Spacer(1, 5*mm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2c3e50")))
    elements.append(Spacer(1, 4*mm))
    elements.append(Paragraph(f"<b>기관명:</b> {institution}", info_style))
    elements.append(Paragraph(f"<b>이메일:</b> {email}", info_style))
    elements.append(Paragraph(f"<b>보고서 생성일:</b> {generated_at}", info_style))
    elements.append(Spacer(1, 5*mm))

    # 성분 그룹별 섹션
    from utils.config import get_component_groups
    groups = get_component_groups()

    for group_name, items in groups.items():
        comp_names = [item["name"] for item in items]
        grp_cols = [c for c in value_cols
                    if get_component_from_col(c, samples) in comp_names
                    and not c.startswith("NIR_")]
        if grp_cols:
            elements += _build_section(
                group_name, grp_cols,
                row_data, zscore_row, z_method_row, group_stats, styles, samples,
            )

    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
    elements.append(Spacer(1, 3*mm))
    elements.append(Paragraph("<b>판정 기준 (Robust Z-score)</b>", info_style))
    elements.append(Paragraph("✅ 적합: |Z| ≤ 2.0", note_style))
    elements.append(Paragraph("⚠️ 경고: 2.0 < |Z| ≤ 3.0", note_style))
    elements.append(Paragraph("❌ 부적합: |Z| > 3.0", note_style))
    elements.append(Spacer(1, 2*mm))
    elements.append(Paragraph(
        "* Z전체: 전체 기관 대비 / Z방법별: 동일 방법 기관 3개 미만이면 N/A", note_style,
    ))
    elements.append(Paragraph(
        "* Robust Z-score = (제출값 − 중앙값) / (1.4826 × MAD)", note_style,
    ))

    doc.build(elements)
    return buf.getvalue()
