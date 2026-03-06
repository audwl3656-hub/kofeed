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
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from utils.zscore import zscore_flag, zscore_color
from utils.sheets import (
    SAMPLES, PROXIMATE, CATTLE_ONLY, AMINO_ACIDS, NIR_COMPONENTS,
    get_component, get_sample,
)


def _color_from_hex(hex_str: str):
    hex_str = hex_str.lstrip("#")
    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    return colors.Color(r / 255, g / 255, b / 255)


def _section_table(
    title: str,
    value_cols: list,
    row_data: dict,
    zscore_row: dict,
    z_method_row: dict,
    group_stats: dict,
    styles_obj,
) -> list:
    """성분 그룹 섹션 PDF 요소 반환"""
    elements = []
    section_style = ParagraphStyle(
        "section", parent=styles_obj["Heading2"],
        fontSize=11, spaceAfter=3, spaceBefore=6,
    )
    elements.append(Paragraph(title, section_style))

    header = ["성분", "사료종류", "제출값", "중앙값", "MAD", "n", "Z전체", "Z방법별", "판정"]
    table_data = [header]

    for col in value_cols:
        comp    = get_component(col) or col
        sample  = get_sample(col) or "-"
        val     = row_data.get(col, "")
        z       = zscore_row.get(col, np.nan)
        zm      = z_method_row.get(col, np.nan)
        stats   = group_stats.get(col, {})
        med     = stats.get("median", "")
        mad     = stats.get("mad", "")
        n       = stats.get("n", "")

        try:
            z_str = f"{float(z):.2f}" if not np.isnan(float(z)) else "-"
        except Exception:
            z_str = "-"
        try:
            zm_str = f"{float(zm):.2f}" if not np.isnan(float(zm)) else "N/A"
        except Exception:
            zm_str = "N/A"

        table_data.append([
            comp,
            sample,
            f"{float(val):.4f}" if val != "" else "-",
            f"{float(med):.4f}" if med != "" else "-",
            f"{float(mad):.4f}" if mad != "" else "-",
            str(n),
            z_str,
            zm_str,
            zscore_flag(float(z)) if z_str != "-" else "N/A",
        ])

    col_widths = [20*mm, 22*mm, 20*mm, 22*mm, 18*mm, 10*mm, 18*mm, 18*mm, 18*mm]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)

    ts = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ])

    # Z전체 열(index 6) 색상
    for row_idx, col in enumerate(value_cols, start=1):
        z = zscore_row.get(col, np.nan)
        try:
            if not np.isnan(float(z)):
                bg = _color_from_hex(zscore_color(float(z)))
                ts.add("BACKGROUND", (6, row_idx), (6, row_idx), bg)
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
) -> bytes:
    """
    개별 기관 보고서 PDF 생성.
    value_cols: 일반 값 컬럼 목록 (NIR 제외)
    z_method_row: {col: within-method z-score}
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title", parent=styles["Title"],
        fontSize=16, spaceAfter=4, alignment=TA_CENTER,
    )
    sub_style = ParagraphStyle(
        "sub", parent=styles["Normal"],
        fontSize=10, alignment=TA_CENTER, textColor=colors.grey,
    )
    info_style = ParagraphStyle(
        "info", parent=styles["Normal"], fontSize=10, spaceAfter=2,
    )
    note_style = ParagraphStyle(
        "note", parent=styles["Normal"], fontSize=8,
        textColor=colors.grey, leftIndent=4,
    )

    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M")

    elements = []

    # ── 헤더 ──────────────────────────────────────────────────
    elements.append(Paragraph("사료 숙련도 시험 결과 보고서", title_style))
    elements.append(Paragraph("Feed Proficiency Testing Report", sub_style))
    elements.append(Spacer(1, 5*mm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2c3e50")))
    elements.append(Spacer(1, 4*mm))

    # ── 기관 정보 ─────────────────────────────────────────────
    elements.append(Paragraph(f"<b>기관명:</b> {institution}", info_style))
    elements.append(Paragraph(f"<b>이메일:</b> {email}", info_style))
    elements.append(Paragraph(f"<b>보고서 생성일:</b> {generated_at}", info_style))
    elements.append(Spacer(1, 5*mm))

    # ── 성분 그룹별 섹션 ──────────────────────────────────────
    def get_cols_by_group(comps, samples=SAMPLES):
        return [
            c for c in value_cols
            if get_component(c) in comps and get_sample(c) in samples
        ]

    prox_cols   = get_cols_by_group(PROXIMATE)
    cattle_cols = get_cols_by_group(CATTLE_ONLY, ["축우사료"])
    aa_cols     = get_cols_by_group(AMINO_ACIDS)

    for title, cols in [
        ("일반성분", prox_cols),
        ("ADF / NDF", cattle_cols),
        ("아미노산", aa_cols),
    ]:
        if cols:
            elements += _section_table(
                title, cols, row_data, zscore_row, z_method_row, group_stats, styles
            )

    # ── 판정 기준 ─────────────────────────────────────────────
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
    elements.append(Spacer(1, 3*mm))
    elements.append(Paragraph("<b>판정 기준 (Robust Z-score)</b>", info_style))
    elements.append(Paragraph("✅ 적합: |Z| ≤ 2.0", note_style))
    elements.append(Paragraph("⚠️ 경고: 2.0 < |Z| ≤ 3.0", note_style))
    elements.append(Paragraph("❌ 부적합: |Z| > 3.0", note_style))
    elements.append(Spacer(1, 2*mm))
    elements.append(Paragraph(
        "* Z전체: 전체 기관 대비 / Z방법별: 동일 방법 사용 기관 대비 (3개 미만이면 N/A)",
        note_style,
    ))
    elements.append(Paragraph(
        "* Robust Z-score = (제출값 − 중앙값) / (1.4826 × MAD)",
        note_style,
    ))

    doc.build(elements)
    return buf.getvalue()
