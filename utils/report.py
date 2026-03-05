import io
from datetime import datetime
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from utils.zscore import zscore_flag, zscore_color


def _color_from_hex(hex_str: str):
    hex_str = hex_str.lstrip("#")
    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    return colors.Color(r / 255, g / 255, b / 255)


def generate_pdf(
    email: str,
    institution: str,
    row_data: dict,
    zscore_row: dict,
    group_stats: dict,
    analyte_cols: list,
    generated_at: str = None,
) -> bytes:
    """
    개별 기관 보고서 PDF 생성
    group_stats: {analyte: {"median": float, "mad": float, "n": int}}
    zscore_row:  {analyte: float (z-score)}
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm,
        topMargin=20*mm, bottomMargin=20*mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "title", parent=styles["Title"],
        fontSize=18, spaceAfter=4, alignment=TA_CENTER,
    )
    sub_style = ParagraphStyle(
        "sub", parent=styles["Normal"],
        fontSize=10, alignment=TA_CENTER, textColor=colors.grey,
    )
    info_style = ParagraphStyle(
        "info", parent=styles["Normal"], fontSize=10, spaceAfter=2,
    )

    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M")

    elements = []

    # ── 헤더 ──────────────────────────────────────────────────
    elements.append(Paragraph("아미노산 숙련도 시험 결과 보고서", title_style))
    elements.append(Paragraph("Amino Acid Proficiency Testing Report", sub_style))
    elements.append(Spacer(1, 6*mm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2c3e50")))
    elements.append(Spacer(1, 4*mm))

    # ── 기관 정보 ─────────────────────────────────────────────
    elements.append(Paragraph(f"<b>기관명:</b> {institution}", info_style))
    elements.append(Paragraph(f"<b>이메일:</b> {email}", info_style))
    elements.append(Paragraph(f"<b>보고서 생성일:</b> {generated_at}", info_style))
    elements.append(Spacer(1, 6*mm))

    # ── 결과 테이블 ───────────────────────────────────────────
    header = ["분석항목", "제출값", "중앙값(그룹)", "MAD", "참여기관 수", "Z-score", "판정"]
    table_data = [header]

    for analyte in analyte_cols:
        submitted = row_data.get(analyte, "")
        z = zscore_row.get(analyte, np.nan)
        stats = group_stats.get(analyte, {})
        median_val = stats.get("median", "")
        mad_val = stats.get("mad", "")
        n_val = stats.get("n", "")

        try:
            z_str = f"{z:.2f}" if not np.isnan(z) else "-"
        except Exception:
            z_str = "-"

        flag = zscore_flag(z)
        table_data.append([
            analyte,
            f"{float(submitted):.4f}" if submitted != "" else "-",
            f"{float(median_val):.4f}" if median_val != "" else "-",
            f"{float(mad_val):.4f}" if mad_val != "" else "-",
            str(n_val),
            z_str,
            flag,
        ])

    col_widths = [25*mm, 25*mm, 28*mm, 22*mm, 22*mm, 22*mm, 22*mm]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)

    # 기본 스타일
    ts = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2c3e50")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8f9fa")]),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ])

    # Z-score에 따른 행 색상
    for row_idx, analyte in enumerate(analyte_cols, start=1):
        z = zscore_row.get(analyte, np.nan)
        try:
            if not np.isnan(z):
                bg = _color_from_hex(zscore_color(z))
                ts.add("BACKGROUND", (6, row_idx), (6, row_idx), bg)
        except Exception:
            pass

    t.setStyle(ts)
    elements.append(t)
    elements.append(Spacer(1, 8*mm))

    # ── 판정 기준 안내 ────────────────────────────────────────
    note_style = ParagraphStyle(
        "note", parent=styles["Normal"], fontSize=8,
        textColor=colors.grey, leftIndent=4,
    )
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
    elements.append(Spacer(1, 3*mm))
    elements.append(Paragraph("<b>판정 기준 (Robust Z-score)</b>", info_style))
    elements.append(Paragraph("✅ 적합: |Z| ≤ 2.0", note_style))
    elements.append(Paragraph("⚠️ 경고: 2.0 < |Z| ≤ 3.0", note_style))
    elements.append(Paragraph("❌ 부적합: |Z| > 3.0", note_style))
    elements.append(Spacer(1, 2*mm))
    elements.append(Paragraph(
        "* Robust Z-score = (제출값 − 중앙값) / (1.4826 × MAD)",
        note_style
    ))

    doc.build(elements)
    return buf.getvalue()
