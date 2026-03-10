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
from utils.zscore import zscore_flag, zscore_color
from utils.config import get_component_from_col

# 한글 TTF 폰트 등록 (Streamlit Cloud: packages.txt에 fonts-nanum 필요)
try:
    pdfmetrics.registerFont(TTFont(
        "NanumGothic",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    ))
    KO = "NanumGothic"
except Exception:
    KO = "Helvetica"  # 한글 깨질 수 있음 (로컬 환경 폴백)


# ── 공통 헬퍼 ─────────────────────────────────────────────────

def _color_from_hex(hex_str: str):
    hex_str = hex_str.lstrip("#")
    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    return colors.Color(r / 255, g / 255, b / 255)


def _fmt_val(v) -> str:
    """수치는 소수점 2자리, 빈 값은 '-'."""
    if v is None or v == "":
        return "-"
    try:
        return f"{float(v):.2f}"
    except (ValueError, TypeError):
        return str(v)


def _cv_pct(median, mad) -> str:
    """Robust CV(%) = 1.4826 × MAD / |중앙값| × 100."""
    try:
        m, d = float(median), float(mad)
        if abs(m) < 1e-10:
            return "-"
        return f"{(1.4826 * d / abs(m)) * 100:.2f}"
    except Exception:
        return "-"


def _table_style(z_col_idx: int, z_highlights: list) -> TableStyle:
    """공통 TableStyle + Z-score 셀 색상 강조."""
    ts = TableStyle([
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
    ])
    for row_i, z_f in z_highlights:
        try:
            if not np.isnan(float(z_f)):
                ts.add("BACKGROUND", (z_col_idx, row_i), (z_col_idx, row_i),
                       _color_from_hex(zscore_color(float(z_f))))
        except Exception:
            pass
    return ts


def _pdf_header(elements, styles, title_text: str, subtitle_text: str,
                institution: str, email: str, generated_at: str):
    title_sty = ParagraphStyle(
        "rh_title", parent=styles["Title"], fontSize=15,
        spaceAfter=2, alignment=TA_CENTER, fontName=KO,
    )
    sub_sty = ParagraphStyle(
        "rh_sub", parent=styles["Normal"], fontSize=11, alignment=TA_CENTER,
        textColor=colors.HexColor("#2c3e50"), spaceAfter=4, fontName=KO,
    )
    info_sty = ParagraphStyle(
        "rh_info", parent=styles["Normal"], fontSize=10, spaceAfter=2, fontName=KO,
    )
    elements.append(Paragraph(title_text, title_sty))
    elements.append(Paragraph(subtitle_text, sub_sty))
    elements.append(Spacer(1, 5*mm))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#2c3e50")))
    elements.append(Spacer(1, 4*mm))
    elements.append(Paragraph(f"<b>기관명:</b> {institution}", info_sty))
    elements.append(Paragraph(f"<b>이메일:</b> {email}", info_sty))
    elements.append(Paragraph(f"<b>보고서 생성일:</b> {generated_at}", info_sty))
    elements.append(Spacer(1, 5*mm))
    return info_sty


def _pdf_footer(elements, styles, extra_notes: list = None):
    info_sty = ParagraphStyle(
        "rh_fi", parent=styles["Normal"], fontSize=10, spaceAfter=2, fontName=KO,
    )
    note_sty = ParagraphStyle(
        "rh_fn", parent=styles["Normal"], fontSize=8,
        textColor=colors.grey, leftIndent=4, fontName=KO,
    )
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey))
    elements.append(Spacer(1, 3*mm))
    elements.append(Paragraph("<b>판정 기준 (Robust Z-score)</b>", info_sty))
    elements.append(Paragraph("적합: |Z| ≤ 2.0", note_sty))
    elements.append(Paragraph("경고: 2.0 < |Z| ≤ 3.0", note_sty))
    elements.append(Paragraph("부적합: |Z| > 3.0", note_sty))
    elements.append(Spacer(1, 2*mm))
    elements.append(Paragraph(
        "* CV(%) = Robust 변이계수 = (1.4826 × MAD / |중앙값|) × 100  "
        "[MAD: Median Absolute Deviation, 참여기관 간 분산도 지표]", note_sty,
    ))
    elements.append(Paragraph(
        "* Robust Z-score = (제출값 − 중앙값) / (1.4826 × MAD)", note_sty,
    ))
    for note in (extra_notes or []):
        elements.append(Paragraph(f"* {note}", note_sty))


# ── 제출 확인 PDF ──────────────────────────────────────────────

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
        for _gname, items in NIR_GROUPS.items():
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
                    if rest == f"{s}":
                        pass
                    elif rest.endswith(f"_{s}"):
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


# ── Z-score 보고서 (전체) ──────────────────────────────────────

def generate_pdf_overall(
    email: str,
    institution: str,
    row_data: dict,
    zscore_row: dict,
    group_stats: dict,
    value_cols: list,
    generated_at: str = None,
    samples: list = None,
    cfg=None,
) -> bytes:
    """사료별 전체 Robust Z-score 보고서."""
    if samples is None:
        from utils.config import get_samples
        samples = get_samples(cfg)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )
    styles  = getSampleStyleSheet()
    sec_sty = ParagraphStyle(
        "ov_sec", parent=styles["Heading2"], fontSize=12,
        spaceAfter=3, spaceBefore=8, fontName=KO,
    )
    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M")
    elements: list = []

    _pdf_header(elements, styles,
                "사료 숙련도 시험 결과 보고서", "전체 Robust Z-score",
                institution, email, generated_at)

    # 성분 | 제출값 | 중앙값 | CV(%) | n | Z전체 | 판정
    cw = [50*mm, 25*mm, 25*mm, 20*mm, 12*mm, 25*mm, 23*mm]
    header = ["성분", "제출값", "중앙값", "CV(%)", "n", "Z전체", "판정"]

    for sample in samples:
        sample_cols = [
            c for c in value_cols
            if c.endswith(f"_{sample}") and not c.startswith("NIR_")
        ]
        if not sample_cols:
            continue

        elements.append(Paragraph(f"【{sample}】", sec_sty))
        tbl_rows   = [header]
        highlights = []

        for i, col in enumerate(sample_cols, start=1):
            comp  = get_component_from_col(col, samples) or col
            val   = row_data.get(col, "")
            stats = group_stats.get(col, {})
            med   = stats.get("median", "")
            mad   = stats.get("mad",    "")
            n     = stats.get("n",      "")
            z     = zscore_row.get(col, np.nan)

            try:
                z_f   = float(z)
                z_str = f"{z_f:.2f}" if not np.isnan(z_f) else "-"
            except Exception:
                z_f, z_str = np.nan, "-"

            tbl_rows.append([
                comp,
                _fmt_val(val),
                _fmt_val(med),
                _cv_pct(med, mad),
                str(n) if n != "" else "-",
                z_str,
                zscore_flag(z_f) if z_str != "-" else "N/A",
            ])
            if z_str != "-":
                highlights.append((i, z_f))

        t = Table(tbl_rows, colWidths=cw, repeatRows=1)
        t.setStyle(_table_style(z_col_idx=5, z_highlights=highlights))
        elements.append(t)
        elements.append(Spacer(1, 4*mm))

    _pdf_footer(elements, styles)
    doc.build(elements)
    return buf.getvalue()


# ── Z-score 보고서 (방법별) ───────────────────────────────────

def generate_pdf_by_method(
    email: str,
    institution: str,
    full_row: dict,
    z_method_row: dict,
    group_stats: dict,
    value_cols: list,
    generated_at: str = None,
    samples: list = None,
    cfg=None,
) -> bytes:
    """사료별 방법별 Robust Z-score 보고서."""
    if samples is None:
        from utils.config import get_samples
        samples = get_samples(cfg)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )
    styles  = getSampleStyleSheet()
    sec_sty = ParagraphStyle(
        "bm_sec", parent=styles["Heading2"], fontSize=12,
        spaceAfter=3, spaceBefore=8, fontName=KO,
    )
    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M")
    elements: list = []

    _pdf_header(elements, styles,
                "사료 숙련도 시험 결과 보고서", "방법별 Robust Z-score",
                institution, email, generated_at)

    # 성분 | 방법 | 제출값 | 중앙값 | CV(%) | n | Z방법별 | 판정
    cw = [35*mm, 40*mm, 20*mm, 20*mm, 18*mm, 10*mm, 22*mm, 15*mm]
    header = ["성분", "방법", "제출값", "중앙값", "CV(%)", "n", "Z방법별", "판정"]

    for sample in samples:
        sample_cols = [
            c for c in value_cols
            if c.endswith(f"_{sample}") and not c.startswith("NIR_")
        ]
        if not sample_cols:
            continue

        elements.append(Paragraph(f"【{sample}】", sec_sty))
        tbl_rows   = [header]
        highlights = []

        for i, col in enumerate(sample_cols, start=1):
            comp   = get_component_from_col(col, samples) or col
            method = str(full_row.get(f"{comp}_방법", "") or "-")
            val    = full_row.get(col, "")
            stats  = group_stats.get(col, {})
            med    = stats.get("median", "")
            mad    = stats.get("mad",    "")
            n      = stats.get("n",      "")
            zm     = z_method_row.get(col, np.nan)

            try:
                zm_f   = float(zm)
                zm_str = f"{zm_f:.2f}" if not np.isnan(zm_f) else "N/A"
            except Exception:
                zm_f, zm_str = np.nan, "N/A"

            tbl_rows.append([
                comp,
                method,
                _fmt_val(val),
                _fmt_val(med),
                _cv_pct(med, mad),
                str(n) if n != "" else "-",
                zm_str,
                zscore_flag(zm_f) if zm_str not in ("N/A", "-") else "N/A",
            ])
            if zm_str not in ("N/A", "-"):
                try:
                    if not np.isnan(zm_f):
                        highlights.append((i, zm_f))
                except Exception:
                    pass

        t = Table(tbl_rows, colWidths=cw, repeatRows=1)
        t.setStyle(_table_style(z_col_idx=6, z_highlights=highlights))
        elements.append(t)
        elements.append(Spacer(1, 4*mm))

    _pdf_footer(elements, styles,
                extra_notes=["Z방법별: 동일 방법 사용 기관 5개 미만이면 N/A"])
    doc.build(elements)
    return buf.getvalue()
