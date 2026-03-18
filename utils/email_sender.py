import base64
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import streamlit as st


def _rfc2047(filename: str) -> str:
    """한글 파일명을 RFC 2047 base64로 인코딩 (네이버/다음 호환)."""
    b64 = base64.b64encode(filename.encode("utf-8")).decode("ascii")
    return f"=?utf-8?b?{b64}?="


def _attach_pdf(msg: MIMEMultipart, pdf_bytes: bytes, filename: str) -> None:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=_rfc2047(filename))
    msg.attach(part)


def send_report(to_email: str, institution: str,
                pdf_overall: bytes, pdf_method: bytes,
                pdf_summary: bytes = None) -> bool:
    """전체/방법별 Robust Z-score 보고서를 이메일로 발송. 요약 보고서 선택 첨부."""
    cfg = st.secrets["email"]
    sender   = cfg["sender"]
    password = cfg["password"]

    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = to_email
    msg["Subject"] = f"[회원사비교분석] {institution} 회원사 비교분석 결과 보고서"

    body = f"""\
{institution} 귀중

회원사 비교분석 시험 결과 보고서를 첨부 파일로 송부드립니다.
보고서에는 귀 기관의 제출값과 Robust Z-score 판정 결과가 포함되어 있습니다.

문의 사항이 있으시면 회신해 주시기 바랍니다.

감사합니다.
"""
    msg.attach(MIMEText(body, "plain", "utf-8"))
    _attach_pdf(msg, pdf_overall, f"회원사비교분석_{institution}_전체 Robust Z-score.pdf")
    _attach_pdf(msg, pdf_method,  f"회원사비교분석_{institution}_방법별 Robust Z-score.pdf")
    if pdf_summary:
        _attach_pdf(msg, pdf_summary, "회원사비교분석_전체요약.pdf")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, to_email, msg.as_string())

    return True


def send_confirmation(to_email: str, institution: str, row: dict, cfg) -> bool:
    """데이터 제출 즉시 발송하는 접수 확인 이메일 (제출 내역 PDF 첨부, Z-score 없음)."""
    from utils.report import generate_submission_pdf

    email_cfg = st.secrets["email"]
    sender    = email_cfg["sender"]
    password  = email_cfg["password"]

    body = f"""\
{institution} 귀중

데이터가 정상적으로 접수되었습니다.
첨부 파일에서 제출 내역을 확인하실 수 있습니다.

문의 사항이 있으시면 한국사료협회 사료기술연구소로 전화 주시기 바랍니다.
감사합니다.
"""
    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = to_email
    msg["Subject"] = f"[회원사 비교분석] {institution} 데이터 접수 확인"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    pdf_bytes = generate_submission_pdf(row, cfg, generated_at=row.get("제출일시"))
    _attach_pdf(msg, pdf_bytes, f"데이터제출확인서_{institution}.pdf")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, to_email, msg.as_string())

    return True


def send_all_reports(report_list: list) -> dict:
    """
    report_list: [{"email": str, "institution": str,
                   "pdf_overall": bytes, "pdf_method": bytes,
                   "pdf_summary": bytes (optional)}, ...]
    반환: {"success": [...], "fail": [...]}
    """
    result = {"success": [], "fail": []}
    for item in report_list:
        try:
            send_report(item["email"], item["institution"],
                        item["pdf_overall"], item["pdf_method"],
                        item.get("pdf_summary"))
            result["success"].append(item["email"])
        except Exception as e:
            result["fail"].append({"email": item["email"], "error": str(e)})
    return result
