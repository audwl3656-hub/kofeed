import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.utils import encode_rfc2231
import streamlit as st


def _attach_pdf(msg: MIMEMultipart, pdf_bytes: bytes, filename: str) -> None:
    """PDF를 이메일에 첨부. 한글 파일명을 RFC 5987로 인코딩."""
    part = MIMEBase("application", "octet-stream")
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)
    # RFC 5987 인코딩으로 한글 파일명 처리
    encoded_name = encode_rfc2231(filename, charset="utf-8")
    part.add_header("Content-Disposition", "attachment",
                    **{"filename*": encoded_name})
    msg.attach(part)


def send_report(to_email: str, institution: str,
                pdf_overall: bytes, pdf_method: bytes) -> bool:
    cfg = st.secrets["email"]
    sender = cfg["sender"]
    password = cfg["password"]

    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = f"[회원사비교분석] {institution} 회원사 비교분석 결과 보고서"

    body = f"""\
{institution} 귀중

회원사 비교분석 시험 결과 보고서를 첨부 파일로 송부드립니다.
보고서에는 귀 기관의 제출값과 Robust Z-score 판정 결과가 포함되어 있습니다.

  첨부 1: 전체 Robust Z-score 보고서
  첨부 2: 방법별 Robust Z-score 보고서

문의 사항이 있으시면 회신해 주시기 바랍니다.

감사합니다.
"""
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for pdf_bytes, label in [
        (pdf_overall, "전체"),
        (pdf_method,  "방법별"),
    ]:
        _attach_pdf(msg, pdf_bytes,
                    f"회원사비교분석_{institution}_{label} Robust Z-score.pdf")

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
    report_list: [{"email": str, "institution": str, "pdf_bytes": bytes}, ...]
    반환: {"success": [...], "fail": [...]}
    """
    result = {"success": [], "fail": []}
    for item in report_list:
        try:
            send_report(item["email"], item["institution"],
                        item["pdf_overall"], item["pdf_method"])
            result["success"].append(item["email"])
        except Exception as e:
            result["fail"].append({"email": item["email"], "error": str(e)})
    return result
