"""Gmail API 기반 성공/실패 리포트 메일 발송."""

from __future__ import annotations

import base64
import logging
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

log = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets",
]


def _get_gmail_service(credentials_path: Path, token_path: Path):
    creds: Credentials | None = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path), SCOPES
            )
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds)


def _build_message(sender: str, to: list[str], subject: str, body: str) -> dict:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to)
    msg.attach(MIMEText(body, "plain", "utf-8"))
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}


def send_email(
    credentials_path: Path,
    token_path: Path,
    sender: str,
    recipients: list[str],
    subject: str,
    body: str,
) -> None:
    if not recipients:
        log.info("  수신자 목록 없음 — 메일 발송 생략")
        return

    try:
        service = _get_gmail_service(credentials_path, token_path)
        message = _build_message(sender, recipients, subject, body)
        service.users().messages().send(userId="me", body=message).execute()
        log.info(f"  메일 발송 완료 → {recipients}")
    except Exception as exc:
        log.error(f"  메일 발송 실패: {exc}")


# ── 편의 함수 ────────────────────────────────────────────────────────

def notify_success(
    credentials_path: Path,
    token_path: Path,
    sender: str,
    recipients: list[str],
    target_date: date,
    metrics_summary: dict[str, Any],
) -> None:
    subject = f"[고객지표 RPA] {target_date} 업데이트 완료"
    lines = [f"[{target_date}] 고객지표 자동 업데이트가 완료되었습니다.\n"]
    for k, v in metrics_summary.items():
        lines.append(f"  {k}: {v:,}" if isinstance(v, (int, float)) and v == v else f"  {k}: {v}")
    body = "\n".join(lines)
    send_email(credentials_path, token_path, sender, recipients, subject, body)


def notify_failure(
    credentials_path: Path,
    token_path: Path,
    sender: str,
    recipients: list[str],
    target_date: date,
    reason: str,
    attempts: int,
) -> None:
    subject = f"[고객지표 RPA] {target_date} 업데이트 실패 — 확인 필요"
    body = (
        f"[{target_date}] 고객지표 자동 업데이트가 실패했습니다.\n\n"
        f"시도 횟수: {attempts}회\n"
        f"실패 사유: {reason}\n\n"
        "수동으로 업데이트하거나 RPA 로그를 확인하세요."
    )
    send_email(credentials_path, token_path, sender, recipients, subject, body)


def notify_data_not_ready(
    credentials_path: Path,
    token_path: Path,
    sender: str,
    recipients: list[str],
    target_date: date,
    attempts: int,
) -> None:
    subject = f"[고객지표 RPA] {target_date} 데이터 미준비 — 자동 재시도 예정"
    body = (
        f"[{target_date}] OLAP 데이터가 아직 준비되지 않아 업데이트를 건너뜁니다.\n\n"
        f"시도 횟수: {attempts}회\n"
        "30분 후 자동으로 재시도합니다.\n"
        "11시 30분까지 계속 실패하면 별도 알림이 발송됩니다."
    )
    send_email(credentials_path, token_path, sender, recipients, subject, body)
