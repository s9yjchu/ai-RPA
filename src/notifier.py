"""Gmail API 기반 성공/실패 리포트 메일 발송."""

from __future__ import annotations

import base64
import logging
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

from email.mime.base import MIMEBase
from email import encoders

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from .config import GOOGLE_OAUTH_SCOPES

log = logging.getLogger(__name__)

SCOPES = GOOGLE_OAUTH_SCOPES

# 첨부 총합 상한 (Gmail 25MB 제한 대비 여유). 초과 시 PNG 부터 제외.
_MAX_ATTACH_BYTES = 20 * 1024 * 1024


def _get_gmail_service(credentials_path: Path, token_path: Path):
    creds: Credentials | None = None

    if token_path.exists():
        # 스코프 미강제 로드 — 부여된 스코프로 refresh (drive.file 강제 시 invalid_scope).
        # 전체 스코프(SCOPES)는 신규 동의 시에만 사용.
        creds = Credentials.from_authorized_user_file(str(token_path))

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


def _select_attachments(paths: list[Path]) -> list[Path]:
    """존재하는 파일만, 총합 상한 내에서 선택. 초과 시 큰 파일(주로 PNG)부터 제외."""
    existing = [p for p in paths if p and p.exists()]
    # 작은 파일(로그·HTML)을 우선 보존하기 위해 크기 오름차순으로 채운다.
    existing.sort(key=lambda p: p.stat().st_size)
    selected: list[Path] = []
    total = 0
    for p in existing:
        size = p.stat().st_size
        if total + size > _MAX_ATTACH_BYTES:
            continue
        selected.append(p)
        total += size
    return selected


def _build_message(
    sender: str,
    to: list[str],
    subject: str,
    body: str,
    attachments: list[Path] | None = None,
) -> dict:
    files = _select_attachments(attachments or [])
    # 첨부가 있으면 multipart/mixed, 없으면 기존 동작 유지.
    msg = MIMEMultipart("mixed") if files else MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(to)
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for path in files:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(path.read_bytes())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=path.name)
        msg.attach(part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}


def send_email(
    credentials_path: Path,
    token_path: Path,
    sender: str,
    recipients: list[str],
    subject: str,
    body: str,
    attachments: list[Path] | None = None,
) -> None:
    if not recipients:
        log.info("  수신자 목록 없음 — 메일 발송 생략")
        return

    try:
        service = _get_gmail_service(credentials_path, token_path)
        message = _build_message(sender, recipients, subject, body, attachments)
        service.users().messages().send(userId="me", body=message).execute()
        if attachments:
            log.info(f"  메일 발송 완료 → {recipients} (첨부 {len(_select_attachments(attachments))}개)")
        else:
            log.info(f"  메일 발송 완료 → {recipients}")
    except Exception as exc:
        log.error(f"  메일 발송 실패: {exc}")


# ── 실패 진단 첨부 수집 ───────────────────────────────────────────────

def collect_log_artifacts(logs_dir: Path, max_snapshots: int = 6) -> list[Path]:
    """실패 메일/드라이브 업로드에 첨부할 진단 파일 목록.

    - 현재 런 로그(가장 최근 logs/*.log) 1개
    - logs/debug/ 의 최신 스냅샷(PNG/HTML) max_snapshots 개
      (로그인 실패·sis 실패 스냅샷이 최신이므로 자연히 우선 포함됨)
    """
    artifacts: list[Path] = []
    try:
        logs = sorted(
            logs_dir.glob("*.log"), key=lambda p: p.stat().st_mtime, reverse=True
        )
        if logs:
            artifacts.append(logs[0])
    except Exception as exc:
        log.debug(f"  로그 파일 수집 실패: {exc}")

    try:
        debug_dir = logs_dir / "debug"
        if debug_dir.exists():
            snaps = sorted(
                (p for p in debug_dir.iterdir() if p.suffix in (".png", ".html")),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            artifacts.extend(snaps[:max_snapshots])
    except Exception as exc:
        log.debug(f"  디버그 스냅샷 수집 실패: {exc}")

    return artifacts


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
    attachments: list[Path] | None = None,
) -> None:
    subject = f"[고객지표 RPA] {target_date} 업데이트 실패 — 확인 필요"
    body = (
        f"[{target_date}] 고객지표 자동 업데이트가 실패했습니다.\n\n"
        f"시도 횟수: {attempts}회\n"
        f"실패 사유: {reason}\n\n"
        "수동으로 업데이트하거나 RPA 로그를 확인하세요.\n"
        "(첨부된 로그·화면 캡처로 원인을 진단할 수 있습니다.)"
    )
    send_email(
        credentials_path, token_path, sender, recipients, subject, body, attachments
    )


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
