"""실패 로그·디버그 스냅샷을 구글 드라이브 공유 폴더로 자동 업로드.

목적: 사용자에게 로그 파일을 수동으로 요청하지 않고, 관리자가 한곳에서
모든 사용자의 실패 로그를 열람한다.

설계:
  - 기존 Google OAuth(Sheets/Gmail) 인증을 재사용 (token.json + drive.file 스코프).
  - drive.file 스코프 → 앱이 만든 폴더/파일에만 접근 (전체 드라이브 권한 아님).
  - 머신별 폴더 `B2C_RPA_logs / <머신키>` 를 최초 1회 생성하고, 그 폴더를
    관리자 이메일에 reader 권한으로 공유 → 관리자가 모든 사용자 폴더 열람.
  - 모든 드라이브 오류는 삼켜서(WARNING) 메인 RPA·실패 메일을 막지 않는다.
    드라이브는 보조 채널일 뿐이다.

토글: config.logging.upload_drive (LOG_UPLOAD_DRIVE=1) 일 때만 동작.
"""

from __future__ import annotations

import logging
import os
import socket
from pathlib import Path

log = logging.getLogger(__name__)

_ROOT_FOLDER_NAME = "B2C_RPA_logs"


def _machine_key(config) -> str:
    """폴더 이름에 쓸 머신/사용자 식별자 — SPC 아이디 우선, 없으면 호스트명."""
    user = (getattr(config.hub, "user_id", "") or "").strip()
    if user:
        return user
    try:
        return socket.gethostname()
    except Exception:
        return os.environ.get("COMPUTERNAME", "unknown")


def _admin_email(config) -> str:
    if config.logging.admin_email:
        return config.logging.admin_email
    recipients = config.notify.report_recipients
    return recipients[0] if recipients else ""


def _get_drive_service(config):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    token_path = config.notify.gmail_token_path
    if not token_path.exists():
        raise RuntimeError("token.json 없음 — 드라이브 업로드 건너뜀")

    # 스코프 미강제 로드 — 토큰에 drive.file 가 없으면 이후 Drive API 호출이
    # 403 으로 실패하고 upload_logs 가 이를 삼킨다(메일·시트 경로엔 영향 없음).
    creds = Credentials.from_authorized_user_file(str(token_path))
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            token_path.write_text(creds.to_json(), encoding="utf-8")
        else:
            raise RuntimeError("드라이브 자격증명 무효 — 재인증 필요(setup 재실행)")

    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _find_or_create_folder(service, name: str, parent_id: str | None = None) -> str:
    """이름(+부모)으로 폴더를 찾고 없으면 생성. drive.file 범위 내 앱 소유 파일만 검색됨."""
    query = (
        "mimeType='application/vnd.google-apps.folder' and trashed=false "
        f"and name='{name}'"
    )
    if parent_id:
        query += f" and '{parent_id}' in parents"
    resp = service.files().list(
        q=query, spaces="drive", fields="files(id,name)", pageSize=1
    ).execute()
    files = resp.get("files", [])
    if files:
        return files[0]["id"]

    metadata = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    if parent_id:
        metadata["parents"] = [parent_id]
    folder = service.files().create(body=metadata, fields="id").execute()
    return folder["id"]


def _share_with_admin(service, folder_id: str, admin_email: str) -> None:
    if not admin_email:
        return
    try:
        service.permissions().create(
            fileId=folder_id,
            body={"type": "user", "role": "reader", "emailAddress": admin_email},
            sendNotificationEmail=False,
            fields="id",
        ).execute()
    except Exception as exc:
        # 이미 공유됨 등은 무시
        log.debug(f"  드라이브 폴더 공유 생략/실패: {exc}")


def upload_logs(config, label: str, files: list[Path]) -> None:
    """실패 진단 파일을 머신별 드라이브 폴더에 업로드 (best-effort).

    label: 업로드 묶음 식별용 접두어 (예: "daily_2026-06-22").
    오류는 모두 삼켜서 호출자 흐름을 막지 않는다.
    """
    if not config.logging.upload_drive:
        return
    files = [p for p in files if p and p.exists()]
    if not files:
        return

    try:
        from googleapiclient.http import MediaFileUpload

        service = _get_drive_service(config)
        root_id = _find_or_create_folder(service, _ROOT_FOLDER_NAME)
        machine_id = _find_or_create_folder(service, _machine_key(config), root_id)
        _share_with_admin(service, machine_id, _admin_email(config))

        for path in files:
            media = MediaFileUpload(str(path), resumable=False)
            service.files().create(
                body={"name": f"{label}__{path.name}", "parents": [machine_id]},
                media_body=media,
                fields="id",
            ).execute()

        log.info(f"  [DRIVE] 진단 파일 {len(files)}개 업로드 완료 → {_ROOT_FOLDER_NAME}/{_machine_key(config)}")
    except Exception as exc:
        log.warning(f"  [DRIVE] 로그 업로드 실패(무시): {exc}")
