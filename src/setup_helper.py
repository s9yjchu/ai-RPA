"""setup_user.bat 에서 호출하는 설치 헬퍼.

사용법:
  python -m src.setup_helper auth [google_email]
  python -m src.setup_helper validate
"""
import sys


def _reconfigure():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def do_auth(email_hint: str = "") -> None:
    _reconfigure()
    from pathlib import Path

    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    from src.sheets_writer import SCOPES

    creds_path = Path("./credentials.json")
    token_path = Path("./token.json")

    # 스코프 미강제 로드 — 기존 토큰의 부여 스코프로 다룬다.
    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path))

    # 필요한 전체 스코프(sheets/gmail/drive)가 모두 부여돼 있는지 확인.
    have = set(creds.scopes or []) if creds else set()
    has_all_scopes = set(SCOPES) <= have

    if creds and creds.valid and has_all_scopes:
        pass  # 이미 전체 스코프로 유효 — 재동의 불필요
    elif creds and creds.expired and creds.refresh_token and has_all_scopes:
        creds.refresh(Request())  # 만료만 — refresh (부여 스코프 ⊇ 요청)
        token_path.write_text(creds.to_json(), encoding="utf-8")
    else:
        # 토큰 없음/무효 또는 신규 스코프(drive.file) 미부여 → 전체 스코프로 재동의.
        flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
        kwargs = {"port": 0}
        if email_hint:
            kwargs["login_hint"] = email_hint
        creds = flow.run_local_server(**kwargs)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    print("[OK] Google 인증 완료")


def do_validate() -> None:
    _reconfigure()
    from src.config import load_config

    c = load_config()
    print("[OK] SPC Hub 계정:", c.hub.user_id)
    print("[OK] Sheets 연결 설정 확인")


def main() -> None:
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""

    if cmd == "auth":
        email_hint = sys.argv[2] if len(sys.argv) > 2 else ""
        try:
            do_auth(email_hint)
        except Exception as e:
            print("[오류]", e)
            sys.exit(1)

    elif cmd == "validate":
        try:
            do_validate()
        except Exception as e:
            print("[오류]", e)
            sys.exit(1)

    else:
        print("Usage: python -m src.setup_helper auth [email] | validate")
        sys.exit(1)


if __name__ == "__main__":
    main()
