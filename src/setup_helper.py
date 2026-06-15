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

    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
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
