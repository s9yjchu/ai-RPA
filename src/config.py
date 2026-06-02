"""환경변수 로딩 및 설정 객체 구성."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


def _get(name: str, default: str = "", required: bool = False) -> str:
    val = os.getenv(name, default)
    if required and not val:
        raise RuntimeError(f"환경변수 {name}이(가) 설정되지 않았습니다.")
    return val or ""


def _get_bool(name: str, default: bool = False) -> bool:
    raw = os.getenv(name, "").strip().lower()
    return raw in ("1", "true", "yes", "y", "on") if raw else default


@dataclass(frozen=True)
class HubConfig:
    base_url: str
    user_id: str
    password: str


@dataclass(frozen=True)
class OlapConfig:
    base_url: str


@dataclass(frozen=True)
class SheetsConfig:
    spreadsheet_id: str
    credentials_path: Path
    token_path: Path


@dataclass(frozen=True)
class NotifyConfig:
    gmail_credentials_path: Path
    gmail_token_path: Path
    report_recipients: list[str]
    report_sender: str


@dataclass(frozen=True)
class RuntimeConfig:
    headless: bool
    dry_run: bool
    download_dir: Path
    state_dir: Path
    logs_dir: Path


@dataclass(frozen=True)
class LogReportConfig:
    base_url: str


@dataclass(frozen=True)
class VisualReportConfig:
    base_url: str


@dataclass(frozen=True)
class GcpConfig:
    project_id: str
    daily_subscription: str
    monthly_subscription: str
    service_account_path: Path  # Pub/Sub 전용 SA JSON (Sheets/Gmail 과 별개)


@dataclass(frozen=True)
class Config:
    hub: HubConfig
    olap: OlapConfig
    log_report: LogReportConfig
    visual_report: VisualReportConfig
    sheets: SheetsConfig
    notify: NotifyConfig
    runtime: RuntimeConfig
    gcp: GcpConfig


def load_config() -> Config:
    recipients_raw = _get("REPORT_RECIPIENTS")

    return Config(
        hub=HubConfig(
            base_url=_get(
                "SPCHUB_BASE_URL",
                "https://hub.spc.co.kr/ekp/view/login/userLogin",
            ),
            user_id=_get("SPCHUB_ID", required=True),
            password=_get("SPCHUB_PW", required=True),
        ),
        olap=OlapConfig(
            base_url=_get(
                "OLAP_BASE_URL",
                "https://dwweb.spc.co.kr:7980/SASHBI/main.jsp"
                "?board=/01.%20SPC/01.%20OLAP/100.NEW_REPORT",
            ),
        ),
        log_report=LogReportConfig(
            base_url=_get(
                "LOG_REPORT_URL",
                "https://hplog.spc.co.kr:8000/datastory/home",
            ),
        ),
        visual_report=VisualReportConfig(
            base_url=_get(
                "VISUAL_REPORT_URL",
                "https://va.spc.co.kr/SASReportViewer/",
            ),
        ),
        sheets=SheetsConfig(
            spreadsheet_id=_get(
                "SHEETS_SPREADSHEET_ID",
                "1gIEbHzyfh4TG21M1etxWxh0u99ylGMsX4B_VGEu3Yxw",
            ),
            credentials_path=Path(
                _get("SHEETS_CREDENTIALS_PATH", "./credentials.json")
            ).resolve(),
            token_path=Path(
                _get("SHEETS_TOKEN_PATH", "./token.json")
            ).resolve(),
        ),
        notify=NotifyConfig(
            gmail_credentials_path=Path(
                _get("GMAIL_CREDENTIALS_PATH", "./credentials.json")
            ).resolve(),
            gmail_token_path=Path(
                _get("GMAIL_TOKEN_PATH", "./token.json")
            ).resolve(),
            report_recipients=[
                r.strip() for r in recipients_raw.split(",") if r.strip()
            ],
            report_sender=_get("REPORT_SENDER"),
        ),
        runtime=RuntimeConfig(
            headless=_get_bool("HEADLESS", True),
            dry_run=_get_bool("DRY_RUN", False),
            download_dir=Path(_get("DOWNLOAD_DIR", "./downloads")).resolve(),
            state_dir=Path(_get("STATE_DIR", "./state")).resolve(),
            logs_dir=Path(_get("LOGS_DIR", "./logs")).resolve(),
        ),
        gcp=GcpConfig(
            project_id=_get("GCP_PROJECT_ID"),
            daily_subscription=_get("PUBSUB_DAILY_SUBSCRIPTION", "ai-rpa-daily-sub"),
            monthly_subscription=_get("PUBSUB_MONTHLY_SUBSCRIPTION", "ai-rpa-monthly-sub"),
            service_account_path=Path(
                _get("GOOGLE_APPLICATION_CREDENTIALS", "./service_account.json")
            ).resolve(),
        ),
    )
