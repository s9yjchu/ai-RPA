"""일별 고객지표 업데이트 오케스트레이션."""

from __future__ import annotations

import logging
from datetime import date, timedelta, timezone, datetime
from pathlib import Path
from typing import Any

from .browser import BrowserSession
from .config import Config
from .excel_parser import (
    parse_member_metrics,
    parse_channel_metrics,
    parse_closing_report,
)
from .notifier import notify_data_not_ready, notify_failure, notify_success
from .olap_scraper import DataNotReadyError, login, scrape_report
from .sheets_writer import open_spreadsheet, write_hpc_daily, write_store_daily
from .state_manager import DAILY_CUTOFF_HOURS, DailyState

log = logging.getLogger(__name__)
KST = timezone(timedelta(hours=9))


def yesterday_kst() -> date:
    return datetime.now(KST).date() - timedelta(days=1)


def run_daily(config: Config, target_date: date | None = None, force: bool = False) -> None:
    """
    어제 날짜 기준으로 3개 OLAP 리포트를 다운로드하고
    Google Sheets 의 2개 시트를 업데이트합니다.

    - 이미 성공한 날짜면 건너뜁니다 (force=True 로 강제 재실행 가능).
    - 데이터 미준비(DataNotReadyError) 시 알림만 발송하고 종료.
      → Windows Task Scheduler 가 30분마다 재실행.
    - 재시도 시간 초과(>3시간) 시 최종 실패 처리 후 알림.
    """
    if target_date is None:
        target_date = yesterday_kst()

    state = DailyState(config.runtime.state_dir, target_date)

    if state.is_done and not force:
        log.info(f"[SKIP] {target_date} 이미 완료됨")
        return

    if state.should_give_up() and not force:
        log.error(f"[ABORT] {target_date} 최대 재시도 시간 초과 — 수동 확인 필요")
        # 포기 시 1회만 실패 알림 (30분마다 재기동되므로 중복 발송 방지).
        if not state.give_up_notified:
            try:
                notify_failure(
                    config.notify.gmail_credentials_path,
                    config.notify.gmail_token_path,
                    config.notify.report_sender,
                    config.notify.report_recipients,
                    target_date,
                    f"최대 재시도 시간({DAILY_CUTOFF_HOURS}시간) 초과 — 데이터 미준비 또는 반복 실패. 수동 확인 필요.",
                    state.attempts,
                )
                state.mark_give_up_notified()
                log.info("  [STATE] 포기 알림 메일 발송 완료")
            except Exception as exc:
                log.warning(f"  포기 알림 메일 발송 실패 (다음 실행에서 재시도): {exc}")
        return

    log.info(f"[START] 일별 업데이트 시작: {target_date} (시도 #{state.attempts + 1})")
    state.record_attempt()

    try:
        member_file, channel_file, closing_file = _download_all(config, target_date, state)
    except DataNotReadyError as exc:
        log.warning(f"[WAIT] {exc}")
        notify_data_not_ready(
            config.notify.gmail_credentials_path,
            config.notify.gmail_token_path,
            config.notify.report_sender,
            config.notify.report_recipients,
            target_date,
            state.attempts,
        )
        return
    except Exception as exc:
        _handle_error(config, state, target_date, str(exc))
        raise

    try:
        metrics = _parse_all(member_file, channel_file, closing_file, target_date, state)
    except Exception as exc:
        _handle_error(config, state, target_date, f"파싱 오류: {exc}")
        raise

    try:
        _write_to_sheets(config, target_date, metrics, state)
    except Exception as exc:
        _handle_error(config, state, target_date, f"Sheets 쓰기 오류: {exc}")
        raise

    state.mark_success()

    if not config.runtime.dry_run:
        notify_success(
            config.notify.gmail_credentials_path,
            config.notify.gmail_token_path,
            config.notify.report_sender,
            config.notify.report_recipients,
            target_date,
            {k: v for k, v in metrics.items() if v is not None},
        )
    log.info(f"[DONE] {target_date} 업데이트 완료")


# ── 내부 헬퍼 ─────────────────────────────────────────────────────────

def _download_all(
    config: Config,
    target_date: date,
    state: DailyState,
) -> tuple[Path, Path, Path]:
    """3개 OLAP 리포트를 순서대로 다운로드."""
    dl_dir = config.runtime.download_dir
    debug_dir = config.runtime.logs_dir / "debug"

    with BrowserSession(
        headless=config.runtime.headless,
        download_dir=dl_dir,
        debug_dir=debug_dir,
    ) as session:
        page = session.new_page()
        page = login(page, config, session)  # may return new tab if OLAP opens in popup

        member_file = scrape_report(page, "member_metrics", target_date, dl_dir, session)
        state.mark_source_done("member_metrics")

        channel_file = scrape_report(page, "channel_metrics", target_date, dl_dir, session)
        state.mark_source_done("channel_metrics")

        closing_file = scrape_report(
            page, "closing_report", target_date, dl_dir, session,
            verify_date=False,  # 일마감 리포트는 당일 날짜 포함 여부 보장 안됨
        )
        state.mark_source_done("closing_report")

    return member_file, channel_file, closing_file


def _parse_all(
    member_file: Path,
    channel_file: Path,
    closing_file: Path,
    target_date: date,
    state: DailyState,
) -> dict[str, Any]:
    member  = parse_member_metrics(member_file, target_date)
    channel = parse_channel_metrics(channel_file, target_date)
    closing = parse_closing_report(closing_file)

    return {**member, **channel, **closing}


def _write_to_sheets(
    config: Config,
    target_date: date,
    metrics: dict[str, Any],
    state: DailyState,
) -> None:
    hpc_fields = {"신규회원수", "해피앱 로그인수", "해피앱 DAU", "해피오더 DAU"}
    hpc_metrics = {k: v for k, v in metrics.items() if k in hpc_fields}

    store_fields = {
        "POS 총매출액", "POS 영수증건수", "POS 거래점포수",
        "HPC 매출액", "HPC 거래점포수", "HPC 총적립액", "HPC 적립건수",
        "객단가", "HPC 총사용액", "HPC 사용건수", "APP 제시건수",
    }
    store_metrics = {k: v for k, v in metrics.items() if k in store_fields}

    if config.runtime.dry_run:
        log.info(f"[DRY_RUN] HPC 일별 쓰기 생략: {hpc_metrics}")
        log.info(f"[DRY_RUN] 전사 매장실적 쓰기 생략: {store_metrics}")
        state.mark_sheet_written("HPC 실적 (일별)")
        state.mark_sheet_written("전사 매장실적 (일별)")
        return

    spreadsheet = open_spreadsheet(
        config.sheets.credentials_path,
        config.sheets.token_path,
        config.sheets.spreadsheet_id,
    )
    write_hpc_daily(spreadsheet, target_date, hpc_metrics, dry_run=False)
    state.mark_sheet_written("HPC 실적 (일별)")

    write_store_daily(spreadsheet, target_date, store_metrics, dry_run=False)
    state.mark_sheet_written("전사 매장실적 (일별)")


def _handle_error(config: Config, state: DailyState, target_date: date, reason: str) -> None:
    state.mark_failed()
    if state.should_give_up():
        notify_failure(
            config.notify.gmail_credentials_path,
            config.notify.gmail_token_path,
            config.notify.report_sender,
            config.notify.report_recipients,
            target_date,
            reason,
            state.attempts,
        )
