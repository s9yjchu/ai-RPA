"""월별 고객지표 업데이트 — LOG REPORT(로그인객수) + VISUAL REPORT(MAU)."""

from __future__ import annotations

import logging

from .config import Config
from .log_report_scraper import scrape_login_count
from .notifier import notify_data_not_ready, notify_failure, notify_success
from .olap_scraper import DataNotReadyError
from .sheets_writer import open_spreadsheet, write_hpc_monthly
from .state_manager import MONTHLY_MAX_DAYS, MonthlyState
from .visual_report_scraper import scrape_mau

log = logging.getLogger(__name__)


def run_monthly(config: Config, year: int, month: int, force: bool = False) -> None:
    """
    지정 연월의 월별 지표를 수집하여 Google Sheets "HPC 실적 (월별)" 시트를 업데이트합니다.

    - 이미 성공한 월이면 건너뜁니다 (force=True 로 강제 재실행 가능).
    - 데이터 미준비(DataNotReadyError) 시 알림 발송 후 종료 → 다음 실행에서 재시도.
    - 재시도 기간 초과 시 최종 실패 처리.
    """
    state = MonthlyState(config.runtime.state_dir, year, month)

    if state.is_done and not force:
        log.info(f"[SKIP] {year}-{month:02d} 월별 업데이트 이미 완료됨")
        return

    if state.should_give_up() and not force:
        log.error(f"[ABORT] {year}-{month:02d} 최대 재시도 기간 초과 — 수동 확인 필요")
        # 포기 시 1회만 실패 알림 (재기동마다 중복 발송 방지).
        if not state.give_up_notified:
            try:
                notify_failure(
                    config.notify.gmail_credentials_path,
                    config.notify.gmail_token_path,
                    config.notify.report_sender,
                    config.notify.report_recipients,
                    _month_label(year, month),
                    f"최대 재시도 기간({MONTHLY_MAX_DAYS}일) 초과 — 데이터 미준비 또는 반복 실패. 수동 확인 필요.",
                    state.attempts,
                )
                state.mark_give_up_notified()
                log.info("  [STATE] 포기 알림 메일 발송 완료")
            except Exception as exc:
                log.warning(f"  포기 알림 메일 발송 실패 (다음 실행에서 재시도): {exc}")
        return

    log.info(f"[START] 월별 업데이트 시작: {year}-{month:02d} (시도 #{state.attempts + 1})")
    state.record_attempt()

    try:
        # ── Step 1: LOG REPORT → 순 로그인 회원수 ──────────────────────
        log.info("[STEP] LOG REPORT 스크래핑")
        login_count = scrape_login_count(config, year, month)
        state.mark_source_done("log_report")

        # ── Step 2: VISUAL REPORT → MAU 당월 (KPI 툴팁에서 직접 추출) ──
        log.info("[STEP] VISUAL REPORT 스크래핑")
        mau_value = scrape_mau(config, year, month)
        state.mark_source_done("visual_report")

        # ── Step 3: Google Sheets 쓰기 ─────────────────────────────────
        log.info("[STEP] Google Sheets 업데이트")
        spreadsheet = open_spreadsheet(
            config.sheets.credentials_path,
            config.sheets.token_path,
            config.sheets.spreadsheet_id,
        )
        write_hpc_monthly(
            spreadsheet,
            year,
            month,
            {
                "해피앱 월 로그인객수": login_count,
                "해피앱 MAU": mau_value,
            },
            dry_run=config.runtime.dry_run,
        )
        state.mark_sheet_written("HPC 실적 (월별)")

        state.mark_success()

        if not config.runtime.dry_run:
            notify_success(
                config.notify.gmail_credentials_path,
                config.notify.gmail_token_path,
                config.notify.report_sender,
                config.notify.report_recipients,
                _month_label(year, month),
                {
                    "해피앱 월 로그인객수": login_count,
                    "해피앱 MAU": mau_value,
                },
            )

        log.info(f"[DONE] {year}-{month:02d} 월별 업데이트 완료")

    except DataNotReadyError as exc:
        log.warning(f"[WAIT] {exc}")
        notify_data_not_ready(
            config.notify.gmail_credentials_path,
            config.notify.gmail_token_path,
            config.notify.report_sender,
            config.notify.report_recipients,
            _month_label(year, month),
            state.attempts,
        )
        # 상태를 failed 로 전환하지 않음 — 다음 실행에서 재시도

    except Exception as exc:
        reason = str(exc)
        log.error(f"[FAIL] {year}-{month:02d} 오류: {reason}", exc_info=True)
        state.mark_failed()
        if state.should_give_up():
            notify_failure(
                config.notify.gmail_credentials_path,
                config.notify.gmail_token_path,
                config.notify.report_sender,
                config.notify.report_recipients,
                _month_label(year, month),
                reason,
                state.attempts,
            )
        raise


def _month_label(year: int, month: int):
    """notifier 의 date 인자 자리에 넘길 표현값 (문자열 형태)."""
    from datetime import date
    return date(year, month, 1)
