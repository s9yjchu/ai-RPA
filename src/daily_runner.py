"""일별 고객지표 업데이트 오케스트레이션."""

from __future__ import annotations

import logging
from datetime import date, timedelta, timezone, datetime
from pathlib import Path
from typing import Any

from .browser import BrowserSession
from .config import Config
from .excel_parser import (
    ClosingDateNotFoundError,
    parse_member_metrics,
    parse_channel_metrics,
    parse_closing_report,
)
from .integrity import current_month_window, find_blank_dates
from .log_uploader import upload_logs
from .notifier import (
    collect_log_artifacts,
    notify_data_not_ready,
    notify_failure,
    notify_success,
)
from .olap_scraper import DataNotReadyError, login, scrape_report
from .sheets_writer import (
    SHEET_HPC_DAILY,
    SHEET_STORE_DAILY,
    open_spreadsheet,
    write_hpc_daily,
    write_store_daily,
)
from .state_manager import DAILY_CUTOFF_HOURS, DailyState

# 일별 자동 갭필 1회당 최대 처리 날짜 수(OLAP 부하·스케줄 창 보호). 남은 날짜는 다음 실행이 처리.
MAX_AUTO_BACKFILL = 7
# 자동 갭필 트리거 컬럼: HPC 는 신규회원수(4) 제외(원천 지연으로 상시 공백 가능 → 무한 재처리 방지).
# 매장은 전체 RPA 컬럼. (수동 backfill 은 전체 컬럼 강제 처리.)
_AUTO_BLANK_SPECS = [
    (SHEET_HPC_DAILY, [5, 6, 7]),
    (SHEET_STORE_DAILY, [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14]),
]

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
                _send_failure(
                    config, state, target_date,
                    f"최대 재시도 시간({DAILY_CUTOFF_HOURS}시간) 초과 — 데이터 미준비 또는 반복 실패. 수동 확인 필요.",
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

    # 주 날짜 성공 후, 당월의 RPA 공백 날짜를 best-effort 로 채움(여러 실행에 걸쳐 수렴).
    _auto_backfill(config, exclude=target_date)


# ── 내부 헬퍼 ─────────────────────────────────────────────────────────

def _download_one(
    page,
    config: Config,
    target_date: date,
    session: BrowserSession,
    state: DailyState | None = None,
) -> tuple[Path, Path, Path]:
    """이미 로그인된 page 로 한 날짜의 3개 리포트를 다운로드(세션 재사용용)."""
    dl_dir = config.runtime.download_dir

    member_file = scrape_report(page, "member_metrics", target_date, dl_dir, session)
    if state:
        state.mark_source_done("member_metrics")

    channel_file = scrape_report(page, "channel_metrics", target_date, dl_dir, session)
    if state:
        state.mark_source_done("channel_metrics")

    closing_file = scrape_report(
        page, "closing_report", target_date, dl_dir, session,
        verify_date=True,  # 일마감 날짜는 유니코드 대시 정규화로 정상 검증됨
    )
    if state:
        state.mark_source_done("closing_report")

    return member_file, channel_file, closing_file


def _download_all(
    config: Config,
    target_date: date,
    state: DailyState,
) -> tuple[Path, Path, Path]:
    """단일 날짜용: 세션 열기 → 로그인 → 3개 리포트 다운로드."""
    debug_dir = config.runtime.logs_dir / "debug"
    with BrowserSession(
        headless=config.runtime.headless,
        download_dir=config.runtime.download_dir,
        debug_dir=debug_dir,
    ) as session:
        page = session.new_page()
        page = login(page, config, session)  # may return new tab if OLAP opens in popup
        return _download_one(page, config, target_date, session, state)


def _parse_all(
    member_file: Path,
    channel_file: Path,
    closing_file: Path,
    target_date: date,
    state: DailyState,
) -> dict[str, Any]:
    member  = parse_member_metrics(member_file, target_date)
    channel = parse_channel_metrics(channel_file, target_date)
    try:
        closing = parse_closing_report(closing_file, target_date)  # target_date 열그룹 고정
    except ClosingDateNotFoundError as exc:
        # 일마감에 대상일 열그룹이 없음 → 일시적 미준비로 간주(재시도 흐름).
        raise DataNotReadyError(str(exc)) from exc

    return {**member, **channel, **closing}


_HPC_FIELDS = {"신규회원수", "해피앱 로그인수", "해피앱 DAU", "해피오더 DAU"}
_STORE_FIELDS = {
    "POS 총매출액", "POS 영수증건수", "POS 거래점포수",
    "HPC 매출액", "HPC 거래점포수", "HPC 총적립액", "HPC 적립건수",
    "객단가", "HPC 총사용액", "HPC 사용건수", "APP 제시건수",
}


def _split_metrics(metrics: dict[str, Any]) -> tuple[dict, dict]:
    hpc = {k: v for k, v in metrics.items() if k in _HPC_FIELDS}
    store = {k: v for k, v in metrics.items() if k in _STORE_FIELDS}
    return hpc, store


def _write_both(
    spreadsheet,
    target_date: date,
    metrics: dict[str, Any],
    dry_run: bool,
    state: DailyState | None = None,
) -> None:
    """이미 열린 spreadsheet 로 두 일별 시트에 기입(세션·인증 재사용)."""
    hpc_metrics, store_metrics = _split_metrics(metrics)
    if dry_run:
        log.info(f"[DRY_RUN] HPC 일별 쓰기 생략: {hpc_metrics}")
        log.info(f"[DRY_RUN] 전사 매장실적 쓰기 생략: {store_metrics}")
    else:
        write_hpc_daily(spreadsheet, target_date, hpc_metrics, dry_run=False)
        write_store_daily(spreadsheet, target_date, store_metrics, dry_run=False)
    if state:
        state.mark_sheet_written("HPC 실적 (일별)")
        state.mark_sheet_written("전사 매장실적 (일별)")


def _open_sheets(config: Config):
    return open_spreadsheet(
        config.sheets.credentials_path,
        config.sheets.token_path,
        config.sheets.spreadsheet_id,
    )


def _write_to_sheets(
    config: Config,
    target_date: date,
    metrics: dict[str, Any],
    state: DailyState,
) -> None:
    spreadsheet = None if config.runtime.dry_run else _open_sheets(config)
    _write_both(spreadsheet, target_date, metrics, config.runtime.dry_run, state)


def _handle_error(config: Config, state: DailyState, target_date: date, reason: str) -> None:
    state.mark_failed()
    if state.should_give_up():
        _send_failure(config, state, target_date, reason)


def _send_failure(config: Config, state: DailyState, target_date: date, reason: str) -> None:
    """실패 알림 메일(로그·스냅샷 첨부) + 구글 드라이브 업로드 (보조)."""
    artifacts = collect_log_artifacts(config.runtime.logs_dir)
    notify_failure(
        config.notify.gmail_credentials_path,
        config.notify.gmail_token_path,
        config.notify.report_sender,
        config.notify.report_recipients,
        target_date,
        reason,
        state.attempts,
        attachments=artifacts,
    )
    upload_logs(config, f"daily_{target_date}", artifacts)


# ── 백필 (여러 날짜 일괄 처리) ────────────────────────────────────────

def run_backfill(
    config: Config,
    dates: list[date],
    force: bool = True,
) -> dict[str, list[date]]:
    """여러 날짜를 한 세션(로그인 1회)으로 처리. 각 날짜는 독립적으로 격리.

    - force=True(기본): 이미 성공한 날짜도 재처리(잘못 쓰인 값 교정용).
    - 데이터 미준비/대상일 미존재 날짜는 건너뜀(skip), 기타 오류는 fail 로 기록.
    반환: {"ok": [...], "skip": [...], "fail": [...]}
    """
    dates = sorted(set(dates))
    summary: dict[str, list[date]] = {"ok": [], "skip": [], "fail": []}
    if not dates:
        log.info("[BACKFILL] 대상 날짜 없음")
        return summary

    log.info(f"[BACKFILL] {len(dates)}일 처리 시작: {dates[0]} ~ {dates[-1]}")
    debug_dir = config.runtime.logs_dir / "debug"
    spreadsheet = None if config.runtime.dry_run else _open_sheets(config)

    with BrowserSession(
        headless=config.runtime.headless,
        download_dir=config.runtime.download_dir,
        debug_dir=debug_dir,
    ) as session:
        page = session.new_page()
        page = login(page, config, session)

        for d in dates:
            state = DailyState(config.runtime.state_dir, d)
            if state.is_done and not force:
                summary["skip"].append(d)
                continue
            try:
                state.record_attempt()
                mf, cf, zf = _download_one(page, config, d, session, state)
                metrics = _parse_all(mf, cf, zf, d, state)
                _write_both(spreadsheet, d, metrics, config.runtime.dry_run, state)
                state.mark_success()
                summary["ok"].append(d)
                log.info(f"  [BACKFILL] {d} 완료")
            except DataNotReadyError as exc:
                log.warning(f"  [BACKFILL] {d} 데이터 미준비 — 건너뜀: {exc}")
                summary["skip"].append(d)
            except Exception as exc:
                log.error(f"  [BACKFILL] {d} 실패: {exc}", exc_info=True)
                state.mark_failed()
                summary["fail"].append(d)

    log.info(
        f"[BACKFILL] 완료 — 성공 {len(summary['ok'])} / 건너뜀 {len(summary['skip'])} "
        f"/ 실패 {len(summary['fail'])}"
    )
    return summary


def _auto_backfill(config: Config, exclude: date) -> None:
    """당월 RPA 공백 날짜를 best-effort 로 채움. 메인 흐름을 절대 막지 않는다(예외 격리)."""
    if config.runtime.dry_run:
        return
    try:
        spreadsheet = _open_sheets(config)
        start, end = current_month_window()
        blanks = find_blank_dates(spreadsheet, start, end, specs=_AUTO_BLANK_SPECS)
        targets = sorted(d for d in blanks if d != exclude)
        if not targets:
            return
        capped = targets[:MAX_AUTO_BACKFILL]
        log.info(
            f"[AUTO-BACKFILL] 당월 공백 {len(targets)}일 중 {len(capped)}일 처리: {capped}"
        )
        run_backfill(config, capped, force=True)
    except Exception as exc:
        log.warning(f"[AUTO-BACKFILL] 건너뜀(무시): {exc}")
