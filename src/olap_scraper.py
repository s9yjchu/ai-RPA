"""SASHBI OLAP 시스템 자동화 — 로그인, 리포트 탐색, 날짜 필터, 엑셀 다운로드.

# ── 셀렉터 튜닝 가이드 ───────────────────────────────────────────────
# HEADLESS=0, DRY_RUN=1 로 실행하면 브라우저가 열립니다.
# 각 단계에서 logs/debug/ 폴더에 스크린샷(.png)과 HTML(.html)이 저장됩니다.
# 해당 파일로 실제 DOM 구조를 확인한 뒤 아래 SELECTORS를 수정하세요.
#
# 빠른 탐색 방법:
#   브라우저 개발자 도구(F12) → Elements 탭 → 해당 요소 우클릭 → Copy selector
# ────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import logging
import time
from datetime import date, timedelta, timezone, datetime
from pathlib import Path
from typing import Optional

from playwright.sync_api import Page, TimeoutError as PwTimeout

from .browser import BrowserSession

log = logging.getLogger(__name__)
KST = timezone(timedelta(hours=9))

# ── 셀렉터 상수 (실제 DOM 확인 후 수정) ──────────────────────────────
# 로그인 페이지
SEL_LOGIN_ID = "input[name='userId'], input[name='id'], input#userId, input#id"
SEL_LOGIN_PW = "input[name='password'], input[name='passwd'], input#password, input#passwd"
SEL_LOGIN_BTN = "button[type='submit'], input[type='submit'], button:has-text('로그인'), a:has-text('로그인')"

# 리포트 트리 (좌측 탐색 패널)
# SAS BI 트리 노드는 보통 data-* 속성이나 title 속성으로 식별
SEL_TREE_NODE = "span[title='{label}'], li[title='{label}'], a[title='{label}']"

# 날짜 필터 영역 — SAS WRS 기본 패턴
SEL_DATE_START = (
    "input[name='startDate'], input[id*='startDate'], input[placeholder*='시작']"
)
SEL_DATE_END = (
    "input[name='endDate'], input[id*='endDate'], input[placeholder*='종료'], input[placeholder*='끝']"
)
# 날짜 입력 방식 A: input에 직접 타이핑
# 날짜 입력 방식 B: 캘린더 위젯 (open_calendar → click day)
SEL_CALENDAR_OPEN = "button[aria-label*='날짜'], img[src*='cal'], button[class*='cal']"
SEL_CALENDAR_DAY = "td[data-day='{day}']:not(.disabled), button[data-day='{day}']:not(.disabled)"
SEL_CALENDAR_YESTERDAY_BTN = "button:has-text('어제'), td:has-text('{day}')"

# 실행/조회 버튼
SEL_RUN_BTN = (
    "button:has-text('실행'), button:has-text('조회'), button:has-text('검색'), "
    "input[value='실행'], input[value='조회']"
)

# 다운로드 — SAS WRS 는 toolbar 메뉴 클릭 후 서브메뉴에서 Excel 선택하는 경우가 많음
SEL_EXPORT_BTN = (
    "button[title*='내보내기'], button[title*='Export'], button[aria-label*='다운로드'], "
    "a:has-text('다운로드'), button:has-text('다운로드'), span:has-text('내보내기')"
)
SEL_EXPORT_EXCEL = (
    "li:has-text('Excel'), a:has-text('Excel'), button:has-text('Excel'), "
    "li:has-text('엑셀'), a:has-text('엑셀')"
)

# 리포트 이름 경로 (트리 클릭 순서)
REPORT_PATHS = {
    "member_metrics": [
        "정형데이터",
        "회원분석",
        "[HPC] 일별 회원관리지표",
    ],
    "channel_metrics": [
        "정형데이터",
        "[HPC] 채널별 적립, 사용건수 현황",
    ],
    "closing_report": [
        "전사",
        "마감리포트",
        "2. [HPC, POS] HPC 일마감(브랜드)",
    ],
}

# iFrame 이름/ID 힌트 — SAS WRS 가 iFrame을 사용할 경우 아래를 설정
# None 이면 iFrame 없이 메인 페이지에서 조작
REPORT_FRAME_SELECTOR: Optional[str] = "iframe#reportFrame, iframe[name='reportFrame']"


def yesterday_kst() -> date:
    return datetime.now(KST).date() - timedelta(days=1)


# ── 로그인 ────────────────────────────────────────────────────────────

def login(page: Page, base_url: str, user_id: str, password: str, session: BrowserSession) -> None:
    log.info("[STEP] OLAP 로그인 페이지 접속")
    page.goto(base_url, wait_until="domcontentloaded", timeout=30_000)
    session.snapshot(page, "01_login_page")

    try:
        page.locator(SEL_LOGIN_ID).first.fill(user_id, timeout=10_000)
        page.locator(SEL_LOGIN_PW).first.fill(password, timeout=5_000)
        page.locator(SEL_LOGIN_BTN).first.click(timeout=5_000)
        # 로그인 후 메인 화면 로딩 대기
        page.wait_for_load_state("networkidle", timeout=30_000)
        log.info("  로그인 완료")
        session.snapshot(page, "02_after_login")
    except PwTimeout as exc:
        session.snapshot(page, "02_login_timeout")
        raise RuntimeError(f"로그인 실패 (셀렉터 확인 필요): {exc}") from exc


# ── 리포트 트리 탐색 ─────────────────────────────────────────────────

def _get_active_page(page: Page) -> Page:
    """iFrame이 있으면 frame을 반환, 없으면 page 자체 반환."""
    if REPORT_FRAME_SELECTOR:
        try:
            frame = page.frame_locator(REPORT_FRAME_SELECTOR).first
            # frame이 존재하는지 확인
            frame.locator("body").wait_for(state="attached", timeout=3_000)
            return frame  # type: ignore[return-value]
        except Exception:
            pass
    return page


def navigate_to_report(page: Page, report_key: str, session: BrowserSession) -> None:
    """트리에서 순서대로 노드를 클릭하여 리포트를 엽니다."""
    path = REPORT_PATHS[report_key]
    log.info(f"[STEP] 리포트 탐색: {' > '.join(path)}")

    target = _get_active_page(page)

    for i, label in enumerate(path):
        sel = SEL_TREE_NODE.format(label=label)
        try:
            node = target.locator(sel).first
            node.wait_for(state="visible", timeout=10_000)
            node.click()
            time.sleep(0.8)  # 트리 확장 애니메이션 대기
            log.info(f"  클릭: {label}")
        except PwTimeout:
            session.snapshot(page, f"tree_fail_{i}_{label[:20]}")
            raise RuntimeError(
                f"트리 노드 '{label}' 를 찾을 수 없습니다. "
                f"logs/debug/ 스크린샷을 확인해 SEL_TREE_NODE 셀렉터를 수정하세요."
            )

    # 리포트 로딩 대기
    page.wait_for_load_state("networkidle", timeout=30_000)
    session.snapshot(page, f"03_report_loaded_{report_key}")
    log.info("  리포트 로딩 완료")


# ── 날짜 필터 ─────────────────────────────────────────────────────────

def _fill_date_input(target: Page, selector: str, target_date: date) -> bool:
    """input 에 직접 날짜를 입력. 성공하면 True."""
    date_str = target_date.strftime("%Y-%m-%d")
    alt_formats = [
        target_date.strftime("%Y-%m-%d"),
        target_date.strftime("%Y/%m/%d"),
        target_date.strftime("%Y.%m.%d"),
        target_date.strftime("%m/%d/%Y"),
    ]
    try:
        inp = target.locator(selector).first
        inp.wait_for(state="visible", timeout=5_000)
        for fmt in alt_formats:
            inp.triple_click()
            inp.fill(fmt)
            inp.press("Tab")
            time.sleep(0.3)
            val = inp.input_value()
            if target_date.strftime("%Y") in val or target_date.strftime("%m") in val:
                log.info(f"  날짜 입력 성공: {fmt}")
                return True
        log.warning(f"  날짜 직접 입력 실패 ({date_str}), 캘린더 방식 시도")
        return False
    except PwTimeout:
        return False


def _click_calendar_day(target: Page, selector_open: str, target_date: date) -> None:
    """캘린더 위젯을 열고 해당 날짜를 클릭합니다."""
    day_str = str(target_date.day)
    try:
        target.locator(selector_open).first.click(timeout=5_000)
        time.sleep(0.5)
    except PwTimeout:
        log.warning("  캘린더 열기 버튼 없음 — 직접 클릭 시도 생략")
        return

    day_sel = SEL_CALENDAR_DAY.format(day=day_str)
    try:
        target.locator(day_sel).first.click(timeout=5_000)
        log.info(f"  캘린더에서 {day_str}일 선택")
    except PwTimeout:
        # 날짜 숫자만으로 td 탐색 (fallback)
        target.locator(f"td:has-text('{day_str}')").first.click(timeout=5_000)
        log.info(f"  캘린더 fallback: td '{day_str}' 클릭")


def set_date_filter(page: Page, target_date: date, session: BrowserSession) -> None:
    """시작일, 종료일을 target_date 로 설정합니다."""
    log.info(f"[STEP] 날짜 필터 설정: {target_date}")
    target = _get_active_page(page)

    for selector, label, cal_open_sel in [
        (SEL_DATE_START, "시작일", SEL_CALENDAR_OPEN),
        (SEL_DATE_END,   "종료일", SEL_CALENDAR_OPEN),
    ]:
        ok = _fill_date_input(target, selector, target_date)
        if not ok:
            _click_calendar_day(target, cal_open_sel, target_date)

    session.snapshot(page, "04_date_filter_set")


# ── 실행 및 다운로드 ──────────────────────────────────────────────────

def click_run(page: Page, session: BrowserSession) -> None:
    log.info("[STEP] 조회 실행 버튼 클릭")
    target = _get_active_page(page)
    try:
        target.locator(SEL_RUN_BTN).first.click(timeout=10_000)
        page.wait_for_load_state("networkidle", timeout=60_000)
        log.info("  조회 완료")
        session.snapshot(page, "05_after_run")
    except PwTimeout as exc:
        session.snapshot(page, "05_run_timeout")
        raise RuntimeError(
            f"조회 버튼 클릭 실패: {exc}\n"
            "SEL_RUN_BTN 셀렉터를 확인하세요."
        ) from exc


def _verify_report_date(page: Page, target_date: date) -> bool:
    """리포트 본문에 어제 날짜가 포함되어 있는지 확인 (데이터 준비 여부 체크)."""
    date_str = target_date.strftime("%Y-%m-%d")
    alt = target_date.strftime("%Y/%m/%d")
    content = page.content()
    if date_str in content or alt in content:
        return True
    # YYYYMMDD 형식도 확인
    if target_date.strftime("%Y%m%d") in content:
        return True
    log.warning(f"  리포트에서 날짜 {date_str} 미확인 — 데이터가 준비 중일 수 있습니다.")
    return False


def download_excel(
    page: Page,
    download_dir: Path,
    report_key: str,
    session: BrowserSession,
) -> Path:
    """내보내기 메뉴를 통해 Excel 파일을 다운로드합니다."""
    log.info(f"[STEP] Excel 다운로드: {report_key}")
    target = _get_active_page(page)

    try:
        # 내보내기 버튼 클릭 → 서브메뉴 표시
        target.locator(SEL_EXPORT_BTN).first.click(timeout=10_000)
        time.sleep(0.5)
        # Excel 옵션 선택
        with page.expect_download(timeout=120_000) as dl_info:
            target.locator(SEL_EXPORT_EXCEL).first.click(timeout=10_000)

        dl = dl_info.value
        save_path = download_dir / f"{report_key}_{dl.suggested_filename}"
        dl.save_as(save_path)
        log.info(f"  저장 완료: {save_path}")
        session.snapshot(page, f"06_downloaded_{report_key}")
        return save_path

    except PwTimeout as exc:
        session.snapshot(page, f"06_download_fail_{report_key}")
        raise RuntimeError(
            f"Excel 다운로드 실패 ({report_key}): {exc}\n"
            "SEL_EXPORT_BTN / SEL_EXPORT_EXCEL 셀렉터를 확인하세요."
        ) from exc


# ── 통합 실행 함수 ────────────────────────────────────────────────────

def scrape_report(
    page: Page,
    report_key: str,
    target_date: date,
    download_dir: Path,
    session: BrowserSession,
    verify_date: bool = True,
) -> Path:
    """리포트 탐색 → 날짜 설정 → 조회 → 다운로드 전체 플로우."""
    navigate_to_report(page, report_key, session)
    set_date_filter(page, target_date, session)
    click_run(page, session)

    if verify_date and not _verify_report_date(page, target_date):
        raise DataNotReadyError(
            f"리포트 '{report_key}' 에서 {target_date} 데이터가 아직 준비되지 않았습니다."
        )

    return download_excel(page, download_dir, report_key, session)


class DataNotReadyError(RuntimeError):
    """OLAP 에서 해당 날짜 데이터가 아직 확인되지 않을 때."""
