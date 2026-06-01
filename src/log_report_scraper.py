"""LOG REPORT 자동화 — 해피앱 월 로그인 회원수(순 로그인 회원수) 추출.

URL: https://hplog.spc.co.kr:8000/datastory/home
경로: 해피앱 > 종합 > 누적추이 (유저 접속 추이) > 월별 보기 > 해당 월 선택

# ── 셀렉터 튜닝 가이드 ───────────────────────────────────────────────
# HEADLESS=0, DRY_RUN=1 로 실행하면 브라우저가 열립니다.
# logs/debug/ 폴더의 스크린샷(.png)·HTML(.html)로 실제 DOM 확인 후
# 아래 SEL_LR_* 상수를 수정하세요.
# ────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import logging
import time
from pathlib import Path

from playwright.sync_api import Page, TimeoutError as PwTimeout

from .browser import BrowserSession
from .config import Config
from .olap_scraper import DataNotReadyError
from . import hub_login

log = logging.getLogger(__name__)

# ── 셀렉터 상수 (실제 DOM 확인 후 수정) ──────────────────────────────
# 로그인
SEL_LR_LOGIN_ID  = "input[name='userId'], input[id='userId'], input[name='id'], input[type='text']"
SEL_LR_LOGIN_PW  = "input[name='password'], input[id='password'], input[type='password']"
SEL_LR_LOGIN_BTN = "button[type='submit'], input[type='submit'], button:has-text('로그인')"

# 좌측 메뉴 탐색
SEL_LR_NAV_HAPPYAPP = (
    "a:has-text('해피앱'), span:has-text('해피앱'), li:has-text('해피앱')"
)
SEL_LR_NAV_SUMMARY = (
    "a:has-text('종합'), span:has-text('종합'), li:has-text('종합')"
)
SEL_LR_NAV_TREND = (
    "a:has-text('누적추이'), a:has-text('유저 접속 추이'), "
    "span:has-text('누적추이'), span:has-text('유저 접속 추이')"
)

# 월별 보기 전환 — 설정/기간 아이콘 or 탭
SEL_LR_PERIOD_SETTINGS = (
    "button[aria-label*='설정'], button[title*='설정'], "
    "img[title*='설정'], span[title*='기간'], button:has-text('주별')"
)
SEL_LR_MONTHLY_VIEW = (
    "li:has-text('월별'), a:has-text('월별'), button:has-text('월별'), "
    "option:has-text('월별')"
)

# 월 선택 피커
SEL_LR_MONTH_PICKER = (
    "select[name*='month'], select[id*='month'], "
    "input[placeholder*='월'], button[aria-label*='월']"
)

# 순 로그인 회원수 셀 — 테이블 행 레이블로 탐색
SEL_LR_LOGIN_COUNT_ROW = (
    "td:has-text('순 로그인 회원수'), td:has-text('신규 로그인 회원수'), "
    "th:has-text('순 로그인'), tr:has-text('순 로그인 회원수')"
)


def scrape_login_count(config: Config, year: int, month: int) -> int:
    """LOG REPORT 에서 해당 연월의 순 로그인 회원수를 반환합니다.

    데이터 미준비 시 DataNotReadyError 를 raise 합니다.
    """
    dl_dir = config.runtime.download_dir
    debug_dir = config.runtime.logs_dir / "debug"

    with BrowserSession(
        headless=config.runtime.headless,
        download_dir=dl_dir,
        debug_dir=debug_dir,
    ) as session:
        page = session.new_page()
        hub_login.login_to_hub(page, config, session)
        page = hub_login.navigate_to_log_report(page, config, session)
        _navigate_to_trend(page, session)
        _switch_to_monthly(page, session)
        _select_month(page, year, month, session)
        return _extract_login_count(page, year, month, session)


# ── 내부 함수 ─────────────────────────────────────────────────────────

def _login(page: Page, base_url: str, user_id: str, password: str, session: BrowserSession) -> None:
    log.info("[STEP] LOG REPORT 로그인 페이지 접속")
    page.goto(base_url, wait_until="domcontentloaded", timeout=30_000)
    session.snapshot(page, "lr_01_login_page")

    try:
        id_input = page.locator(SEL_LR_LOGIN_ID).first
        id_input.wait_for(state="visible", timeout=10_000)
        id_input.fill(user_id)
        page.locator(SEL_LR_LOGIN_PW).first.fill(password)
        page.locator(SEL_LR_LOGIN_BTN).first.click(timeout=5_000)
        page.wait_for_load_state("networkidle", timeout=30_000)
        log.info("  로그인 완료")
        session.snapshot(page, "lr_02_after_login")
    except PwTimeout as exc:
        session.snapshot(page, "lr_02_login_timeout")
        raise RuntimeError(f"LOG REPORT 로그인 실패 (셀렉터 확인 필요): {exc}") from exc


def _navigate_to_trend(page: Page, session: BrowserSession) -> None:
    log.info("[STEP] 해피앱 > 종합 > 누적추이 탐색")
    for label, sel, step in [
        ("해피앱",    SEL_LR_NAV_HAPPYAPP, "lr_03a_happyapp"),
        ("종합",      SEL_LR_NAV_SUMMARY,  "lr_03b_summary"),
        ("누적추이",  SEL_LR_NAV_TREND,    "lr_03c_trend"),
    ]:
        try:
            node = page.locator(sel).first
            node.wait_for(state="visible", timeout=10_000)
            node.click()
            time.sleep(0.8)
            log.info(f"  클릭: {label}")
            session.snapshot(page, step)
        except PwTimeout:
            session.snapshot(page, f"lr_nav_fail_{label}")
            raise RuntimeError(
                f"'{label}' 메뉴를 찾을 수 없습니다. "
                f"logs/debug/{step}.png 를 확인해 셀렉터를 수정하세요."
            )
    page.wait_for_load_state("networkidle", timeout=30_000)


def _switch_to_monthly(page: Page, session: BrowserSession) -> None:
    log.info("[STEP] 월별 보기 전환")
    try:
        btn = page.locator(SEL_LR_PERIOD_SETTINGS).first
        btn.wait_for(state="visible", timeout=8_000)
        btn.click()
        time.sleep(0.5)
        page.locator(SEL_LR_MONTHLY_VIEW).first.click(timeout=5_000)
        page.wait_for_load_state("networkidle", timeout=20_000)
        log.info("  월별 보기 전환 완료")
        session.snapshot(page, "lr_04_monthly_view")
    except PwTimeout:
        session.snapshot(page, "lr_04_monthly_fail")
        raise RuntimeError(
            "월별 보기 전환 실패 (SEL_LR_PERIOD_SETTINGS / SEL_LR_MONTHLY_VIEW 확인 필요)"
        )


def _select_month(page: Page, year: int, month: int, session: BrowserSession) -> None:
    log.info(f"[STEP] {year}-{month:02d} 월 선택")
    try:
        picker = page.locator(SEL_LR_MONTH_PICKER).first
        picker.wait_for(state="visible", timeout=8_000)
        tag = picker.evaluate("el => el.tagName.toLowerCase()")
        if tag == "select":
            picker.select_option(label=f"{year}년 {month}월", timeout=5_000)
        else:
            picker.fill(f"{year}-{month:02d}")
            picker.press("Enter")
        page.wait_for_load_state("networkidle", timeout=20_000)
        log.info(f"  {year}-{month:02d} 선택 완료")
        session.snapshot(page, "lr_05_month_selected")
    except PwTimeout:
        session.snapshot(page, "lr_05_month_fail")
        raise RuntimeError(
            f"{year}-{month:02d} 월 선택 실패 (SEL_LR_MONTH_PICKER 확인 필요)"
        )


def _extract_login_count(page: Page, year: int, month: int, session: BrowserSession) -> int:
    log.info("[STEP] 순 로그인 회원수 추출")
    session.snapshot(page, "lr_06_before_extract")

    try:
        row_el = page.locator(SEL_LR_LOGIN_COUNT_ROW).first
        row_el.wait_for(state="visible", timeout=10_000)

        # 해당 행의 값 셀 — 레이블 다음 td (형제 탐색)
        # 단일 값이 있는 경우와 여러 달 열이 있는 경우 모두 대응
        row_html = row_el.evaluate("el => el.closest('tr') ? el.closest('tr').innerText : el.innerText")
        log.debug(f"  행 텍스트: {row_html!r}")

        # 숫자만 추출 (쉼표 포함)
        import re
        numbers = re.findall(r"[\d,]+", row_html.replace(" ", ""))
        if not numbers:
            raise DataNotReadyError(
                f"순 로그인 회원수 행에서 숫자를 찾을 수 없습니다 ({year}-{month:02d})"
            )

        # 여러 숫자가 있으면 레이블 제외한 첫 번째 유효한 숫자 사용
        for raw in numbers:
            val = int(raw.replace(",", ""))
            if val > 0:
                log.info(f"  순 로그인 회원수: {val:,}")
                return val

        raise DataNotReadyError(
            f"순 로그인 회원수 값이 0 또는 미준비 상태입니다 ({year}-{month:02d})"
        )

    except PwTimeout:
        session.snapshot(page, "lr_06_extract_fail")
        raise DataNotReadyError(
            f"순 로그인 회원수 셀을 찾을 수 없습니다 ({year}-{month:02d}). "
            "데이터 미준비 또는 SEL_LR_LOGIN_COUNT_ROW 셀렉터 확인 필요."
        )
