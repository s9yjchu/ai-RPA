"""SPC Hub SSO 로그인 — 모든 내부 시스템 공통 진입점.

실제 로그인 플로우:
  1. Hub 로그인 (hub.spc.co.kr) → 확인 버튼 클릭
  2. 상단 nav "System Link" → "정보화시스템" 클릭
  3. 정보화시스템 메뉴 페이지에서 OLAP / LOG REPORT / VISUAL REPORT 링크 클릭
  4. VISUAL REPORT 는 로그인 페이지 표시 시 "OLAP 계정으로 로그인" 클릭

# ── 셀렉터 튜닝 가이드 ───────────────────────────────────────────────
# HEADLESS=0, DRY_RUN=1 로 실행 → logs/debug/hub_*.png 확인
# 브라우저 F12 → Elements → 요소 우클릭 → Copy selector
# ────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import logging
import time

from playwright.sync_api import Page, TimeoutError as PwTimeout

from .browser import BrowserSession

log = logging.getLogger(__name__)

# ── Hub 로그인 셀렉터 ────────────────────────────────────────────────
SEL_HUB_ID  = "input[name='userId'], input[id='userId'], input[name='id']"
SEL_HUB_PW  = "input[name='password'], input[id='password'], input[type='password']"
SEL_HUB_BTN = "button[type='submit'], input[type='submit'], button:has-text('로그인')"

# 로그인 직후 확인/동의 버튼 (없으면 건너뜀)
SEL_HUB_CONFIRM = (
    "button:has-text('확인'), input[value='확인'], a:has-text('확인'), "
    "button:has-text('동의'), a:has-text('동의')"
)

# ── 정보화시스템 메뉴 탐색 셀렉터 ────────────────────────────────────
# 상단 네비게이션 "System Link" 드롭다운
SEL_NAV_SYSTEM_LINK = (
    "a:has-text('System Link'), span:has-text('System Link'), "
    "li:has-text('System Link'), a[title*='System']"
)
# 드롭다운 내 "정보화시스템" 항목
SEL_NAV_INFO_SYSTEM = (
    "a:has-text('정보화시스템'), li:has-text('정보화시스템'), "
    "span:has-text('정보화시스템')"
)

# ── 정보화시스템 메뉴 페이지 링크 ────────────────────────────────────
SEL_MENU_OLAP = (
    "a:has-text('OLAP'), a[href*='SASHBI'], a[href*='sashbi'], "
    "a[href*='dwweb'], td:has-text('OLAP') a"
)
SEL_MENU_LOG_REPORT = (
    "a:has-text('LOG REPORT'), a:has-text('Log Report'), "
    "a[href*='hplog'], a:has-text('로그 리포트')"
)
SEL_MENU_VISUAL_REPORT = (
    "a:has-text('VISUAL REPORT'), a:has-text('Visual Report'), "
    "a[href*='va.spc'], a:has-text('비주얼 리포트')"
)

# VISUAL REPORT 로그인 페이지 하단 "OLAP 계정으로 로그인" 버튼
SEL_VR_OLAP_LOGIN = (
    "a:has-text('OLAP 계정으로 로그인'), button:has-text('OLAP 계정으로 로그인'), "
    "a:has-text('OLAP'), a[href*='olap']"
)

# ── 공개 함수 ─────────────────────────────────────────────────────────

def login_to_hub(page: Page, config, session: BrowserSession) -> None:
    """SPC Hub 에 로그인하고 메인 페이지에 도달합니다."""
    log.info("[STEP] SPC Hub 로그인")
    page.goto(config.hub.base_url, wait_until="domcontentloaded", timeout=30_000)
    session.snapshot(page, "hub_01_login_page")

    try:
        id_inp = page.locator(SEL_HUB_ID).first
        id_inp.wait_for(state="visible", timeout=10_000)
        id_inp.fill(config.hub.user_id)
        page.locator(SEL_HUB_PW).first.fill(config.hub.password)
        page.locator(SEL_HUB_BTN).first.click(timeout=5_000)
        page.wait_for_load_state("networkidle", timeout=30_000)
        log.info("  로그인 버튼 클릭 완료")
        session.snapshot(page, "hub_02_after_login")
    except PwTimeout as exc:
        session.snapshot(page, "hub_02_login_fail")
        raise RuntimeError(
            f"SPC Hub 로그인 실패 (SEL_HUB_* 셀렉터 확인 필요): {exc}"
        ) from exc

    # 확인/동의 버튼 — 있으면 클릭, 없으면 건너뜀
    try:
        confirm = page.locator(SEL_HUB_CONFIRM).first
        confirm.wait_for(state="visible", timeout=5_000)
        confirm.click()
        page.wait_for_load_state("networkidle", timeout=20_000)
        log.info("  확인 버튼 클릭")
        session.snapshot(page, "hub_03_after_confirm")
    except PwTimeout:
        log.info("  확인 버튼 없음 — 건너뜀")

    log.info("  Hub 로그인 완료")


def navigate_to_olap(page: Page, config, session: BrowserSession) -> Page:
    """정보화시스템 메뉴를 통해 OLAP 에 진입하고 활성 Page 를 반환합니다."""
    return _open_via_menu(page, SEL_MENU_OLAP, "OLAP", config, session)


def navigate_to_log_report(page: Page, config, session: BrowserSession) -> Page:
    """정보화시스템 메뉴를 통해 LOG REPORT 에 진입하고 활성 Page 를 반환합니다."""
    return _open_via_menu(page, SEL_MENU_LOG_REPORT, "LOG REPORT", config, session)


def navigate_to_visual_report(page: Page, config, session: BrowserSession) -> Page:
    """정보화시스템 메뉴를 통해 VISUAL REPORT 에 진입합니다.

    VISUAL REPORT 로그인 페이지가 표시되면 'OLAP 계정으로 로그인' 을 클릭합니다.
    """
    active = _open_via_menu(page, SEL_MENU_VISUAL_REPORT, "VISUAL REPORT", config, session)

    # VISUAL REPORT 가 자체 로그인 페이지를 보여주는 경우 처리
    try:
        olap_btn = active.locator(SEL_VR_OLAP_LOGIN).first
        olap_btn.wait_for(state="visible", timeout=6_000)
        olap_btn.click()
        active.wait_for_load_state("networkidle", timeout=30_000)
        log.info("  'OLAP 계정으로 로그인' 클릭 완료")
        session.snapshot(active, "hub_vr_after_olap_login")
    except PwTimeout:
        log.info("  VISUAL REPORT 로그인 페이지 없음 — SSO 자동 인증됨")

    return active


# ── 내부 헬퍼 ─────────────────────────────────────────────────────────

def _open_info_system_menu(page: Page, session: BrowserSession) -> None:
    """Hub 메인에서 System Link 드롭다운 → 정보화시스템 클릭."""
    log.info("[STEP] 정보화시스템 메뉴 탐색")
    try:
        sys_link = page.locator(SEL_NAV_SYSTEM_LINK).first
        sys_link.wait_for(state="visible", timeout=10_000)
        sys_link.click()
        time.sleep(0.5)
        page.locator(SEL_NAV_INFO_SYSTEM).first.click(timeout=8_000)
        page.wait_for_load_state("networkidle", timeout=20_000)
        log.info("  정보화시스템 메뉴 페이지 도달")
        session.snapshot(page, "hub_04_info_system_menu")
    except PwTimeout as exc:
        session.snapshot(page, "hub_04_menu_fail")
        raise RuntimeError(
            f"정보화시스템 메뉴 접근 실패 (SEL_NAV_* 셀렉터 확인 필요): {exc}"
        ) from exc


def _open_via_menu(
    page: Page,
    link_selector: str,
    label: str,
    config,
    session: BrowserSession,
) -> Page:
    """정보화시스템 메뉴에서 시스템 링크를 클릭하고 활성 Page 를 반환합니다.

    링크가 새 탭(popup)으로 열리면 해당 탭을, 같은 탭이면 현재 page 를 반환합니다.
    """
    _open_info_system_menu(page, session)

    log.info(f"  {label} 링크 클릭")
    try:
        link = page.locator(link_selector).first
        link.wait_for(state="visible", timeout=10_000)

        # 새 탭으로 열리는 경우를 expect_popup 으로 감지
        pages_before = len(page.context.pages)
        link.click()
        time.sleep(1.5)  # popup 열릴 시간 대기

        pages_after = page.context.pages
        if len(pages_after) > pages_before:
            # 새 탭으로 열림
            active = pages_after[-1]
            active.wait_for_load_state("networkidle", timeout=30_000)
            log.info(f"  {label} 새 탭에서 열림: {active.url[:60]}")
        else:
            # 같은 탭에서 열림
            page.wait_for_load_state("networkidle", timeout=30_000)
            active = page
            log.info(f"  {label} 현재 탭에서 열림: {active.url[:60]}")

        session.snapshot(active, f"hub_05_{label.lower().replace(' ', '_')}_opened")
        return active

    except PwTimeout as exc:
        session.snapshot(page, f"hub_05_{label.lower().replace(' ', '_')}_fail")
        raise RuntimeError(
            f"{label} 링크 클릭 실패 (셀렉터 확인 필요): {exc}"
        ) from exc
