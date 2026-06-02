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

# ── Hub 로그인 셀렉터 (hub.spc.co.kr DOM 확인 완료) ─────────────────
SEL_HUB_ID  = "input#userId"
# input#password_fake 는 display:none — .input_pw 클래스로 실제 필드만 선택
SEL_HUB_PW  = "input#password.input_pw, input#password:not([style*='display: none'])"
SEL_HUB_BTN = "button#btnLogin"

# 로그인 직후 확인/동의 버튼 (없으면 건너뜀)
SEL_HUB_CONFIRM = (
    "button:has-text('확인'), input[value='확인'], a:has-text('확인'), "
    "button:has-text('동의'), a:has-text('동의')"
)

# ── 정보화시스템 접속 ──────────────────────────────────────────────
# Hub HTML 확인: 정보화시스템 링크는 SSO 게이트를 통해 sis.spc.co.kr 으로 연결됨
# onclick="GwMainMenu.fn.goPage(..., 'https://hubsso.spc.co.kr:9443/dnsagent/gate/legacySSOGate.jsp?nurl=https://sis.spc.co.kr', ...)"
# 드롭다운 클릭 대신 SSO 게이트 URL 직접 접속 (더 안정적)
_SIS_SSO_GATE = (
    "https://hubsso.spc.co.kr:9443/dnsagent/gate/legacySSOGate.jsp"
    "?nurl=https://sis.spc.co.kr"
)

# ── sis.spc.co.kr 메뉴 페이지 버튼 (DOM 확인 완료) ────────────────
# 각 버튼은 onclick 으로 새 팝업 창을 열고 form POST 로 SSO 로그인
SEL_MENU_OLAP          = "img[onclick*='olapFunc'], img[src*='btn_olap']"
SEL_MENU_LOG_REPORT    = "img[onclick*='logFunc'],  img[src*='btn_log.png']"
SEL_MENU_VISUAL_REPORT = "img[onclick*='vaFunc'],   img[src*='btn_visualreports']"

# VISUAL REPORT 로그인 페이지 하단 "OLAP 계정으로 로그인" 버튼 (표시 시 클릭)
SEL_VR_OLAP_LOGIN = (
    "a:has-text('OLAP 계정으로 로그인'), button:has-text('OLAP 계정으로 로그인')"
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
    """SSO 게이트를 통해 정보화시스템(sis.spc.co.kr)으로 직접 이동."""
    log.info("[STEP] 정보화시스템 접속 (SSO 게이트)")
    try:
        page.goto(_SIS_SSO_GATE, wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_load_state("networkidle", timeout=20_000)
        log.info(f"  정보화시스템 URL: {page.url[:80]}")
        session.snapshot(page, "hub_04_sis_menu")
    except Exception as exc:
        session.snapshot(page, "hub_04_sis_fail")
        raise RuntimeError(f"정보화시스템 접속 실패: {exc}") from exc


def _open_via_menu(
    page: Page,
    link_selector: str,
    label: str,
    config,
    session: BrowserSession,
) -> Page:
    """sis.spc.co.kr 메뉴에서 버튼을 클릭하고 열린 팝업 Page 를 반환합니다.

    각 버튼은 window.open() + form POST 로 새 팝업 창을 엽니다.
    Playwright expect_page() 로 팝업을 캡처하고 로딩 완료를 기다립니다.
    """
    _open_info_system_menu(page, session)

    log.info(f"  {label} 버튼 클릭")
    slug = label.lower().replace(" ", "_")
    try:
        link = page.locator(link_selector).first
        link.wait_for(state="visible", timeout=10_000)

        # window.open() 팝업을 expect_page 로 캡처
        with page.context.expect_page(timeout=15_000) as popup_info:
            link.click()

        popup = popup_info.value
        # 팝업은 window.open("") 으로 먼저 about:blank 로 열리고
        # form.submit() 으로 SSO URL 로 이동함 → about:blank 에서 떠날 때까지 대기
        try:
            popup.wait_for_url(lambda url: url not in ("about:blank", ""), timeout=30_000)
        except PwTimeout:
            pass  # 이미 이동했거나 빠르게 진행된 경우
        popup.wait_for_load_state("load", timeout=60_000)
        log.info(f"  {label} 팝업 열림: {popup.url[:80]}")
        session.snapshot(popup, f"hub_05_{slug}_opened")
        return popup

    except PwTimeout as exc:
        session.snapshot(page, f"hub_05_{slug}_fail")
        raise RuntimeError(
            f"{label} 팝업 열기 실패 (셀렉터 또는 팝업 타임아웃 확인 필요): {exc}"
        ) from exc
