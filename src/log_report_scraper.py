"""LOG REPORT 자동화 — 해피앱 월 로그인 회원수(순 로그인 회원수) 추출.

URL: https://hplog.spc.co.kr:8000/datastory/home
경로: 해피앱 > 종합 > 종합추이(MONTHLY 설정) > 누적추이 > 테이블 월별 보기
추출값: 순 로그인 회원수 (해당 연월 행)

# ── DOM 확인 완료 ────────────────────────────────────────────────────
# 프로파일 목록 페이지: td.profile[title='해피앱'] 클릭 → window.open → 새 탭
# 트리 탐색: div[title='...'] 또는 p:has-text('...') (div.tree_node 구조)
# 월별 전환: 종합추이에서 div.period-select select 를 MONTHLY 로 설정 후 누적추이로 이동
# 테이블 월별: #period_select select_option("MONTHLY")
# 데이터 테이블: table 내 thead가 '순 로그인 회원수' 포함, 행 형식 "YYYY.MM월 [start~end]"
# ────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path

from playwright.sync_api import Page, TimeoutError as PwTimeout

from .browser import BrowserSession
from .config import Config
from .olap_scraper import DataNotReadyError
from . import hub_login

log = logging.getLogger(__name__)

# ── 셀렉터 상수 ────────────────────────────────────────────────────────

# 프로파일 목록 화면의 해피앱 행
SEL_LR_PROFILE_ROW = "td.profile[title='해피앱'], td[title='해피앱']"

# 좌측 트리 메뉴 (div.tree_node 구조)
SEL_LR_NAV_HAPPYAPP  = "a:has-text('해피앱'), span:has-text('해피앱'), li:has-text('해피앱')"
SEL_LR_NAV_SUMMARY   = "a:has-text('종합'), span:has-text('종합'), li:has-text('종합')"
SEL_LR_NAV_SUMTREND  = "div[title='종합추이'], p:has-text('종합추이')"
SEL_LR_NAV_CUMTREND  = "div[title='누적추이'], p:has-text('누적추이')"

# 테이블 #period_select (tr 내부, inline-block 으로 표시됨)
SEL_LR_PERIOD_SELECT = "#period_select"

# 로그인 회원수 컬럼 인덱스 (0-based, 빈 열 포함 후 5열 구조)
# 열: [날짜, (empty), 순방문자수, 순로그인회원수, 일합계방문자수]
LR_LOGIN_COL_IDX = 3


def scrape_login_count(config: Config, year: int, month: int) -> int:
    """LOG REPORT 에서 해당 연월의 순 로그인 회원수를 반환합니다.

    데이터 미준비 시 DataNotReadyError 를 raise 합니다.
    """
    dl_dir    = config.runtime.download_dir
    debug_dir = config.runtime.logs_dir / "debug"

    with BrowserSession(
        headless=config.runtime.headless,
        download_dir=dl_dir,
        debug_dir=debug_dir,
    ) as session:
        page = session.new_page()
        hub_login.login_to_hub(page, config, session)
        page = hub_login.navigate_to_log_report(page, config, session)
        page = _select_profile(page, session)
        _navigate_to_trend(page, session)
        _switch_to_monthly(page, session)
        return _extract_login_count(page, year, month, session)


# ── 내부 함수 ─────────────────────────────────────────────────────────

def _select_profile(page: Page, session: BrowserSession) -> Page:
    """프로파일 목록 화면이면 '해피앱' 클릭 → 새 탭(report?profileId=)을 열고 반환합니다.

    jQuery handler: $(document).on('click', '#report td.profile', function(t) {
        window.open("report?profileId=" + $(t.target).attr("profileid"), "_blank")
    })
    """
    try:
        profile = page.locator(SEL_LR_PROFILE_ROW).first
        profile.wait_for(state="visible", timeout=5_000)
    except PwTimeout:
        log.info("  프로파일 선택 화면 없음 — 현재 페이지 사용")
        return page

    log.info("  프로파일 선택 화면 감지 — 새 탭으로 열기")
    with page.context.expect_page(timeout=15_000) as popup_info:
        profile.click()

    report_page = popup_info.value
    report_page.wait_for_load_state("networkidle", timeout=30_000)
    log.info(f"  대시보드 탭 열림: {report_page.url[:70]}")
    session.snapshot(report_page, "lr_02b_dashboard")
    return report_page


def _navigate_to_trend(page: Page, session: BrowserSession) -> None:
    """해피앱 > 종합 > 종합추이 로 이동 후 페이지 수준 월별 모드를 설정,
    이후 누적추이 로 이동합니다.

    종합추이를 먼저 방문해야 div.period-select가 활성화되고
    누적추이에서 #period_select 월별 전환이 정상 동작합니다.
    """
    log.info("[STEP] 해피앱 > 종합 > 종합추이 → (MONTHLY) → 누적추이 탐색")

    # ① 해피앱 > 종합 > 종합추이
    for label, sel in [
        ("해피앱",   SEL_LR_NAV_HAPPYAPP),
        ("종합",     SEL_LR_NAV_SUMMARY),
        ("종합추이", SEL_LR_NAV_SUMTREND),
    ]:
        try:
            node = page.locator(sel).first
            node.wait_for(state="visible", timeout=10_000)
            node.click()
            time.sleep(0.8)
            log.info(f"  클릭: {label}")
        except PwTimeout:
            session.snapshot(page, f"lr_nav_fail_{label}")
            raise RuntimeError(f"'{label}' 메뉴를 찾을 수 없습니다.")
    page.wait_for_load_state("networkidle", timeout=20_000)
    session.snapshot(page, "lr_03a_summary_trend")

    # ② 종합추이에서 페이지 수준 MONTHLY 설정
    try:
        page.evaluate(
            "() => { var d = document.querySelector('div.period-select');"
            " if (d) d.style.display = 'block'; }"
        )
        page.locator("div.period-select select").select_option("MONTHLY", timeout=5_000)
        time.sleep(2)
        page.wait_for_load_state("networkidle", timeout=15_000)
        log.info("  페이지 수준 MONTHLY 설정 완료")
    except PwTimeout:
        log.warning("  div.period-select select 를 찾을 수 없음 — 건너뜀")

    # ③ 누적추이로 이동
    try:
        node = page.locator(SEL_LR_NAV_CUMTREND).first
        node.wait_for(state="visible", timeout=10_000)
        node.click()
        time.sleep(1.5)
        log.info("  클릭: 누적추이")
    except PwTimeout:
        session.snapshot(page, "lr_nav_fail_누적추이")
        raise RuntimeError("'누적추이' 메뉴를 찾을 수 없습니다.")
    page.wait_for_load_state("networkidle", timeout=20_000)
    session.snapshot(page, "lr_03b_cumulate_trend")


def _switch_to_monthly(page: Page, session: BrowserSession) -> None:
    """누적추이 데이터 테이블의 #period_select를 MONTHLY 로 전환합니다."""
    log.info("[STEP] 누적추이 테이블 월별 보기 전환")
    try:
        page.locator(SEL_LR_PERIOD_SELECT).select_option("MONTHLY", timeout=8_000)
        time.sleep(3)
        page.wait_for_load_state("networkidle", timeout=20_000)
        log.info("  테이블 월별 전환 완료")
        session.snapshot(page, "lr_04_monthly_table")
    except PwTimeout:
        session.snapshot(page, "lr_04_monthly_fail")
        raise RuntimeError(
            "누적추이 테이블 #period_select 를 찾을 수 없습니다. "
            "SEL_LR_PERIOD_SELECT 확인 필요."
        )


def _extract_login_count(page: Page, year: int, month: int, session: BrowserSession) -> int:
    """누적추이 테이블 월별 보기에서 순 로그인 회원수를 추출합니다.

    테이블 구조 (확인 완료):
      thead: [주별/월별 토글, 순방문자수, 순로그인회원수, 일합계방문자수]
      tbody row: [YYYY.MM월 [start~end], (empty), 순방문자수, 순로그인회원수, 일합계방문자수]
    → LR_LOGIN_COL_IDX = 3 (0-based)
    """
    log.info(f"[STEP] {year}-{month:02d} 순 로그인 회원수 추출")
    session.snapshot(page, "lr_05_before_extract")

    target_prefix = f"{year}.{month:02d}월"

    rows = page.evaluate("""() => {
        var tables = Array.from(document.querySelectorAll('table'));
        var target = tables.find(function(t){
            return t.innerText && t.innerText.includes('순 로그인 회원수');
        });
        if (!target) return [];
        return Array.from(target.querySelectorAll('tr')).map(function(r){
            return Array.from(r.querySelectorAll('th,td')).map(function(c){
                return c.innerText.trim();
            });
        });
    }""")

    if not rows:
        session.snapshot(page, "lr_05_table_missing")
        raise DataNotReadyError(
            f"순 로그인 회원수 테이블을 찾을 수 없습니다 ({year}-{month:02d}). "
            "데이터 미준비 또는 페이지 구조 변경 가능성."
        )

    for row in rows:
        if not row or not row[0].startswith(target_prefix):
            continue
        # row 형식: [YYYY.MM월 [...], '', 순방문자수, 순로그인회원수, 일합계방문자수]
        if len(row) > LR_LOGIN_COL_IDX:
            raw = row[LR_LOGIN_COL_IDX].replace(",", "").strip()
            if raw.isdigit():
                val = int(raw)
                log.info(f"  순 로그인 회원수 ({target_prefix}): {val:,}")
                return val

    # 대상 월 행 없음 → 아직 데이터 미생성
    session.snapshot(page, "lr_05_month_missing")
    raise DataNotReadyError(
        f"{target_prefix} 행을 찾을 수 없습니다. "
        "데이터가 아직 LOG REPORT 에 반영되지 않았거나 월별 보기 전환이 실패했습니다."
    )
