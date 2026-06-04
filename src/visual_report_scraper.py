"""VISUAL REPORT 자동화 — 해피앱 MAU 당월 추출.

URL: https://va.spc.co.kr/SASReportViewer/
경로: 리포트 찾아보기 > 클라우드 > 프로모션 > 해피앱 DAU 리포트 > 열기
탭: 해피앱 접속현황
필터: 연도 / 로그인월(YYYY/MM) / 로그인일자
추출: MAU(당월) KPI 타일 canvas hover → 툴팁 "MAU(당월):N,NNN,NNN" 파싱

# ── DOM 확인 완료 ────────────────────────────────────────────────────
# 프로파일 선택 없음 — OLAP SSO 로그인 후 바로 리포트 뷰어 로드
# 리포트 open: 목록 아이템 Playwright click → #__CntSelDlg0-commitBtn force click
# 탭 로드: 약 30초, 전체 KPI 렌더링: 약 60초 추가 소요
# 월 필터: #__select1 (SAP UI5 Select), title="로그인월: YYYY/MM"
# MAU 값: canvas[id*="birdvisualization"] hover → [role="tooltip"] 텍스트
#          형식: "MAU(당월):N,NNN,NNN"
# ────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path

from playwright.sync_api import Page, Frame, TimeoutError as PwTimeout

from .browser import BrowserSession
from .config import Config
from .olap_scraper import DataNotReadyError
from . import hub_login

log = logging.getLogger(__name__)

# ── 셀렉터 상수 (DOM 확인 완료) ────────────────────────────────────────

# 리포트 찾아보기 → 폴더 탐색
SEL_VR_REPORT_BROWSE = "button:has-text('리포트 찾아보기'), a:has-text('리포트 찾아보기')"
SEL_VR_CLOUD_FOLDER  = "li:has-text('클라우드'), span:has-text('클라우드'), a:has-text('클라우드')"
SEL_VR_PROMO_FOLDER  = "li:has-text('프로모션'), span:has-text('프로모션'), a:has-text('프로모션')"

# 보고서 이름 (해피앱 GA 리포트 → 해피앱 DAU 리포트로 변경됨)
VR_REPORT_NAME = "해피앱 DAU 리포트"

# 열기 다이얼로그 확인 버튼
SEL_VR_COMMIT_BTN = "#__CntSelDlg0-commitBtn"

# 탭: 해피앱 접속현황
VR_TAB_NAME = "해피앱 접속현황"

# 월 필터 select (SAP UI5 Select, id는 세션마다 동일)
SEL_VR_MONTH_SELECT  = "#__select1"

# MAU KPI 타일 canvas (aria-label 기반)
SEL_VR_MAU_CANVAS    = "[aria-label*='MAU(당월)'] canvas"

# 툴팁 (hover 후 표시)
SEL_VR_TOOLTIP       = "[role='tooltip']"

# 리포트 로딩 최대 대기 (탭 등장까지 최대 5분)
VR_TAB_WAIT_MAX_S    = 300
VR_KPI_RENDER_WAIT_S = 90   # 탭 클릭 후 KPI 렌더링 대기


def scrape_mau(config: Config, year: int, month: int) -> int:
    """VISUAL REPORT 에서 해당 연월의 해피앱 MAU 를 반환합니다.

    반환값: 정수 (예: 2_693_419)
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
        page = hub_login.navigate_to_visual_report(page, config, session)

        _open_report(page, session)
        dau_frame = _wait_for_tab(page, session)
        _set_month_filter(dau_frame, year, month, session)
        return _extract_mau(dau_frame, year, month, session)


# ── 내부 함수 ─────────────────────────────────────────────────────────

def _find_in_page_or_frames(page: Page, selector: str, timeout: int = 2_000):
    """메인 페이지 또는 모든 하위 iFrame 에서 selector 를 찾아 Locator 를 반환합니다."""
    el = page.locator(selector)
    try:
        el.first.wait_for(state="visible", timeout=timeout)
        return el.first
    except PwTimeout:
        pass
    for frame in page.frames[1:]:
        try:
            fel = frame.locator(selector)
            fel.first.wait_for(state="visible", timeout=timeout)
            return fel.first
        except Exception:
            pass
    return page.locator(selector).first


def _open_report(page: Page, session: BrowserSession) -> None:
    """리포트 찾아보기 → 클라우드 → 프로모션 → 해피앱 DAU 리포트 → 열기."""
    log.info(f"[STEP] {VR_REPORT_NAME} 열기")

    for label, sel in [
        ("리포트 찾아보기", SEL_VR_REPORT_BROWSE),
        ("클라우드",        SEL_VR_CLOUD_FOLDER),
        ("프로모션",        SEL_VR_PROMO_FOLDER),
    ]:
        try:
            el = _find_in_page_or_frames(page, sel)
            el.wait_for(state="visible", timeout=10_000)
            el.click()
            time.sleep(2)
            log.info(f"  클릭: {label}")
        except PwTimeout:
            session.snapshot(page, f"vr_nav_fail_{label}")
            raise RuntimeError(f"'{label}' 버튼을 찾을 수 없습니다.")

    # 보고서 선택 — Playwright native click (exact text match across frames)
    selected = False
    for frame in page.frames:
        try:
            items = frame.get_by_text(VR_REPORT_NAME, exact=True)
            if items.count() > 0:
                items.first.click(timeout=5_000)
                selected = True
                log.info(f"  선택: {VR_REPORT_NAME}")
                break
        except Exception:
            pass
    if not selected:
        session.snapshot(page, "vr_report_select_fail")
        raise RuntimeError(f"'{VR_REPORT_NAME}' 항목을 찾을 수 없습니다.")

    time.sleep(1)

    # 열기 버튼 — Primary button (SAP native click 필수)
    opened = False
    for frame in page.frames:
        try:
            btn = frame.locator(SEL_VR_COMMIT_BTN)
            if btn.count() > 0:
                btn.click(timeout=5_000, force=True)
                opened = True
                log.info("  열기 클릭")
                break
        except Exception:
            pass
    if not opened:
        session.snapshot(page, "vr_open_fail")
        raise RuntimeError("열기 버튼을 찾을 수 없습니다.")

    session.snapshot(page, "vr_01_report_opening")


def _wait_for_tab(page: Page, session: BrowserSession) -> Frame:
    """해피앱 접속현황 탭이 나타날 때까지 대기하고 클릭 후 Frame 을 반환합니다."""
    log.info(f"[STEP] '{VR_TAB_NAME}' 탭 대기 (최대 {VR_TAB_WAIT_MAX_S}초)")
    steps = VR_TAB_WAIT_MAX_S // 15
    dau_frame = None

    for i in range(steps):
        time.sleep(15)
        for frame in page.frames:
            try:
                tab = frame.get_by_text(VR_TAB_NAME, exact=True)
                if tab.count() > 0:
                    tab.first.click(timeout=5_000)
                    dau_frame = frame
                    log.info(f"  탭 클릭 ({(i+1)*15}s 경과)")
                    break
            except Exception:
                pass
        if dau_frame:
            break

    if dau_frame is None:
        session.snapshot(page, "vr_tab_timeout")
        raise DataNotReadyError(
            f"'{VR_TAB_NAME}' 탭이 {VR_TAB_WAIT_MAX_S}초 이내에 나타나지 않았습니다. "
            "VISUAL REPORT 데이터 미준비 또는 로딩 지연."
        )

    log.info(f"  KPI 렌더링 대기 ({VR_KPI_RENDER_WAIT_S}초)...")
    time.sleep(VR_KPI_RENDER_WAIT_S)
    session.snapshot(page, "vr_02_tab_loaded")
    return dau_frame


def _set_month_filter(frame: Frame, year: int, month: int, session: BrowserSession) -> None:
    """로그인월 필터를 YYYY/MM 으로 설정합니다 (이미 맞으면 건너뜀)."""
    target = f"{year}/{month:02d}"
    log.info(f"[STEP] 로그인월 필터 설정: {target}")

    try:
        select = frame.locator(SEL_VR_MONTH_SELECT)
        current = select.get_attribute("title") or ""
        if target in current:
            log.info(f"  이미 {target} 설정됨 — 건너뜀")
            return

        # SAP UI5 Select: click to open dropdown, then click target option
        select.click(timeout=5_000)
        time.sleep(1)
        option = frame.get_by_text(target, exact=True)
        option.first.click(timeout=5_000)
        time.sleep(3)
        log.info(f"  {target} 선택 완료")
        session.snapshot(frame, "vr_03_month_set")  # type: ignore[arg-type]
    except PwTimeout:
        log.warning(f"  월 필터 변경 실패 — 현재 값으로 진행")


def _extract_mau(frame: Frame, year: int, month: int, session: BrowserSession) -> int:
    """MAU(당월) KPI canvas에 hover → 툴팁 텍스트에서 정수 추출.

    툴팁 형식: "MAU(당월):N,NNN,NNN"
    """
    log.info("[STEP] MAU(당월) 값 추출")
    target_month = f"{year}/{month:02d}"

    # canvas hover
    try:
        canvas = frame.locator(SEL_VR_MAU_CANVAS).first
        canvas.wait_for(state="attached", timeout=30_000)
        canvas.hover()
        time.sleep(2)
    except PwTimeout:
        session.snapshot(frame, "vr_mau_canvas_fail")  # type: ignore[arg-type]
        raise DataNotReadyError(
            "MAU(당월) canvas를 찾을 수 없습니다. "
            "KPI 타일이 아직 렌더링되지 않았거나 필터 설정을 확인하세요."
        )

    # 툴팁 읽기 — JS evaluate (inner_text 은 Canvas overlay에서 빈값 반환)
    time.sleep(1)  # hover 후 툴팁 렌더링 대기
    tip_text = frame.evaluate("""() => {
        var tips = document.querySelectorAll(
            '[role="tooltip"], [class*="tooltip"], [class*="Tooltip"]');
        return Array.from(tips)
            .map(e => e.textContent.trim())
            .filter(t => t.length > 0)
            .join(' | ');
    }""")
    log.debug(f"  툴팁: {tip_text!r}")

    # 파싱: "MAU(당월):2,693,419"
    m = re.search(r"MAU\(당월\)\s*:\s*([\d,]+)", tip_text)
    if m:
        val = int(m.group(1).replace(",", ""))
        log.info(f"  MAU(당월) [{target_month}]: {val:,}")
        session.snapshot(frame, "vr_04_mau_extracted")  # type: ignore[arg-type]
        return val

    # 숫자만 있는 경우 fallback
    nums = re.findall(r"[\d,]+", tip_text)
    for raw in nums:
        v = int(raw.replace(",", ""))
        if v > 100_000:  # MAU는 최소 10만 이상
            log.info(f"  MAU(당월) fallback [{target_month}]: {v:,}")
            return v

    session.snapshot(frame, "vr_04_mau_parse_fail")  # type: ignore[arg-type]
    raise DataNotReadyError(
        f"MAU(당월) 값을 파싱할 수 없습니다 (툴팁: {tip_text!r}). "
        f"필터가 {target_month}로 설정되었는지 확인하세요."
    )
