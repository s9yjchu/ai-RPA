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
from . import hub_login

log = logging.getLogger(__name__)
KST = timezone(timedelta(hours=9))

# ── 셀렉터 상수 (DOM 확인 완료) ─────────────────────────────────────
# OLAP 리포트 탐색 — 3단계 구조 (DOM 확인 완료)
#   1단계 탭:      a.ui-tabs-anchor              (전사, 정형데이터)
#   2단계 아코디언: h3.ui-accordion-header        (마감리포트, 회원분석, …)
#   3단계 리포트:  span.STPRVmenuItem a           (실제 리포트 링크)
SEL_TREE_NODE = (
    "a.ui-tabs-anchor:has-text('{label}'), "
    "h3.ui-accordion-header:has-text('{label}'), "
    "span.STPRVmenuItem a:has-text('{label}')"
)

# 날짜 필터 — iFrame 내 id 확인 완료 (StoredProcess HBI 리포트별 패턴 상이)
# member_metrics:   sltdate_start / sltdate_end
# channel_metrics:  sltPERIOD_start / sltPERIOD_end
SEL_DATE_START = (
    "input#sltdate_start, input#sltPERIOD_start, "
    "input[id*='date_start'], input[id*='PERIOD_start'], input[id*='Date_start']"
)
SEL_DATE_END = (
    "input#sltdate_end, input#sltPERIOD_end, "
    "input[id*='date_end'], input[id*='PERIOD_end'], input[id*='Date_end']"
)
# jQuery UI datepicker — input에 직접 타이핑. 캘린더 위젯은 불필요.
SEL_CALENDAR_OPEN = "button.ui-datepicker-trigger, img.ui-datepicker-trigger"
SEL_CALENDAR_DAY = "td[data-handler='selectDay'] a:has-text('{day}')"

# 실행 버튼 — iFrame 내 id 확인 완료 (input#btnRun, value="Submit")
SEL_RUN_BTN = "input#btnRun, input[value='Submit'], input[id*='Run'][type='button']"

# Excel 내보내기
# HBI StoredProcess: input#btnExcel (직접 클릭, 서브메뉴 없음)
# SAS WRS: button#citationFileButton (파일 메뉴) → "내보내기..." 클릭
SEL_EXPORT_BTN      = "input#btnExcel, input[value='엑셀'], input[id*='Excel'][type='button']"
SEL_EXPORT_EXCEL    = SEL_EXPORT_BTN  # HBI: 동일 버튼
SEL_WRS_FILE_MENU   = "button#citationFileButton"                # WRS 파일(F) 메뉴 버튼
SEL_WRS_EXPORT_ITEM = "li:has-text('내보내기'), a:has-text('내보내기')"  # 메뉴 항목

# 리포트 이름 경로 (트리 클릭 순서)
REPORT_PATHS = {
    "member_metrics": [
        "정형데이터",
        "회원분석",
        "[HPC] 일별 회원관리지표",
    ],
    "channel_metrics": [
        "정형데이터",
        "회원분석",                              # 채널별 보고서는 회원분석 아코디언 안에 있음
        "[HPC] 채널별 적립, 사용건수 현황",
    ],
    "closing_report": [
        "전사",
        "마감리포트",
        "2. [HPC, POS] HPC 일마감(브랜드)",
    ],
}

# 리포트 콘텐츠 iFrame — aria-hidden='false' 인 활성 탭 패널 내의 iFrame 만 선택
# (비활성 탭에도 frmSASHBI iFrame이 있으므로 전체 매칭 시 잘못된 iFrame 선택 가능)
REPORT_FRAME_SELECTOR: Optional[str] = (
    "div[aria-hidden='false'] iframe[name^='frmSASHBI'], "
    "div[role='tabpanel'][aria-hidden='false'] iframe[name^='frmSASHBI']"
)


def yesterday_kst() -> date:
    return datetime.now(KST).date() - timedelta(days=1)


# ── 로그인 ────────────────────────────────────────────────────────────

def login(page: Page, config, session: BrowserSession) -> Page:
    """SPC Hub SSO 로그인 후 OLAP 으로 이동합니다.

    OLAP 이 새 탭으로 열리는 경우 해당 탭의 Page 를 반환합니다.
    caller 는 반환된 Page 를 이후 작업에 사용해야 합니다.
    """
    hub_login.login_to_hub(page, config, session)
    active = hub_login.navigate_to_olap(page, config, session)
    log.info("  OLAP 진입 완료")
    return active


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

    # 탐색은 항상 메인 페이지에서 (iFrame 내부가 아님)
    target = page

    for i, label in enumerate(path):
        sel = SEL_TREE_NODE.format(label=label)
        try:
            node = target.locator(sel).first
            node.wait_for(state="visible", timeout=10_000)
        except PwTimeout:
            session.snapshot(page, f"tree_fail_{i}_{label[:20]}")
            raise RuntimeError(
                f"트리 노드 '{label}' 를 찾을 수 없습니다. "
                f"logs/debug/ 스크린샷을 확인해 SEL_TREE_NODE 셀렉터를 수정하세요."
            )

        # 마지막 노드가 아닐 때: 다음 노드가 이미 보이면 클릭 생략 (아코디언 중복 토글 방지)
        if i < len(path) - 1:
            next_sel = SEL_TREE_NODE.format(label=path[i + 1])
            try:
                target.locator(next_sel).first.wait_for(state="visible", timeout=1_500)
                log.info(f"  건너뜀 (이미 펼쳐짐): {label}")
                continue
            except PwTimeout:
                pass  # 다음 노드 안 보임 → 클릭해서 펼침

        node.click()
        time.sleep(1.0)  # 탭/아코디언 애니메이션 대기
        log.info(f"  클릭: {label}")

    # 리포트 폼 로딩 대기 — Submit 버튼이 보일 때까지 (networkidle 은 AJAX로 인해 비사용)
    try:
        _get_active_page(page).locator(SEL_RUN_BTN).first.wait_for(state="visible", timeout=15_000)
        log.info("  리포트 폼 로딩 완료")
    except PwTimeout:
        time.sleep(3)
        log.info("  리포트 폼 대기 완료 (타임아웃 후 진행)")
    session.snapshot(page, f"03_report_loaded_{report_key}")


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
            inp.click()           # 포커스
            inp.press("Control+a")
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
        log.info("  조회 버튼 클릭 완료 — 결과 대기 중")

        # StoredProcess HBI: #progressIndicatorWIP 가 사라지면 조회 완료
        try:
            target.locator("#progressIndicatorWIP").wait_for(state="hidden", timeout=120_000)
            log.info("  조회 완료 (progressIndicator 숨김)")
        except PwTimeout:
            time.sleep(3)
            log.info("  조회 완료 (타임아웃 후 진행)")

        session.snapshot(page, "05_after_run")

    except PwTimeout:
        # SAS WRS 리포트 등 Submit 버튼 없이 자동 렌더링되는 경우
        log.info("  조회 버튼 없음 — SAS WRS 자동 로드 리포트로 간주, 대기 중")
        time.sleep(5)  # WRS 리포트 렌더링 대기
        session.snapshot(page, "05_wrs_auto_loaded")


def _verify_report_date(page: Page, target_date: date) -> bool:
    """리포트 본문(iFrame 포함)에 어제 날짜가 포함되어 있는지 확인."""
    date_str = target_date.strftime("%Y-%m-%d")
    alt      = target_date.strftime("%Y/%m/%d")
    yyyymmdd = target_date.strftime("%Y%m%d")

    def _has_date(text: str) -> bool:
        return date_str in text or alt in text or yyyymmdd in text

    # 외부 페이지 확인
    if _has_date(page.content()):
        return True

    # iFrame 내용 확인 (리포트 결과는 frmSASHBI iFrame 안에 있음)
    try:
        for frame in page.frames:
            if "frmSASHBI" in (frame.name or ""):
                if _has_date(frame.content()):
                    return True
    except Exception:
        pass

    # dvData 에 테이블이 있으면 데이터 준비됨으로 간주 (날짜가 숨겨진 경우 대비)
    try:
        active = _get_active_page(page)
        dv = active.locator("#dvData table")
        dv.first.wait_for(state="visible", timeout=2_000)
        cell_count = dv.first.locator("td").count()
        if cell_count > 0:
            log.info(f"  dvData 테이블 존재 ({cell_count}개 셀) → 데이터 준비됨으로 간주")
            return True
    except Exception:
        pass

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

    # HBI 방식: input#btnExcel 직접 클릭
    hbi_btn = target.locator(SEL_EXPORT_BTN).first
    try:
        hbi_btn.wait_for(state="visible", timeout=5_000)
        hbi_found = True
    except PwTimeout:
        hbi_found = False

    if hbi_found:
        try:
            with page.expect_download(timeout=120_000) as dl_info:
                hbi_btn.click(timeout=5_000)
            dl = dl_info.value
            save_path = download_dir / f"{report_key}_{dl.suggested_filename}"
            dl.save_as(save_path)
            log.info(f"  저장 완료: {save_path}")
            session.snapshot(page, f"06_downloaded_{report_key}")
            return save_path
        except Exception as exc:
            session.snapshot(page, f"06_hbi_download_fail_{report_key}")
            raise RuntimeError(f"HBI Excel 다운로드 실패 ({report_key}): {exc}") from exc

    # WRS 방식: cwMenuBarExport() JS 직접 호출 → exportReport.do 이동 → 다운로드
    # (툴바가 embedded 모드에서 숨겨져 있으므로 JS 직접 호출)
    log.info("  HBI 버튼 없음 → SAS WRS cwMenuBarExport() 직접 호출")
    try:
        # page.frames 에서 frmSASHBI* 이름의 활성 프레임 탐색
        wrs_frame = None
        for f in page.frames:
            if "frmSASHBI" in (f.name or "") and "AA00012A" in (f.name or ""):
                wrs_frame = f
                break
        if wrs_frame is None:
            for f in page.frames:
                if "frmSASHBI" in (f.name or ""):
                    wrs_frame = f
                    break

        if wrs_frame is None:
            raise RuntimeError("WRS iFrame을 찾을 수 없습니다.")

        log.info(f"  WRS iFrame: {wrs_frame.name}, URL: {wrs_frame.url[:60]}")

        with page.expect_download(timeout=120_000) as dl_info:
            wrs_frame.evaluate("cwMenuBarExport()")

        dl = dl_info.value
        save_path = download_dir / f"{report_key}_{dl.suggested_filename}"
        dl.save_as(save_path)
        log.info(f"  저장 완료: {save_path}")
        session.snapshot(page, f"06_downloaded_{report_key}")
        return save_path

    except Exception as exc:
        session.snapshot(page, f"06_wrs_download_fail_{report_key}")
        raise RuntimeError(
            f"WRS 내보내기 실패 ({report_key}): {exc}\n"
            "cwMenuBarExport() 호출 또는 exportReport.do 응답을 확인하세요."
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
