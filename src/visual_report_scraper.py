"""VISUAL REPORT 자동화 — 해피앱 MAU 당월 Excel 다운로드.

URL: https://va.spc.co.kr/SASReportViewer/
경로: 리포트 찾아보기 > 클라우드 > 프로모션 > 해피앱 GA 리포트 > 열기
작업: 필터 지우기 → MAU 당월 내보내기 → Excel 다운로드

주의: 리포트 로딩에 최대 5분 소요됩니다 (timeout=360s).

# ── 셀렉터 튜닝 가이드 ───────────────────────────────────────────────
# HEADLESS=0, DRY_RUN=1 로 실행하면 브라우저가 열립니다.
# logs/debug/ 폴더의 스크린샷(.png)·HTML(.html)로 실제 DOM 확인 후
# 아래 SEL_VR_* 상수를 수정하세요.
# ────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import logging
import time
from pathlib import Path

from playwright.sync_api import Page, TimeoutError as PwTimeout

from .browser import BrowserSession
from .config import Config

log = logging.getLogger(__name__)

# ── 셀렉터 상수 (실제 DOM 확인 후 수정) ──────────────────────────────
# 로그인
SEL_VR_LOGIN_ID  = "input[name='userId'], input[id='userId'], input[name='id'], input[type='text']"
SEL_VR_LOGIN_PW  = "input[name='password'], input[id='password'], input[type='password']"
SEL_VR_LOGIN_BTN = "button[type='submit'], input[type='submit'], button:has-text('로그인')"

# 리포트 찾아보기 → 폴더 탐색
SEL_VR_REPORT_BROWSE  = "button:has-text('리포트 찾아보기'), a:has-text('리포트 찾아보기')"
SEL_VR_CLOUD_FOLDER   = "li:has-text('클라우드'), span:has-text('클라우드'), a:has-text('클라우드')"
SEL_VR_PROMO_FOLDER   = "li:has-text('프로모션'), span:has-text('프로모션'), a:has-text('프로모션')"
SEL_VR_HAPPYAPP_ITEM  = (
    "li:has-text('해피앱 GA 리포트'), span:has-text('해피앱 GA 리포트'), "
    "a:has-text('해피앱 GA 리포트')"
)
SEL_VR_OPEN_BTN = "button:has-text('열기'), a:has-text('열기')"

# 필터 조작
SEL_VR_FILTER_SETTINGS = (
    "button:has-text('필터 설정'), button[aria-label*='필터'], "
    "span:has-text('필터'), a:has-text('필터 설정')"
)
SEL_VR_FILTER_CLEAR = (
    "button:has-text('필터 지우기'), button:has-text('지우기'), "
    "a:has-text('필터 지우기'), span:has-text('필터 지우기')"
)

# MAU 당월 위젯 → 더보기(···) → 데이터 내보내기
SEL_VR_MAU_WIDGET = (
    "div:has-text('MAU 당월'), section:has-text('MAU 당월'), "
    "span:has-text('MAU 당월')"
)
SEL_VR_MAU_MORE_BTN = (
    "[aria-label*='옵션'], button[title*='더보기'], button[aria-label*='more'], "
    "button:has-text('⋯'), button:has-text('···')"
)
SEL_VR_EXPORT_DATA    = "li:has-text('데이터 내보내기'), a:has-text('데이터 내보내기')"
SEL_VR_EXPORT_CONFIRM = "button:has-text('확인'), button:has-text('OK')"

# 리포트 로딩 최대 대기 시간 (5분 관측값 기준)
REPORT_LOAD_TIMEOUT = 360_000  # ms


def scrape_mau_excel(config: Config, year: int, month: int) -> Path:
    """VISUAL REPORT 에서 MAU 당월 Excel 파일을 다운로드하고 경로를 반환합니다."""
    dl_dir = config.runtime.download_dir
    debug_dir = config.runtime.logs_dir / "debug"

    with BrowserSession(
        headless=config.runtime.headless,
        download_dir=dl_dir,
        debug_dir=debug_dir,
    ) as session:
        page = session.new_page()
        _login(page, config.visual_report.base_url, config.olap.user_id, config.olap.password, session)
        _open_happyapp_report(page, session)
        _clear_date_filter(page, session)
        return _export_mau_excel(page, dl_dir, session)


# ── 내부 함수 ─────────────────────────────────────────────────────────

def _login(page: Page, base_url: str, user_id: str, password: str, session: BrowserSession) -> None:
    log.info("[STEP] VISUAL REPORT 로그인 페이지 접속")
    page.goto(base_url, wait_until="domcontentloaded", timeout=30_000)
    session.snapshot(page, "vr_01_login_page")

    try:
        id_input = page.locator(SEL_VR_LOGIN_ID).first
        id_input.wait_for(state="visible", timeout=10_000)
        id_input.fill(user_id)
        page.locator(SEL_VR_LOGIN_PW).first.fill(password)
        page.locator(SEL_VR_LOGIN_BTN).first.click(timeout=5_000)
        page.wait_for_load_state("networkidle", timeout=30_000)
        log.info("  로그인 완료")
        session.snapshot(page, "vr_02_after_login")
    except PwTimeout as exc:
        session.snapshot(page, "vr_02_login_timeout")
        raise RuntimeError(f"VISUAL REPORT 로그인 실패 (셀렉터 확인 필요): {exc}") from exc


def _open_happyapp_report(page: Page, session: BrowserSession) -> None:
    log.info("[STEP] 해피앱 GA 리포트 열기")

    # 리포트 찾아보기 클릭
    _click_step(page, session, SEL_VR_REPORT_BROWSE, "리포트 찾아보기", "vr_03a_browse")
    time.sleep(1.0)

    # 클라우드 폴더 클릭
    _click_step(page, session, SEL_VR_CLOUD_FOLDER, "클라우드", "vr_03b_cloud")
    time.sleep(0.8)

    # 프로모션 폴더 클릭
    _click_step(page, session, SEL_VR_PROMO_FOLDER, "프로모션", "vr_03c_promo")
    time.sleep(0.8)

    # 해피앱 GA 리포트 항목 클릭
    _click_step(page, session, SEL_VR_HAPPYAPP_ITEM, "해피앱 GA 리포트", "vr_03d_item")
    time.sleep(0.5)

    # 열기 버튼 클릭
    _click_step(page, session, SEL_VR_OPEN_BTN, "열기", "vr_03e_open")

    # 리포트 로딩 대기 (최대 6분)
    log.info(f"  리포트 로딩 대기 중 (최대 {REPORT_LOAD_TIMEOUT // 1000}초)…")
    try:
        page.wait_for_load_state("networkidle", timeout=REPORT_LOAD_TIMEOUT)
    except PwTimeout:
        pass  # networkidle 미도달이어도 진행 — 리포트가 iframe 안에서 자체 로딩할 수 있음
    session.snapshot(page, "vr_04_report_loaded")
    log.info("  리포트 열기 완료")


def _clear_date_filter(page: Page, session: BrowserSession) -> None:
    log.info("[STEP] 로그인일자 필터 삭제")
    try:
        btn = page.locator(SEL_VR_FILTER_SETTINGS).first
        btn.wait_for(state="visible", timeout=15_000)
        btn.click()
        time.sleep(0.8)
        page.locator(SEL_VR_FILTER_CLEAR).first.click(timeout=8_000)
        time.sleep(1.0)
        log.info("  필터 지우기 완료")
        session.snapshot(page, "vr_05_filter_cleared")
    except PwTimeout:
        session.snapshot(page, "vr_05_filter_fail")
        raise RuntimeError(
            "필터 지우기 실패 (SEL_VR_FILTER_SETTINGS / SEL_VR_FILTER_CLEAR 확인 필요)"
        )


def _export_mau_excel(page: Page, dl_dir: Path, session: BrowserSession) -> Path:
    log.info("[STEP] MAU 당월 데이터 내보내기")

    # MAU 당월 위젯의 더보기(···) 버튼 탐색
    # 위젯을 먼저 찾아 그 안의 ··· 버튼을 클릭하는 방식
    try:
        mau_widget = page.locator(SEL_VR_MAU_WIDGET).first
        mau_widget.wait_for(state="visible", timeout=15_000)
        mau_widget.hover()  # ··· 버튼이 hover 시 표시되는 경우 대비
        time.sleep(0.5)

        more_btn = mau_widget.locator(SEL_VR_MAU_MORE_BTN).first
        try:
            more_btn.wait_for(state="visible", timeout=5_000)
        except PwTimeout:
            # hover 후에도 없으면 page 전체에서 탐색
            more_btn = page.locator(SEL_VR_MAU_MORE_BTN).first
            more_btn.wait_for(state="visible", timeout=5_000)

        more_btn.click(timeout=5_000)
        time.sleep(0.5)
        log.info("  ··· 버튼 클릭")
        session.snapshot(page, "vr_06a_more_menu")

    except PwTimeout:
        session.snapshot(page, "vr_06a_more_fail")
        raise RuntimeError(
            "MAU 당월 ··· 버튼을 찾을 수 없습니다. "
            "SEL_VR_MAU_WIDGET / SEL_VR_MAU_MORE_BTN 셀렉터를 확인하세요."
        )

    # 데이터 내보내기 클릭
    try:
        page.locator(SEL_VR_EXPORT_DATA).first.click(timeout=8_000)
        time.sleep(0.5)
        session.snapshot(page, "vr_06b_export_dialog")
    except PwTimeout:
        session.snapshot(page, "vr_06b_export_fail")
        raise RuntimeError("데이터 내보내기 메뉴를 찾을 수 없습니다 (SEL_VR_EXPORT_DATA 확인 필요)")

    # 확인 버튼 클릭 + 파일 다운로드 대기
    try:
        with page.expect_download(timeout=120_000) as dl_info:
            page.locator(SEL_VR_EXPORT_CONFIRM).first.click(timeout=8_000)
        dl = dl_info.value
        save_path = dl_dir / dl.suggested_filename
        dl.save_as(save_path)
        log.info(f"  MAU Excel 다운로드 완료: {save_path.name}")
        session.snapshot(page, "vr_07_downloaded")
        return save_path

    except PwTimeout as exc:
        session.snapshot(page, "vr_07_download_fail")
        raise RuntimeError(
            f"MAU Excel 다운로드 실패 (확인 버튼 또는 파일 대기 시간 초과): {exc}\n"
            "SEL_VR_EXPORT_CONFIRM 셀렉터 및 다운로드 경로를 확인하세요."
        ) from exc


def _click_step(
    page: Page,
    session: BrowserSession,
    selector: str,
    label: str,
    snapshot_label: str,
) -> None:
    try:
        el = page.locator(selector).first
        el.wait_for(state="visible", timeout=10_000)
        el.click()
        session.snapshot(page, snapshot_label)
    except PwTimeout:
        session.snapshot(page, f"{snapshot_label}_fail")
        raise RuntimeError(
            f"'{label}' 요소를 찾을 수 없습니다. "
            f"logs/debug/{snapshot_label}_fail.png 를 확인해 셀렉터를 수정하세요."
        )
