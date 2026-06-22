"""Playwright 브라우저 세션 헬퍼 (생성/종료/스크린샷/덤프)."""

from __future__ import annotations

import logging
import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional

from playwright.sync_api import Browser, BrowserContext, Page, Playwright, sync_playwright

# Node.js v22 crashes with ACCESS_VIOLATION when JavaScript module paths
# contain non-ASCII characters (e.g. Korean folder names). If playwright's
# package directory is on a non-ASCII path, copy it to a temp ASCII location.
def _ensure_ascii_cli(cli_path: str) -> str:
    if cli_path.isascii():
        return cli_path
    import hashlib
    import shutil
    import tempfile
    pkg_dir = os.path.dirname(cli_path)
    digest = hashlib.md5(Path(cli_path).read_bytes()).hexdigest()[:8]
    dest = Path(tempfile.gettempdir()) / f"pw_pkg_{digest}"
    if not (dest / "cli.js").exists():
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(pkg_dir, dest)
    return str(dest / "cli.js")

import playwright._impl._driver as _pw_driver
_orig_compute = _pw_driver.compute_driver_executable
def _patched_compute():
    node, cli = _orig_compute()
    return node, _ensure_ascii_cli(cli)
_pw_driver.compute_driver_executable = _patched_compute

log = logging.getLogger(__name__)
KST = timezone(timedelta(hours=9))


class BrowserSession:
    """Playwright 기반 브라우저 세션 — context manager."""

    def __init__(
        self,
        headless: bool = True,
        download_dir: Optional[Path] = None,
        debug_dir: Optional[Path] = None,
    ):
        self.headless = headless
        self.download_dir = download_dir
        self.debug_dir = debug_dir or Path("./logs/debug")
        self.debug_dir.mkdir(parents=True, exist_ok=True)
        if download_dir:
            download_dir.mkdir(parents=True, exist_ok=True)

        self._pw: Optional[Playwright] = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None

    def __enter__(self) -> "BrowserSession":
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(
            headless=self.headless,
            args=["--disable-blink-features=AutomationControlled"],
        )
        ctx_opts: dict = dict(
            locale="ko-KR",
            timezone_id="Asia/Seoul",
            ignore_https_errors=True,  # 내부망 자체서명 인증서 대응
        )
        if self.download_dir:
            ctx_opts["accept_downloads"] = True
        self._context = self._browser.new_context(**ctx_opts)
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        try:
            if self._context is not None:
                self._context.close()
            if self._browser is not None:
                self._browser.close()
        finally:
            if self._pw is not None:
                self._pw.stop()

    def new_page(self) -> Page:
        assert self._context is not None
        page = self._context.new_page()
        page.on("console", lambda msg: log.debug(f"  [CONSOLE:{msg.type}] {msg.text}"))
        page.on("pageerror", lambda err: log.error(f"  [PAGE_ERROR] {err}"))
        return page

    def snapshot(self, page: Page, label: str) -> None:
        """단계별 스크린샷 + HTML 덤프 저장 (셀렉터 튜닝·디버그용)."""
        ts = datetime.now(KST).strftime("%Y%m%d_%H%M%S")
        base = self.debug_dir / f"{ts}_{label}"
        try:
            page.screenshot(path=str(base.with_suffix(".png")), full_page=True, timeout=10_000)
        except Exception as exc:
            log.debug(f"  스크린샷 실패: {exc}")
        try:
            base.with_suffix(".html").write_text(page.content(), encoding="utf-8")
        except Exception as exc:
            log.debug(f"  HTML 덤프 실패: {exc}")
        # iFrame 내용 덤프 (날짜 필터·버튼 셀렉터 튜닝용)
        try:
            for frame in page.frames[1:]:  # 첫 번째는 메인 프레임
                if not frame.name:
                    continue
                frame_path = self.debug_dir / f"{ts}_{label}_frame_{frame.name[:30]}.html"
                frame_path.write_text(frame.content(), encoding="utf-8")
        except Exception:
            pass
