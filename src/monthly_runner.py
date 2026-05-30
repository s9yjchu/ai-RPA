"""월별 고객지표 업데이트 — LOG REPORT / VISUAL REPORT.

[미구현 — 추후 완성 예정]

완성에 필요한 정보:
  - LOG REPORT > 해피앱 > 종합 > 누적추이 > 월별 보기: 순 로그인 회원수 추출 방법
  - VISUAL REPORT > 해피앱 DAU 리포트 > 해피앱 접속현황: MAU Excel 파일 구조
"""

from __future__ import annotations

import logging
from datetime import date
from .config import Config
from .state_manager import MonthlyState

log = logging.getLogger(__name__)


def run_monthly(config: Config, year: int, month: int, force: bool = False) -> None:
    state = MonthlyState(config.runtime.state_dir, year, month)

    if state.is_done and not force:
        log.info(f"[SKIP] {year}-{month:02d} 월별 업데이트 이미 완료됨")
        return

    if state.should_give_up() and not force:
        log.error(f"[ABORT] {year}-{month:02d} 최대 재시도 기간 초과 — 수동 확인 필요")
        return

    log.warning(
        "[TODO] 월별 업데이트 미구현.\n"
        "LOG REPORT 및 VISUAL REPORT 데이터 추출 방법 확인 후 구현 예정."
    )
    raise NotImplementedError("월별 업데이트는 아직 구현되지 않았습니다.")
