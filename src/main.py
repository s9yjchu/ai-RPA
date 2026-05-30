"""고객지표 RPA 진입점.

사용법:
  python -m src.main daily              # 어제 날짜 일별 업데이트
  python -m src.main daily --date 2026-05-29   # 특정 날짜 (재실행)
  python -m src.main daily --force      # 이미 완료된 날도 강제 재실행
  python -m src.main monthly            # 당월 월별 업데이트 (추후 구현)
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import date, datetime, timezone, timedelta

from .config import load_config
from .daily_runner import run_daily, yesterday_kst
from .logger import setup_logging
from .monthly_runner import run_monthly

KST = timezone(timedelta(hours=9))
log = logging.getLogger(__name__)


def _parse_date(s: str) -> date:
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        raise argparse.ArgumentTypeError(f"날짜 형식 오류 (YYYY-MM-DD 필요): {s}")


def main() -> int:
    parser = argparse.ArgumentParser(description="B2C 고객지표 자동 업데이트 RPA")
    sub = parser.add_subparsers(dest="mode", required=True)

    # daily
    p_daily = sub.add_parser("daily", help="일별 업데이트 (3개 OLAP 리포트 → 2개 시트)")
    p_daily.add_argument("--date", type=_parse_date, help="대상 날짜 (기본: 어제, YYYY-MM-DD)")
    p_daily.add_argument("--force", action="store_true", help="이미 완료된 날짜도 재실행")

    # monthly
    p_monthly = sub.add_parser("monthly", help="월별 업데이트 (LOG REPORT / VISUAL REPORT)")
    p_monthly.add_argument("--year",  type=int, help="대상 연도 (기본: 이번 달)")
    p_monthly.add_argument("--month", type=int, help="대상 월 (기본: 이번 달)")
    p_monthly.add_argument("--force", action="store_true", help="이미 완료된 달도 재실행")

    args = parser.parse_args()

    try:
        config = load_config()
    except RuntimeError as exc:
        print(f"[ERROR] 설정 오류: {exc}", file=sys.stderr)
        return 2

    setup_logging(config.runtime.logs_dir)

    try:
        if args.mode == "daily":
            target = args.date or yesterday_kst()
            run_daily(config, target_date=target, force=args.force)

        elif args.mode == "monthly":
            now = datetime.now(KST)
            year  = args.year  or now.year
            month = args.month or now.month
            run_monthly(config, year=year, month=month, force=args.force)

    except SystemExit:
        raise
    except Exception as exc:
        log.exception(f"[FATAL] 예기치 않은 오류: {exc}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
