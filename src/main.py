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
from .daily_runner import run_backfill, run_daily, yesterday_kst
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

    # backfill — 여러 날짜 일괄 처리 (빈 셀 채움 / 잘못 쓰인 값 교정)
    p_bf = sub.add_parser("backfill", help="여러 날짜 일괄 백필 (당월 공백 또는 지정 월 전체)")
    p_bf.add_argument("--month", help="YYYY-MM 전체 강제 재처리 (생략 시: 당월 공백만)")

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

        elif args.mode == "backfill":
            # NOTE: yesterday_kst / timedelta 는 모듈 상단에서 이미 import 됨. 함수 안에서
            # 다시 import 하면 그 이름이 함수 전체의 지역변수로 취급돼 daily 분기(line 64)가
            # UnboundLocalError 가 난다. 여기서는 모듈 전역만 사용한다.
            from .integrity import current_month_window, find_blank_dates
            from .sheets_writer import open_spreadsheet

            if args.month:
                y, m = map(int, args.month.split("-"))
                start = date(y, m, 1)
                nxt = date(y + (m // 12), (m % 12) + 1, 1)
                end = min(nxt - timedelta(days=1), yesterday_kst())
                dates = []
                d = start
                while d <= end:
                    dates.append(d)
                    d += timedelta(days=1)
                log.info(f"[BACKFILL] {args.month} 전체 {len(dates)}일 강제 재처리")
            else:
                sh = open_spreadsheet(
                    config.sheets.credentials_path,
                    config.sheets.token_path,
                    config.sheets.spreadsheet_id,
                )
                s, e = current_month_window()
                dates = sorted(find_blank_dates(sh, s, e))
                log.info(f"[BACKFILL] 당월 공백 {len(dates)}일: {dates}")
            run_backfill(config, dates, force=True)

    except SystemExit:
        raise
    except Exception as exc:
        log.exception(f"[FATAL] 예기치 않은 오류: {exc}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
