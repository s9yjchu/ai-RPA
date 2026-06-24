"""구글시트 사전 무결성 점검 — 백필 대상 날짜 탐지 + 중복 행 정리.

행은 staff 가 미리 생성(2021/2022~)하므로 "누락"은 *행 부재*가 아니라
**기존 행의 RPA 담당 셀 공백**이다. 이 모듈이 그 공백 날짜를 찾아 백필 작업 목록을 만든다.

CLI:
  python -m src.integrity --blanks [--month YYYY-MM]   # 공백 날짜 목록 출력
  python -m src.integrity --dedupe [--apply]           # 중복 행 탐지(기본 dry-run) / 삭제
"""

from __future__ import annotations

import argparse
import logging
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone

from .sheets_writer import (
    SHEET_HPC_DAILY,
    SHEET_STORE_DAILY,
    _cell_to_date,
    _col_letter,
    build_date_row_map,
)

log = logging.getLogger(__name__)
KST = timezone(timedelta(hours=9))

# RPA 담당 컬럼(1-based). 이 중 하나라도 공백이면 해당 날짜를 백필 대상으로 본다.
# HPC 실적(일별): 신규회원수(4), 해피앱 로그인수(5), 해피앱 DAU(6), 해피오더 DAU(7)
HPC_RPA_COLS = [4, 5, 6, 7]
# 전사 매장실적(일별): POS·HPC 지표(3~12) + APP 제시건수(14). 13(HPC 가입수)은 수동(제외).
STORE_RPA_COLS = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14]

_SHEET_SPECS = [
    (SHEET_HPC_DAILY, HPC_RPA_COLS),
    (SHEET_STORE_DAILY, STORE_RPA_COLS),
]


def _daterange(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def current_month_window(today: date | None = None) -> tuple[date, date]:
    """당월 1일 ~ 어제(KST). 월말이면 자연히 한 달 전체를 포함."""
    if today is None:
        today = datetime.now(KST).date()
    start = today.replace(day=1)
    end = today - timedelta(days=1)
    return start, end


def _is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def find_blank_dates(spreadsheet, start: date, end: date, specs=None) -> set[date]:
    """[start, end] 구간에서 RPA 담당 셀이 하나라도 공백인 기존 행의 날짜 집합.

    specs: [(시트명, [1-based 컬럼...]), ...]. 기본은 전체 RPA 컬럼(_SHEET_SPECS).
    자동 갭필은 신규회원수 등 상시 공백 가능 컬럼을 제외한 specs 를 넘길 수 있다.
    """
    result: set[date] = set()
    if end < start:
        return result
    window = set(_daterange(start, end))

    for title, rpa_cols in (specs or _SHEET_SPECS):
        ws = spreadsheet.worksheet(title)
        dmap = build_date_row_map(ws)  # {date: 1-based row}
        rows = sorted((dmap[d], d) for d in window if d in dmap)
        if not rows:
            continue
        minr, maxr = rows[0][0], rows[-1][0]
        last_col = _col_letter(max(rpa_cols))
        rng = ws.get(
            f"A{minr}:{last_col}{maxr}", value_render_option="UNFORMATTED_VALUE"
        )
        for r, d in rows:
            vals = rng[r - minr] if (r - minr) < len(rng) else []
            for c in rpa_cols:
                v = vals[c - 1] if (c - 1) < len(vals) else None
                if _is_blank(v):
                    result.add(d)
                    break
    return result


def find_duplicate_rows(ws) -> dict[date, list[int]]:
    """동일 날짜가 2개 이상인 경우 {date: [삭제후보 행번호...]} 반환.

    정본 = 최초(가장 작은 행번호), 삭제후보 = 그 이후 중복 행(append 된 행).
    """
    col_b = ws.col_values(2, value_render_option="UNFORMATTED_VALUE")
    first: dict[date, int] = {}
    extras: dict[date, list[int]] = defaultdict(list)
    for i, v in enumerate(col_b, start=1):
        d = _cell_to_date(v)
        if d is None:
            continue
        if d in first:
            extras[d].append(i)
        else:
            first[d] = i
    return dict(extras)


def dedupe(spreadsheet, apply: bool = False) -> dict[str, dict[date, list[int]]]:
    """중복 행 탐지(기본 dry-run). apply=True 시 중복(append) 행 삭제.

    삭제는 비가역이므로 기본은 목록만 반환한다. 삭제 시 행번호 큰 것부터 제거(재인덱싱 방지).
    """
    report: dict[str, dict[date, list[int]]] = {}
    for title, _ in _SHEET_SPECS:
        ws = spreadsheet.worksheet(title)
        dups = find_duplicate_rows(ws)
        report[title] = dups
        if not dups:
            log.info(f"  [{title}] 중복 없음")
            continue
        rows_to_delete = sorted(
            (r for rows in dups.values() for r in rows), reverse=True
        )
        log.warning(
            f"  [{title}] 중복 {len(dups)}일, 삭제후보 {len(rows_to_delete)}행: "
            + ", ".join(f"{d}->{rows}" for d, rows in sorted(dups.items()))
        )
        if apply:
            for r in rows_to_delete:  # 큰 행번호부터
                ws.delete_rows(r)
            log.warning(f"  [{title}] {len(rows_to_delete)}행 삭제 완료")
    return report


# ── CLI ───────────────────────────────────────────────────────────────

def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    ap = argparse.ArgumentParser(description="구글시트 무결성 점검")
    ap.add_argument("--blanks", action="store_true", help="공백(백필 대상) 날짜 목록 출력")
    ap.add_argument("--month", help="YYYY-MM (기본: 당월)")
    ap.add_argument("--dedupe", action="store_true", help="중복 행 탐지")
    ap.add_argument("--apply", action="store_true", help="--dedupe 와 함께: 실제 삭제")
    args = ap.parse_args()

    from .config import load_config
    from .sheets_writer import open_spreadsheet

    c = load_config()
    sh = open_spreadsheet(
        c.sheets.credentials_path, c.sheets.token_path, c.sheets.spreadsheet_id
    )

    if args.dedupe:
        dedupe(sh, apply=args.apply)
        if not args.apply:
            print("\n(dry-run) 실제 삭제하려면 --apply 추가. 행 삭제는 비가역입니다.")

    if args.blanks or not (args.dedupe):
        if args.month:
            y, m = map(int, args.month.split("-"))
            start = date(y, m, 1)
            nxt = date(y + (m // 12), (m % 12) + 1, 1)
            end = min(nxt - timedelta(days=1), datetime.now(KST).date() - timedelta(days=1))
        else:
            start, end = current_month_window()
        blanks = sorted(find_blank_dates(sh, start, end))
        print(f"\n공백(백필 대상) 날짜 {len(blanks)}개 [{start}~{end}]:")
        for d in blanks:
            print(" ", d)


if __name__ == "__main__":
    main()
