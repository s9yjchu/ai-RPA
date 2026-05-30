"""다운로드된 OLAP Excel 파일 파싱 — 3개 리포트별 데이터 추출.

# ── 파싱 튜닝 가이드 ─────────────────────────────────────────────────
# OLAP 에서 다운로드한 Excel 파일의 실제 컬럼명/구조를 확인하려면:
#   python -m src.excel_parser <파일경로>
# 위 명령으로 헤더와 샘플 행이 출력됩니다.
#
# 컬럼명이 다를 경우 아래 COLUMN_MAP_* 상수를 수정하세요.
# ────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import logging
import sys
from datetime import date
from pathlib import Path
from typing import Any

import openpyxl

log = logging.getLogger(__name__)

# ── 리포트 1: [HPC] 일별 회원관리지표 ──────────────────────────────
# Excel 컬럼명 → 대상 시트 컬럼명 매핑 (왼쪽: OLAP Excel, 오른쪽: Google Sheets)
MEMBER_COL_MAP: dict[str, str] = {
    "신규가입회원수":     "신규회원수",
    "해피앱 로그인 회원수": "해피앱 로그인수",
    "해피앱 DAU":        "해피앱 DAU",
    "해피오더 DAU":      "해피오더 DAU",
}
# Excel 에서 날짜가 들어있는 컬럼명 (행 필터링용)
MEMBER_DATE_COL = "일자"  # TODO: 실제 OLAP Excel 헤더 확인 후 수정

# ── 리포트 2: [HPC] 채널별 적립, 사용건수 현황 ─────────────────────
# 채널명 행 식별값
CHANNEL_TARGET_LABEL = "HPCAPP"
# 채널명이 들어있는 컬럼명
CHANNEL_KEY_COL = "채널"       # TODO: 실제 OLAP Excel 헤더 확인 후 수정
# 건수 값이 들어있는 컬럼명
CHANNEL_VALUE_COL = "제시건수"   # TODO: 실제 OLAP Excel 헤더 확인 후 수정

# ── 리포트 3: [HPC, POS] HPC 일마감(브랜드) ────────────────────────
# 브랜드 식별자 — 컬럼 헤더로 존재 (유저 확인)
CLOSING_BRAND_LABEL = "0002. SPC전사(3사)"
# 메트릭 행 레이블이 들어있는 컬럼명 (행 인덱스 열)
CLOSING_ROW_KEY_COL = 0  # 첫 번째 열 (0-indexed) — TODO: 실제 구조 확인 후 수정
# 메트릭 행 레이블 → 대상 시트 컬럼명 매핑
CLOSING_ROW_MAP: dict[str, str] = {
    "POS 총매출액":  "POS 총매출액",
    "POS 영수증건수": "POS 영수증건수",
    "POS 거래점포수": "POS 거래점포수",
    "HPC 매출액":    "HPC 매출액",
    "HPC 거래점포수": "HPC 거래점포수",
    "HPC 총적립액":  "HPC 총적립액",
    "HPC 적립건수":  "HPC 적립건수",
    "객단가":        "객단가",
    "HPC 총사용액":  "HPC 총사용액",
    "HPC 사용건수":  "HPC 사용건수",
}


# ── 공통 유틸 ─────────────────────────────────────────────────────────

def _load_first_sheet(path: Path) -> openpyxl.worksheet.worksheet.Worksheet:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active or wb[wb.sheetnames[0]]
    log.debug(f"  엑셀 로드: {path.name} / 시트={ws.title} / {ws.max_row}행×{ws.max_column}열")
    return ws


def _header_row(ws) -> list[str]:
    """첫 번째 비어있지 않은 행을 헤더로 반환."""
    for row in ws.iter_rows(values_only=True):
        cleaned = [str(c).strip() if c is not None else "" for c in row]
        if any(cleaned):
            return cleaned
    return []


def _to_num(val: Any) -> float | None:
    if val is None:
        return None
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def _dates_match(cell_val: Any, target: date) -> bool:
    """셀 값(날짜 직렬/문자열/datetime)이 target date 와 일치하는지 확인."""
    if cell_val is None:
        return False
    from datetime import datetime as dt
    if isinstance(cell_val, (dt,)):
        return cell_val.date() == target
    if isinstance(cell_val, date):
        return cell_val == target
    # 문자열 형식 시도
    s = str(cell_val).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y"):
        try:
            return dt.strptime(s, fmt).date() == target
        except ValueError:
            continue
    return False


def _debug_dump(path: Path) -> None:
    """헤더 + 상위 5행 출력 (파싱 튜닝용)."""
    ws = _load_first_sheet(path)
    headers = _header_row(ws)
    print(f"\n=== {path.name} ===")
    print("헤더:", headers)
    count = 0
    for row in ws.iter_rows(values_only=True):
        if count >= 5:
            break
        if any(c is not None for c in row):
            print("행:", row[:20])
            count += 1


# ── 리포트 1 파서 ─────────────────────────────────────────────────────

def parse_member_metrics(path: Path, target_date: date) -> dict[str, Any]:
    """[HPC] 일별 회원관리지표 Excel → 어제 날짜 행 값 추출."""
    log.info(f"[PARSE] 회원지표 파싱: {path.name}")
    ws = _load_first_sheet(path)
    headers = _header_row(ws)

    if not headers:
        raise ValueError(f"헤더를 찾을 수 없습니다: {path}")

    # 날짜 컬럼 인덱스
    date_idx: int | None = None
    for i, h in enumerate(headers):
        if h == MEMBER_DATE_COL or "일자" in h or "날짜" in h or "date" in h.lower():
            date_idx = i
            break

    if date_idx is None:
        log.warning(
            f"  날짜 컬럼 '{MEMBER_DATE_COL}' 미발견. 헤더: {headers[:15]}\n"
            "  MEMBER_DATE_COL 상수를 실제 컬럼명으로 수정하세요."
        )
        # fallback: 첫 번째 데이터 행의 값으로 매핑
        date_idx = 0

    # 대상 행 탐색
    target_row: list | None = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        if _dates_match(row[date_idx], target_date):
            target_row = list(row)
            break

    if target_row is None:
        raise ValueError(
            f"날짜 {target_date} 에 해당하는 행이 없습니다 (파일: {path.name}). "
            "데이터가 아직 준비되지 않았거나, MEMBER_DATE_COL 컬럼명을 확인하세요."
        )

    result: dict[str, Any] = {}
    for src_col, dst_col in MEMBER_COL_MAP.items():
        if src_col in headers:
            idx = headers.index(src_col)
            result[dst_col] = _to_num(target_row[idx])
        else:
            log.warning(f"  컬럼 '{src_col}' 미발견 (헤더: {headers[:15]})")
            result[dst_col] = None

    log.info(f"  추출 결과: {result}")
    return result


# ── 리포트 2 파서 ─────────────────────────────────────────────────────

def parse_channel_metrics(path: Path, target_date: date) -> dict[str, Any]:
    """[HPC] 채널별 적립, 사용건수 현황 → HPCAPP 제시건수 추출."""
    log.info(f"[PARSE] 채널별 지표 파싱: {path.name}")
    ws = _load_first_sheet(path)
    headers = _header_row(ws)

    if not headers:
        raise ValueError(f"헤더를 찾을 수 없습니다: {path}")

    # 날짜 컬럼 탐색 (있는 경우)
    date_idx: int | None = None
    for i, h in enumerate(headers):
        if "일자" in h or "날짜" in h or "date" in h.lower():
            date_idx = i
            break

    # 채널 컬럼 인덱스
    key_idx: int | None = None
    for i, h in enumerate(headers):
        if h == CHANNEL_KEY_COL or "채널" in h or "channel" in h.lower():
            key_idx = i
            break

    # 값 컬럼 인덱스
    val_idx: int | None = None
    for i, h in enumerate(headers):
        if h == CHANNEL_VALUE_COL or "제시건수" in h or "건수" in h:
            val_idx = i
            break

    if key_idx is None or val_idx is None:
        log.warning(
            f"  채널 컬럼 미발견. 헤더: {headers[:15]}\n"
            "  CHANNEL_KEY_COL / CHANNEL_VALUE_COL 을 수정하세요."
        )
        raise ValueError(f"채널 컬럼 미발견: {path.name}")

    # HPCAPP 행 탐색 (날짜 필터 적용 가능하면 적용)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if date_idx is not None and not _dates_match(row[date_idx], target_date):
            continue
        cell_key = str(row[key_idx]).strip() if row[key_idx] else ""
        if CHANNEL_TARGET_LABEL in cell_key:
            val = _to_num(row[val_idx])
            result = {"APP 제시건수": val}
            log.info(f"  추출 결과: {result}")
            return result

    raise ValueError(
        f"'{CHANNEL_TARGET_LABEL}' 행이 없습니다 (파일: {path.name}). "
        "CHANNEL_TARGET_LABEL / CHANNEL_KEY_COL 을 확인하세요."
    )


# ── 리포트 3 파서 ─────────────────────────────────────────────────────

def parse_closing_report(path: Path) -> dict[str, Any]:
    """[HPC, POS] HPC 일마감(브랜드) → '0002. SPC전사(3사)' 컬럼 추출.

    리포트 구조 (유저 확인):
        행 = 메트릭명, 열 = 브랜드명 (0002. SPC전사(3사) 포함)
    """
    log.info(f"[PARSE] 일마감 리포트 파싱: {path.name}")
    ws = _load_first_sheet(path)
    headers = _header_row(ws)

    if not headers:
        raise ValueError(f"헤더를 찾을 수 없습니다: {path}")

    # "0002. SPC전사(3사)" 컬럼 인덱스 탐색
    brand_idx: int | None = None
    for i, h in enumerate(headers):
        if CLOSING_BRAND_LABEL in h:
            brand_idx = i
            break

    if brand_idx is None:
        # fallback: 부분 문자열 "SPC전사" 로 재탐색
        for i, h in enumerate(headers):
            if "SPC전사" in h or "전사(3사)" in h:
                brand_idx = i
                log.warning(f"  '{CLOSING_BRAND_LABEL}' 대신 '{headers[i]}' 컬럼 사용")
                break

    if brand_idx is None:
        log.warning(
            f"  브랜드 컬럼 '{CLOSING_BRAND_LABEL}' 미발견. 헤더: {headers[:20]}\n"
            "  CLOSING_BRAND_LABEL 을 실제 컬럼명으로 수정하세요."
        )
        raise ValueError(f"브랜드 컬럼 미발견: {path.name}")

    # 메트릭 행 추출
    result: dict[str, Any] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_key = str(row[CLOSING_ROW_KEY_COL]).strip() if row[CLOSING_ROW_KEY_COL] else ""
        for src_label, dst_col in CLOSING_ROW_MAP.items():
            if src_label in row_key:
                result[dst_col] = _to_num(row[brand_idx])
                break

    missing = [v for v in CLOSING_ROW_MAP.values() if v not in result]
    if missing:
        log.warning(
            f"  미추출 항목: {missing}\n"
            "  CLOSING_ROW_MAP 의 행 레이블을 실제 Excel 값으로 수정하세요."
        )

    log.info(f"  추출 결과: {result}")
    return result


# ── CLI 디버그 모드 ───────────────────────────────────────────────────

if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG, format="%(message)s")
    if len(sys.argv) < 2:
        print("사용법: python -m src.excel_parser <파일경로>")
        sys.exit(1)
    _debug_dump(Path(sys.argv[1]))
