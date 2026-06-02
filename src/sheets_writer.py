"""Google Sheets 쓰기 — 날짜 행 탐색·생성 후 지표 업데이트."""

from __future__ import annotations

import logging
from datetime import date, datetime
from pathlib import Path
from typing import Any

import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

log = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ── 대상 시트 이름 ────────────────────────────────────────────────────
SHEET_HPC_DAILY   = "HPC 실적 (일별)"
SHEET_STORE_DAILY = "전사 매장실적 (일별)"
SHEET_HPC_MONTHLY = "HPC 실적 (월별)"

# ── 컬럼 인덱스 (1-based, A=1) ────────────────────────────────────────
# HPC 실적 (일별)
HPC_COLS: dict[str, int] = {
    "월":           1,
    "날짜":         2,
    "요일":         3,
    "신규회원수":    4,   # ← 신규가입회원수
    "해피앱 로그인수": 5, # ← 해피앱 로그인 회원수
    "해피앱 DAU":   6,
    "해피오더 DAU": 7,
}

# 전사 매장실적 (일별)
STORE_COLS: dict[str, int] = {
    "월":           1,
    "날짜":         2,
    # 요일 없음
    "POS 총매출액": 3,
    "POS 영수증건수": 4,
    "POS 거래점포수": 5,
    "HPC 매출액":   6,
    "HPC 거래점포수": 7,
    "HPC 총적립액": 8,
    "HPC 적립건수": 9,
    "객단가":       10,
    "HPC 총사용액": 11,
    "HPC 사용건수": 12,
    # 13 = HPC 가입수 (비대상)
    "APP 제시건수": 14,
}

# HPC 실적 (월별) — 월별 업데이트 대상 컬럼 (col O=15, col P=16)
HPC_MONTHLY_COLS: dict[str, int] = {
    "해피앱 월 로그인객수": 15,  # col O ← LOG REPORT 순 로그인 회원수
    "해피앱 MAU":          16,  # col P ← VISUAL REPORT MAU 당월
}

KR_DOW = ["월", "화", "수", "목", "금", "토", "일"]  # Monday=0


# ── 인증 ─────────────────────────────────────────────────────────────

def get_client(credentials_path: Path, token_path: Path) -> gspread.Client:
    creds: Credentials | None = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path), SCOPES
            )
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return gspread.authorize(creds)


def open_spreadsheet(
    credentials_path: Path,
    token_path: Path,
    spreadsheet_id: str,
) -> gspread.Spreadsheet:
    client = get_client(credentials_path, token_path)
    return client.open_by_key(spreadsheet_id)


# ── 날짜 행 탐색 / 생성 ──────────────────────────────────────────────

def _parse_cell_date(val: str) -> date | None:
    """셀 문자열 값을 date 로 파싱 (다양한 형식 처리)."""
    if not val:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(val.strip(), fmt).date()
        except ValueError:
            continue
    return None


def _find_row_by_date(ws: gspread.Worksheet, target_date: date) -> int | None:
    """column B (날짜) 에서 target_date 와 일치하는 1-based 행 번호 반환."""
    col_b = ws.col_values(2)  # 1-based → list index 0
    for i, cell in enumerate(col_b):
        parsed = _parse_cell_date(cell)
        if parsed == target_date:
            return i + 1  # 1-based
    return None


def _append_date_row(
    ws: gspread.Worksheet,
    target_date: date,
    include_dow: bool,
) -> int:
    """마지막 행 아래에 날짜 관련 셀(월, 날짜, 요일)을 삽입하고 행 번호 반환."""
    yyyymm = target_date.strftime("%Y%m")
    date_str = target_date.strftime("%Y-%m-%d")
    dow = KR_DOW[target_date.weekday()]

    # 마지막 데이터 행 탐색
    col_a = ws.col_values(1)
    last_row = len(col_a)
    new_row = last_row + 1

    ws.update_cell(new_row, 1, yyyymm)
    ws.update_cell(new_row, 2, date_str)
    if include_dow:
        ws.update_cell(new_row, 3, dow)

    log.info(f"  새 행 추가: row={new_row}, 날짜={date_str}")
    return new_row


def find_or_create_row(
    ws: gspread.Worksheet,
    target_date: date,
    include_dow: bool = False,
) -> int:
    """날짜 행이 있으면 그 번호를, 없으면 새로 생성해서 반환."""
    row = _find_row_by_date(ws, target_date)
    if row is not None:
        log.info(f"  기존 행 발견: row={row}")
        return row
    log.info(f"  행 없음 → 새 행 생성")
    return _append_date_row(ws, target_date, include_dow=include_dow)


# ── 셀 배치 업데이트 ──────────────────────────────────────────────────

def _build_batch(
    row: int, metrics: dict[str, Any], col_map: dict[str, int], sheet_title: str
) -> list[dict]:
    """gspread batch_update 용 cell 리스트 생성. 시트명을 range 에 포함해 대상 시트를 명시."""
    cells = []
    for field, col in col_map.items():
        if field in ("월", "날짜", "요일"):
            continue  # 날짜 필드는 find_or_create_row 에서 이미 처리
        if field not in metrics:
            continue
        val = metrics[field]
        if val is None:
            continue
        col_letter = _col_letter(col)
        cells.append({
            "range": f"'{sheet_title}'!{col_letter}{row}",
            "values": [[val]],
        })
    return cells


def _col_letter(n: int) -> str:
    """1-based 컬럼 번호를 A, B, …, Z, AA, … 로 변환."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ── 공개 쓰기 함수 ────────────────────────────────────────────────────

def write_hpc_daily(
    spreadsheet: gspread.Spreadsheet,
    target_date: date,
    metrics: dict[str, Any],
    dry_run: bool = False,
) -> None:
    """HPC 실적 (일별) 시트에 지표를 씁니다."""
    ws = spreadsheet.worksheet(SHEET_HPC_DAILY)
    row = find_or_create_row(ws, target_date, include_dow=True)
    batch = _build_batch(row, metrics, HPC_COLS, ws.title)

    if dry_run:
        log.info(f"  [DRY_RUN] HPC 일별 — 쓰기 생략: {batch}")
        return

    if batch:
        spreadsheet.values_batch_update(
            {"valueInputOption": "RAW", "data": batch}
        )
    log.info(f"  [SHEETS] HPC 실적 (일별) row={row} 업데이트 완료")


def write_store_daily(
    spreadsheet: gspread.Spreadsheet,
    target_date: date,
    metrics: dict[str, Any],
    dry_run: bool = False,
) -> None:
    """전사 매장실적 (일별) 시트에 지표를 씁니다."""
    ws = spreadsheet.worksheet(SHEET_STORE_DAILY)
    row = find_or_create_row(ws, target_date, include_dow=False)
    batch = _build_batch(row, metrics, STORE_COLS, ws.title)

    if dry_run:
        log.info(f"  [DRY_RUN] 전사 매장실적 — 쓰기 생략: {batch}")
        return

    if batch:
        spreadsheet.values_batch_update(
            {"valueInputOption": "RAW", "data": batch}
        )
    log.info(f"  [SHEETS] 전사 매장실적 (일별) row={row} 업데이트 완료")


# ── 월별 시트 ─────────────────────────────────────────────────────────

def _find_row_by_month(ws: gspread.Worksheet, year: int, month: int) -> int | None:
    """col A (YYYYMM 형식) 에서 연월과 일치하는 1-based 행 번호 반환."""
    target = f"{year}{month:02d}"
    col_a = ws.col_values(1)
    for i, cell in enumerate(col_a):
        if str(cell).strip() == target:
            return i + 1
    return None


def write_hpc_monthly(
    spreadsheet: gspread.Spreadsheet,
    year: int,
    month: int,
    metrics: dict[str, Any],
    dry_run: bool = False,
) -> None:
    """HPC 실적 (월별) 시트에 월별 지표를 씁니다.

    월별 행은 staff 가 수동으로 관리하므로, 행이 없으면 ValueError 를 발생시킵니다.
    """
    ws = spreadsheet.worksheet(SHEET_HPC_MONTHLY)
    row = _find_row_by_month(ws, year, month)
    if row is None:
        raise ValueError(
            f"[SHEETS] {year}-{month:02d} 행을 찾을 수 없음 (col A YYYYMM 형식 확인 필요)"
        )
    batch = _build_batch(row, metrics, HPC_MONTHLY_COLS, ws.title)

    if dry_run:
        log.info(f"  [DRY_RUN] HPC 월별 — 쓰기 생략: {batch}")
        return

    if batch:
        spreadsheet.values_batch_update(
            {"valueInputOption": "RAW", "data": batch}
        )
    log.info(f"  [SHEETS] HPC 실적 (월별) row={row} 업데이트 완료")
