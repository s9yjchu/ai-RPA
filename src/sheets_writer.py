"""Google Sheets 쓰기 — 날짜 행 탐색·생성 후 지표 업데이트."""

from __future__ import annotations

import logging
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from .config import GOOGLE_OAUTH_SCOPES

log = logging.getLogger(__name__)

# 단일 출처(config.GOOGLE_OAUTH_SCOPES) 사용 — 설치 시 동의 범위가
# sheets/gmail/drive 를 모두 포함하도록 통일. setup_helper.do_auth 가 이 목록으로 동의.
SCOPES = GOOGLE_OAUTH_SCOPES

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
        # 스코프를 강제하지 않고 토큰에 부여된 스코프로 로드 — 새 스코프(drive.file)를
        # 강제하면 미재인증 토큰의 refresh 가 invalid_scope 로 깨진다.
        # 전체 스코프(SCOPES) 요청은 신규 동의(InstalledAppFlow) 시에만 적용한다.
        creds = Credentials.from_authorized_user_file(str(token_path))

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

# Google Sheets 날짜 시리얼 기준일 (1899-12-30 == serial 0)
_SHEETS_EPOCH = date(1899, 12, 30)


def _parse_cell_date(val: str) -> date | None:
    """셀 문자열 값을 date 로 파싱 (다양한 형식 처리).

    한국 로케일 표시('2021. 1. 1')·요일 접미('2022-01-01(토)')도 처리하도록
    숫자만 추출해 재구성한다. 유니코드 대시는 ASCII 로 정규화.
    """
    if not val:
        return None
    s = str(val).strip().replace("−", "-").replace("–", "-").replace("—", "-")
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # fallback: 앞쪽 연-월-일 숫자 3개를 직접 추출 ('2021. 1. 1', '2022-01-01(토)' 등)
    import re
    m = re.match(r"\s*(\d{4})\D+(\d{1,2})\D+(\d{1,2})", s)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def _cell_to_date(v: Any) -> date | None:
    """UNFORMATTED 셀 값(날짜 시리얼 정수/실수 또는 문자열)을 date 로 변환."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        try:
            return _SHEETS_EPOCH + timedelta(days=int(v))
        except (ValueError, OverflowError):
            return None
    if isinstance(v, str):
        return _parse_cell_date(v)
    return None


def build_date_row_map(ws: gspread.Worksheet) -> dict[date, int]:
    """col B(일자)를 UNFORMATTED 로 읽어 {date: 1-based 첫 행번호} 맵을 만든다.

    날짜는 실제 시리얼로 저장되어 있으나 표시 형식이 시트마다 달라(FORMATTED 파싱 불가),
    반드시 UNFORMATTED 로 읽어 시리얼→date 로 변환한다. 중복 날짜는 **첫(정본) 행** 유지.
    """
    col_b = ws.col_values(2, value_render_option="UNFORMATTED_VALUE")
    m: dict[date, int] = {}
    for i, cell in enumerate(col_b, start=1):
        d = _cell_to_date(cell)
        if d is not None and d not in m:
            m[d] = i
    return m


def _find_row_by_date(ws: gspread.Worksheet, target_date: date) -> int | None:
    """column B (날짜) 에서 target_date 와 일치하는 1-based 행 번호 반환."""
    return build_date_row_map(ws).get(target_date)


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
    # 행은 staff 가 미리 생성해 두므로, 미발견은 (신규 최신일을 제외하면) 이상 신호.
    log.warning(
        f"  [SHEETS] {target_date} 기존 행 없음 → 새 행 append "
        "(미리 생성된 행 범위를 벗어난 신규일이 아니라면 날짜 형식/시트를 확인하세요)"
    )
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
