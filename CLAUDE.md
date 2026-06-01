# CLAUDE.md

This file provides guidance to Claude Code when working in this repository.

## Purpose

B2C 고객지표 관리 자동화 RPA. 매일 SASHBI OLAP 에서 3개 리포트를 다운로드하고
Google Sheets "B2C사업본부 고객지표" 의 2개 시트를 업데이트한다.
매월 1일에는 LOG REPORT / VISUAL REPORT 에서 월별 지표를 추가로 업데이트한다.

- **OLAP**: `https://dwweb.spc.co.kr:7980/SASHBI/main.jsp` (내부망 SSO)
- **LOG REPORT**: `https://hplog.spc.co.kr:8000/datastory/home` (내부망 SSO)
- **VISUAL REPORT**: `https://va.spc.co.kr/SASReportViewer/` (내부망 SSO)
- **Sheets ID**: `1gIEbHzyfh4TG21M1etxWxh0u99ylGMsX4B_VGEu3Yxw`

## Runtime

- Python 3 (`python` 또는 `python3`)
- 가상환경: `.venv/` (없으면 시스템 Python)
- 브라우저: Playwright Chromium (`python -m playwright install chromium`)
- 외부 의존: `requirements.txt` 참고

## Entry Point

```
python -m src.main daily              # 어제 날짜 일별 업데이트
python -m src.main daily --date YYYY-MM-DD --force
python -m src.main monthly            # 이번 달 월별 업데이트 (매월 1일 실행)
python -m src.main monthly --year 2025 --month 5 --force
```

## Module Layout

```
src/
├── main.py                CLI argparse 진입점. daily / monthly 서브커맨드.
├── config.py              .env → 동결 dataclass (Config, OlapConfig, LogReportConfig,
│                          VisualReportConfig, SheetsConfig, …).
│                          load_config() 단일 함수로 전체 설정 반환.
├── logger.py              setup_logging(logs_dir) — KST 타임스탬프, 파일+콘솔.
├── browser.py             BrowserSession context manager. snapshot() 로 디버그 덤프.
│                          ignore_https_errors=True (내부망 자체서명 인증서 대응).
├── state_manager.py       DailyState / MonthlyState — state/ 폴더에 JSON 저장.
│                          should_give_up() 로 재시도 시간 초과 판단.
├── olap_scraper.py        OLAP 자동화의 핵심. SEL_* 상수로 셀렉터 분리.
│                          scrape_report(page, report_key, …) 가 한 리포트 전체 플로우.
│                          DataNotReadyError — 날짜 미확인 시 raise.
├── log_report_scraper.py  LOG REPORT 자동화. SEL_LR_* 상수로 셀렉터 분리.
│                          scrape_login_count(config, year, month) → int.
│                          해피앱 > 종합 > 유저 접속 추이 > 월별 보기 탐색.
├── visual_report_scraper.py  VISUAL REPORT 자동화. SEL_VR_* 상수로 셀렉터 분리.
│                          scrape_mau_excel(config, year, month) → Path.
│                          리포트 찾아보기 > 클라우드 > 프로모션 > 해피앱 GA 리포트.
│                          필터 지우기 → MAU 당월 Excel 다운로드. 로딩 최대 360초.
├── excel_parser.py        4개 파서 (parse_member_metrics, parse_channel_metrics,
│                          parse_closing_report, parse_mau_excel). 상단 상수로 컬럼명 분리.
│                          CLI: python -m src.excel_parser <파일> 로 헤더 확인.
├── sheets_writer.py       gspread 기반. find_or_create_row() 로 날짜 행 탐색/생성.
│                          _build_batch() → values_batch_update() 로 한 번에 쓰기.
│                          write_hpc_daily / write_store_daily / write_hpc_monthly.
├── notifier.py            Gmail API. notify_success / notify_failure /
│                          notify_data_not_ready 3가지.
├── daily_runner.py        run_daily() — 다운로드 → 파싱 → 쓰기 → 알림 순서 조율.
│                          DataNotReadyError 는 실패가 아닌 재시도 대기로 처리.
└── monthly_runner.py      run_monthly() — LOG REPORT + VISUAL REPORT → Sheets 업데이트.
                           DataNotReadyError 재시도 처리. 매월 1일 실행.
```

## 대상 Sheets 컬럼 인덱스 (1-based)

### HPC 실적 (월별)
| col | 필드 | 작성 주체 |
|---|---|---|
| 15(O) | 해피앱 월 로그인객수 | LOG REPORT |
| 16(P) | 해피앱 MAU | VISUAL REPORT |

행 키: col A (YYYYMM 형식). 행은 staff 가 수동 관리 — RPA 는 기존 행만 업데이트.

### HPC 실적 (일별)
| col | 필드 | 작성 주체 |
|---|---|---|
| 1(A) | 월 | RPA |
| 2(B) | 날짜 | RPA |
| 3(C) | 요일 | RPA |
| 4(D) | 신규회원수 | 리포트1 |
| 5(E) | 해피앱 로그인수 | 리포트1 |
| 6(F) | 해피앱 DAU | 리포트1 |
| 7(G) | 해피오더 DAU | 리포트1 |
| 8–16 | (기타) | 수동 |

### 전사 매장실적 (일별)
| col | 필드 | 작성 주체 |
|---|---|---|
| 1(A) | 월 | RPA |
| 2(B) | 날짜 | RPA |
| 3(C) | POS 총매출액 | 리포트3 |
| 4(D) | POS 영수증건수 | 리포트3 |
| 5(E) | POS 거래점포수 | 리포트3 |
| 6(F) | HPC 매출액 | 리포트3 |
| 7(G) | HPC 거래점포수 | 리포트3 |
| 8(H) | HPC 총적립액 | 리포트3 |
| 9(I) | HPC 적립건수 | 리포트3 |
| 10(J) | 객단가 | 리포트3 |
| 11(K) | HPC 총사용액 | 리포트3 |
| 12(L) | HPC 사용건수 | 리포트3 |
| 13(M) | HPC 가입수 | 수동 (비대상) |
| 14(N) | APP 제시건수 | 리포트2 |
| 15–16 | (기타) | 수동 |

## OLAP 리포트 경로

```python
REPORT_PATHS = {
    "member_metrics":  ["정형데이터", "회원분석", "[HPC] 일별 회원관리지표"],
    "channel_metrics": ["정형데이터", "[HPC] 채널별 적립, 사용건수 현황"],
    "closing_report":  ["전사", "마감리포트", "2. [HPC, POS] HPC 일마감(브랜드)"],
}
```

## LOG REPORT 탐색 경로

해피앱 → 종합 → 누적추이 (유저 접속 추이) → 월별 보기 → 해당 월 선택
추출값: 순 로그인 회원수 (= 해피앱 월 로그인객수)

## VISUAL REPORT 탐색 경로

리포트 찾아보기 → 클라우드 → 프로모션 → 해피앱 GA 리포트 → 열기
→ 필터 설정 → 필터 지우기 → MAU 당월 위젯 → ··· → 데이터 내보내기 → 확인
추출값: 다운로드된 Excel A2 셀 (= 해피앱 MAU)

## 재시도 흐름

### 일별 (daily)
```
Task Scheduler: 08:30 기동, 30분마다 반복, 3시간 동안
    └─ state = success  → skip
    └─ should_give_up() → 실패 메일, exit
    └─ DataNotReadyError → 알림 메일, exit (다음 30분에 재시도)
    └─ 기타 오류 → state = failed, 실패 메일, exit
```

### 월별 (monthly) — 매월 1일
```
Task Scheduler: 매월 1일 08:30 기동, 30분마다 반복, 4일 동안
    └─ state = success  → skip
    └─ should_give_up() → 실패 메일, exit
    └─ DataNotReadyError → 알림 메일, exit (다음 30분에 재시도)
    └─ 기타 오류 → state = failed, 실패 메일, exit
```

## 주요 튜닝 포인트

### 1. OLAP 셀렉터 (`src/olap_scraper.py` 상단 `SEL_*` 상수)

실제 SASHBI DOM 구조 확인 방법:
```
HEADLESS=0, DRY_RUN=1 로 실행 → logs/debug/*.png 확인
브라우저 F12 → 요소 우클릭 → Copy selector
```

`REPORT_FRAME_SELECTOR`: SASHBI 가 iFrame 을 사용하는지 확인 후 None 또는 실제 선택자 설정.

### 2. Excel 컬럼명 (`src/excel_parser.py` 상단 상수)

```
python -m src.excel_parser downloads/<파일>.xlsx
```
헤더를 보고 `MEMBER_DATE_COL`, `CHANNEL_KEY_COL`, `CHANNEL_VALUE_COL`, `CLOSING_ROW_MAP` 수정.

### 3. 일마감 리포트 구조

유저 확인: `"0002. SPC전사(3사)"` 는 **컬럼 헤더** (브랜드가 열, 메트릭이 행).
`CLOSING_BRAND_LABEL` 로 컬럼 인덱스 탐색 → 해당 열 값 추출.
실제 Excel 이 반대 구조(브랜드가 행, 메트릭이 열)이면 `parse_closing_report` 로직 교체 필요.

### 4. LOG REPORT 셀렉터 (`src/log_report_scraper.py` 상단 `SEL_LR_*` 상수)

```
HEADLESS=0, DRY_RUN=1 로 실행 → logs/debug/lr_*.png 확인
브라우저 F12 → 요소 우클릭 → Copy selector
```

주요 튜닝 포인트:
- `SEL_LR_NAV_*` — 해피앱/종합/누적추이 메뉴 경로
- `SEL_LR_PERIOD_SETTINGS` / `SEL_LR_MONTHLY_VIEW` — 월별 보기 전환 UI
- `SEL_LR_MONTH_PICKER` — 월 선택 피커 (select or input)
- `SEL_LR_LOGIN_COUNT_ROW` — 순 로그인 회원수 행 레이블

### 5. VISUAL REPORT 셀렉터 (`src/visual_report_scraper.py` 상단 `SEL_VR_*` 상수)

```
HEADLESS=0, DRY_RUN=1 로 실행 → logs/debug/vr_*.png 확인
```

주요 튜닝 포인트:
- `SEL_VR_REPORT_BROWSE` / `SEL_VR_CLOUD_FOLDER` / `SEL_VR_PROMO_FOLDER` / `SEL_VR_HAPPYAPP_ITEM` — 폴더 탐색
- `SEL_VR_FILTER_CLEAR` — 필터 지우기 버튼
- `SEL_VR_MAU_WIDGET` / `SEL_VR_MAU_MORE_BTN` — MAU 당월 위젯 ··· 버튼
- `REPORT_LOAD_TIMEOUT` — 리포트 로딩 대기 시간 (기본 360초)

## 미완성 항목

| 파일 | 항목 | 비고 |
|---|---|---|
| `src/olap_scraper.py` | `SEL_*` 셀렉터 | HEADLESS=0 실행 후 튜닝 |
| `src/excel_parser.py` | 컬럼명 상수 | 다운로드 후 튜닝 |
| `src/log_report_scraper.py` | `SEL_LR_*` 셀렉터 | HEADLESS=0 실행 후 튜닝 |
| `src/visual_report_scraper.py` | `SEL_VR_*` 셀렉터 | HEADLESS=0 실행 후 튜닝 |

## 코딩 규칙

- 상수는 모듈 상단에 분리 (셀렉터, 컬럼명, 시트명)
- 로그: `[STEP]`, `[PARSE]`, `[SHEETS]`, `[STATE]`, `[DRY_RUN]`, `[DONE]` 접두어
- Playwright 실패 시 반드시 `session.snapshot()` 호출 후 raise
- DRY_RUN 체크는 실제 외부 쓰기 직전에만 (파싱·탐색은 DRY_RUN 과 무관)
- `DataNotReadyError` 는 재시도 가능한 일시적 상태 — 로그 레벨 WARNING, 상태 failed 로 전환 안 함

## 참고 프로젝트 (`reference/`)

| 디렉토리 | 참고 포인트 |
|---|---|
| `mkt_edm_rpa` | browser.py, config.py, logger.py, retry 패턴, Gmail 발송 |
| `hp_sett_rpa` | Playwright 로그인·다운로드, openpyxl 처리, with_retry 데코레이터 |
| `hp_mealal_rpa` | openpyxl 파싱, PyInstaller .exe 빌드, tkinter 알림 |
| `hc_rpa` | withRetry JS 패턴 (Node.js, 참고만) |
