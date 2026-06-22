# CLAUDE.md

This file provides guidance to Claude Code when working in this repository.

## Purpose

B2C 고객지표 관리 자동화 RPA. 매일 SASHBI OLAP 에서 3개 리포트를 다운로드하고
Google Sheets "B2C사업본부 고객지표" 의 2개 시트를 업데이트한다.
매월 1일에는 LOG REPORT / VISUAL REPORT 에서 월별 지표를 추가로 업데이트한다.

## 로그인 플로우 (확인 완료)

모든 내부 시스템은 **SPC Hub SSO** 를 통해 인증합니다.

```
1. SPC Hub 로그인 (hub.spc.co.kr) — SPCHUB_ID / SPCHUB_PW
2. 확인/동의 버튼 클릭 (있는 경우, 없으면 자동 건너뜀)
3. Hub 상단 nav "System Link" → "정보화시스템" 클릭 → sis.spc.co.kr/main/main.jsp
   (SEL_HUB_SYSTEM_LINK 호버 → SEL_HUB_SIS_LINK 클릭)
   ※ sis.spc.co.kr 은 Hub SSO 통해서만 접근 가능 — 직접 URL 접근 시 login.jsp 리다이렉트
   ※ 클릭 결과: 새 창 또는 현재 탭 이동 두 가지 모두 자동 처리
4. sis.spc.co.kr 메뉴에서 img 버튼 클릭 → 팝업 창(window.open + form POST)
   - OLAP       → 별도 로그인 없음 (SSO 릴레이)
   - LOG REPORT → 별도 로그인 없음 (SSO 릴레이)
   - VISUAL REPORT → 로그인 페이지 표시 시 "OLAP 계정으로 로그인" 클릭
```

- **SPC Hub**: `https://hub.spc.co.kr/ekp/view/login/userLogin`
- **SSO 게이트**: `https://hubsso.spc.co.kr:9443/dnsagent/gate/legacySSOGate.jsp?nurl=https://sis.spc.co.kr`
- **OLAP**: `https://dwweb.spc.co.kr:7980/SASHBI/main.jsp` (SSO, 평일 09:00–18:00만 접속 가능)
- **LOG REPORT**: `https://hplog.spc.co.kr:8000/datastory/home` (SSO)
- **VISUAL REPORT**: `https://va.spc.co.kr/SASReportViewer/` (SSO + OLAP 계정 로그인)
- **Sheets ID**: `1gIEbHzyfh4TG21M1etxWxh0u99ylGMsX4B_VGEu3Yxw`

## Runtime

- Python 3 (`python` 또는 `python3`)
- 가상환경: `.venv/` (없으면 시스템 Python)
- 브라우저: Playwright Chromium (`python -m playwright install chromium`)
- 외부 의존: `requirements.txt` 참고 (`xlrd<2.0` 포함 — HTML 위장 XLS 대응)

## Entry Point

```
# ── 최종 사용자 배포 ──────────────────────────────────────────────────
setup_user.bat                        # 원클릭 설치 (비개발자용)
                                      # Python 확인 → venv → 패키지 → Chromium
                                      # → SPC Hub 자격증명 입력 (비밀번호 마스킹)
                                      # → Google OAuth → Task Scheduler 등록
                                      # ※ credentials.json 이 없으면 실행 불가
                                      #   (관리자가 별도 전달, zip 에 미포함)

# ── 개발자용 ──────────────────────────────────────────────────────────
setup.bat                             # 최초 설치 + 설정 마법사 실행
python -m src.setup_gui               # 설정 마법사만 단독 실행

python -m src.main daily              # 어제 날짜 일별 업데이트
python -m src.main daily --date YYYY-MM-DD --force
python -m src.main monthly            # 이번 달 월별 업데이트 (매월 1일 실행)
python -m src.main monthly --year 2026 --month 5 --force
```

## 배포 패키지

`B2C_고객지표_RPA.zip` (프로젝트 루트) — 최종 사용자에게 전달하는 파일.

포함: `setup_user.bat`, `run_rpa.bat`, `python-3.12.9-amd64.exe`, `.env` (템플릿),
      `requirements_user.txt`, `설치_사용_가이드.md`, `src/*.py`

**미포함**: `credentials.json` — 보안상 관리자가 별도 전달. 사용자가 압축 해제 후
           같은 폴더에 복사해야 `setup_user.bat` 이 실행됨.

**주의**: zip 재생성 전 `.bat` 파일이 CRLF 줄바꿈인지 확인. LF 전용이면 `cmd.exe` `^` 줄 이음이 깨져
`'wright'은(는) 내부 또는 외부 명령...` 오류 발생. 확인: `python -c "raw=open('setup_user.bat','rb').read(); print('CRLF:',raw.count(b'\r\n'),'LF:',raw.count(b'\n')-raw.count(b'\r\n'))"` — LF 가 0이어야 함.

zip 재생성:
```powershell
# 프로젝트 루트에서 실행
python -c "
import zipfile, pathlib, shutil, os

stage = pathlib.Path('_pkg_stage')
shutil.rmtree(stage, ignore_errors=True)
(stage / 'src').mkdir(parents=True)

for f in ['setup_user.bat','run_rpa.bat','requirements_user.txt','설치_사용_가이드.md','python-3.12.9-amd64.exe']:
    shutil.copy(f, stage / f)
shutil.copy('.env_user_template', stage / '.env')

for py in pathlib.Path('src').glob('*.py'):
    shutil.copy(py, stage / 'src' / py.name)

with zipfile.ZipFile('B2C_고객지표_RPA.zip','w',zipfile.ZIP_DEFLATED,compresslevel=9) as zf:
    for f in sorted(stage.rglob('*')):
        if f.is_file(): zf.write(f, f.relative_to(stage))

shutil.rmtree(stage)
print('Done')
"
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
│                          screenshot timeout=10s (페이지 로딩 중 블로킹 방지).
├── state_manager.py       DailyState / MonthlyState — state/ 폴더에 JSON 저장.
│                          should_give_up() 로 재시도 시간 초과 판단.
├── hub_login.py           SPC Hub SSO 로그인 공통 모듈.
│                          login_to_hub() → _SIS_SSO_GATE URL 직접 접속
│                          → _open_via_menu() 팝업 캡처.
│                          navigate_to_olap / log_report / visual_report().
├── olap_scraper.py        OLAP 자동화의 핵심. SEL_* 상수로 셀렉터 분리.
│                          login() → hub_login 경유. scrape_report() 전체 플로우.
│                          DataNotReadyError — 날짜 미확인 시 raise.
│                          WRS 리포트: iFrame HTML 저장 방식 (cwMenuBarExport 대체).
├── log_report_scraper.py  LOG REPORT 자동화. SEL_LR_* 상수로 셀렉터 분리.
│                          scrape_login_count(config, year, month) → int.
│                          해피앱 > 종합 > 유저 접속 추이 > 월별 보기 탐색.
├── visual_report_scraper.py  VISUAL REPORT 자동화. SEL_VR_* 상수로 셀렉터 분리.
│                          scrape_mau_excel(config, year, month) → Path.
│                          리포트 찾아보기 > 클라우드 > 프로모션 > 해피앱 GA 리포트.
│                          필터 지우기 → MAU 당월 Excel 다운로드. 로딩 최대 360초.
├── excel_parser.py        4개 파서. 상단 상수로 컬럼명 분리.
│                          HTML 위장 XLS 자동 감지 (OLE magic bytes 확인).
│                          HPCAPP wide/tall 포맷 자동 감지.
│                          SAS 결측값 "." → None 처리.
│                          CLI: python -m src.excel_parser <파일> 로 헤더 확인.
├── sheets_writer.py       gspread 기반. find_or_create_row() 로 날짜 행 탐색/생성.
│                          _build_batch() 에 sheet_title 포함 → 올바른 시트에 쓰기.
│                          write_hpc_daily / write_store_daily / write_hpc_monthly.
├── notifier.py            Gmail API. notify_success / notify_failure /
│                          notify_data_not_ready 3가지.
├── daily_runner.py        run_daily() — 다운로드 → 파싱 → 쓰기 → 알림 순서 조율.
│                          DRY_RUN=1 시 open_spreadsheet 호출 자체를 생략.
│                          DataNotReadyError 는 실패가 아닌 재시도 대기로 처리.
├── monthly_runner.py      run_monthly() — LOG REPORT + VISUAL REPORT → Sheets 업데이트.
│                          DataNotReadyError 재시도 처리. 매월 1일 실행.
└── setup_gui.py           비개발자용 tkinter 설정 마법사. python -m src.setup_gui.
                           SPC 로그인·Google 인증·이메일·Task Scheduler 등록.
                           setup.bat 에서 자동으로 실행됨.
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

## OLAP 리포트 경로 (실제 트리 구조 확인 완료)

```python
REPORT_PATHS = {
    "member_metrics":  ["정형데이터", "회원분석", "[HPC] 일별 회원관리지표"],
    "channel_metrics": ["정형데이터", "회원분석", "[HPC] 채널별 적립, 사용건수 현황"],
    "closing_report":  ["전사", "마감리포트", "2. [HPC, POS] HPC 일마감(브랜드)"],
}
```

closing_report 는 SAS WRS 리포트 — Submit 버튼 없음, 자동 렌더링.
iFrame HTML (`frmSASHBIAA00012A`) 을 직접 파싱해 데이터 추출.

## Excel 파일 형식 (실제 확인 완료)

| 리포트 | 다운로드 형식 | 파서 |
|---|---|---|
| member_metrics | `.xls` (HTML 위장) | `_load_html_table_sheet()` |
| channel_metrics | `.xlsx` | `openpyxl` |
| closing_report | `.html` (WRS iFrame 저장) | `_parse_closing_report_html()` |

- OLAP 이 내보내는 `.xls` 는 실제로는 HTML (meta charset=utf-8).
- `channel_metrics` 는 Wide 포맷: HPCAPP 이 컬럼 헤더, 날짜·브랜드가 행 값.
- SAS 결측값 `"."` 은 `None` 으로 변환.

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

### 0. Hub 로그인 셀렉터 (`src/hub_login.py` 상단 `SEL_HUB_*` / `SEL_MENU_*`)

모든 시스템 공통 진입점.

```
HEADLESS=0, DRY_RUN=1 로 실행 → logs/debug/hub_*.png 확인
```

주요 튜닝 포인트:
- `SEL_HUB_ID/PW/BTN` — Hub 로그인 폼 (확인 완료)
- `SEL_HUB_CONFIRM` — 로그인 후 확인/동의 버튼 (확인 완료)
- `SEL_HUB_SYSTEM_LINK` — Hub nav "System Link" 드롭다운 (튜닝 필요 시 hub_04_before_sis_click.png)
- `SEL_HUB_SIS_LINK` — "정보화시스템" 링크 (튜닝 필요 시 hub_04_before_sis_click.png)
- `SEL_MENU_OLAP/LOG_REPORT/VISUAL_REPORT` — sis.spc.co.kr img 버튼 (확인 완료)
- `SEL_VR_OLAP_LOGIN` — VISUAL REPORT 내 "OLAP 계정으로 로그인"

### 1. OLAP 셀렉터 (`src/olap_scraper.py` 상단 `SEL_*` 상수)

```
HEADLESS=0, DRY_RUN=1 로 실행 → logs/debug/*.png 확인
```

- `SEL_TREE_NODE` — 탭/아코디언/리포트 링크 (확인 완료)
- `SEL_DATE_START/END` — HBI 리포트 날짜 input (확인 완료)
- `SEL_RUN_BTN` — HBI Submit 버튼 (확인 완료)
- `SEL_EXPORT_BTN` — HBI Excel 버튼 (확인 완료)
- `REPORT_FRAME_SELECTOR` — 활성 탭 내 frmSASHBI iFrame (확인 완료)

### 2. Excel 파서 (`src/excel_parser.py` 상단 상수)

```
python -m src.excel_parser downloads/<파일>
```

헤더 확인 후 `MEMBER_COL_MAP`, `CHANNEL_TARGET_LABEL`, `CLOSING_ROW_MAP` 수정.

### 3. LOG REPORT 셀렉터 (`src/log_report_scraper.py` 상단 `SEL_LR_*` 상수)

```
HEADLESS=0, DRY_RUN=1 로 실행 → logs/debug/lr_*.png 확인
```

### 4. VISUAL REPORT 셀렉터 (`src/visual_report_scraper.py` 상단 `SEL_VR_*` 상수)

```
HEADLESS=0, DRY_RUN=1 로 실행 → logs/debug/vr_*.png 확인
```

## 코딩 규칙

- 상수는 모듈 상단에 분리 (셀렉터, 컬럼명, 시트명)
- 로그: `[STEP]`, `[PARSE]`, `[SHEETS]`, `[STATE]`, `[DRY_RUN]`, `[DONE]` 접두어
- Playwright 실패 시 반드시 `session.snapshot()` 호출 후 raise
- DRY_RUN 체크는 `daily_runner._write_to_sheets` 진입부에서 일괄 처리
  (open_spreadsheet 자체를 건너뜀 — credentials.json 없어도 DRY_RUN 가능)
- `DataNotReadyError` 는 재시도 가능한 일시적 상태 — 로그 레벨 WARNING, 상태 failed 로 전환 안 함

## 참고 프로젝트 (`reference/`)

| 디렉토리 | 참고 포인트 |
|---|---|
| `mkt_edm_rpa` | browser.py, config.py, logger.py, retry 패턴, Gmail 발송 |
| `hp_sett_rpa` | Playwright 로그인·다운로드, openpyxl 처리, credentials.json 원본 |
| `hp_mealal_rpa` | openpyxl 파싱, PyInstaller .exe 빌드, tkinter 알림 |
| `hc_rpa` | withRetry JS 패턴 (Node.js, 참고만) |
