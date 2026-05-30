# B2C 고객지표 관리 자동화 RPA

매일 08:30 SASHBI OLAP 시스템에서 3개 리포트를 자동 다운로드하고,
Google Sheets **"B2C사업본부 고객지표"** 의 2개 시트를 업데이트합니다.

---

## 업무 플로우

```
[매일 08:30 Task Scheduler 기동]
        │
        ▼
state/daily_YYYY-MM-DD.json 확인
  └─ 이미 완료? → 종료
  └─ 재시도 시간 초과(>3시간)? → 실패 메일 → 종료
        │
        ▼
SASHBI OLAP 로그인 (SSO)
        │
        ├─ 리포트 1: [HPC] 일별 회원관리지표    → Excel 다운로드
        ├─ 리포트 2: [HPC] 채널별 적립·사용건수  → Excel 다운로드
        └─ 리포트 3: [HPC, POS] HPC 일마감(브랜드) → Excel 다운로드
        │
        ▼
Excel 파싱 (openpyxl)
        │
        ▼
Google Sheets 업데이트
  ├─ HPC 실적 (일별) — 4개 지표
  └─ 전사 매장실적 (일별) — 11개 지표
        │
        ▼
완료 메일 발송 + state 기록

[데이터 미준비 시: 알림 메일 → 30분 후 Task Scheduler 재기동 → 최대 11:30까지]
```

---

## 디렉토리 구조

```
ai-RPA/
├── .env.example            환경변수 템플릿
├── .gitignore
├── requirements.txt
├── run_rpa.bat             Task Scheduler 등록용 진입점
├── run_rpa.ps1             PowerShell 래퍼 (로그 회전 포함)
├── build_windows.ps1       PyInstaller .exe 빌드
├── downloads/              OLAP Excel 임시 파일 (gitignore)
├── state/                  일별/월별 실행 상태 JSON (gitignore)
├── logs/                   실행 로그 + 디버그 스크린샷 (gitignore)
│   └── debug/              단계별 .png + .html (셀렉터 튜닝용)
├── reference/              참고 RPA 프로젝트 (읽기 전용)
└── src/
    ├── main.py             CLI 진입점
    ├── config.py           .env 로딩 및 설정 객체
    ├── logger.py           파일+콘솔 로깅
    ├── browser.py          Playwright 세션 헬퍼
    ├── state_manager.py    실행 상태 JSON 관리
    ├── olap_scraper.py     OLAP 자동화 (로그인·탐색·다운로드)
    ├── excel_parser.py     다운로드 Excel 파싱 (3개 리포트)
    ├── sheets_writer.py    Google Sheets 쓰기
    ├── notifier.py         Gmail 알림
    ├── daily_runner.py     일별 업데이트 오케스트레이션
    └── monthly_runner.py   월별 업데이트 (구현 예정)
```

---

## 설치

```powershell
# 1. 가상환경 생성 및 활성화
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. 패키지 설치
pip install -r requirements.txt

# 3. Playwright 브라우저 설치
python -m playwright install chromium
```

---

## 설정

`.env.example` 을 복사해 `.env` 를 만들고 값을 채웁니다.

```powershell
Copy-Item .env.example .env
notepad .env
```

| 항목 | 설명 |
|---|---|
| `OLAP_ID` | SASHBI SSO 아이디 |
| `OLAP_PW` | SASHBI SSO 비밀번호 |
| `SHEETS_SPREADSHEET_ID` | Google Sheets 문서 ID (기본값 내장) |
| `SHEETS_CREDENTIALS_PATH` | Google OAuth credentials.json 경로 |
| `SHEETS_TOKEN_PATH` | OAuth 토큰 저장 경로 (최초 인증 후 자동 생성) |
| `GMAIL_CREDENTIALS_PATH` | Gmail 발송용 credentials.json (Sheets 와 동일 파일 가능) |
| `REPORT_SENDER` | 발송자 Gmail 주소 |
| `REPORT_RECIPIENTS` | 수신자 주소 (쉼표 구분) |
| `HEADLESS` | `1` = 브라우저 숨김 (기본), `0` = 화면 표시 (튜닝 시) |
| `DRY_RUN` | `1` = Sheets 쓰기·메일 생략 (테스트용) |

---

## Google API 인증

최초 1회만 브라우저 인증이 필요합니다. 이후 `token.json` 이 자동 갱신됩니다.

```powershell
# credentials.json 을 프로젝트 루트에 복사한 뒤 실행
python -m src.main daily --date 2026-01-01   # 아무 날짜로 시범 실행
# → 브라우저 창에서 Google 계정 로그인 → token.json 자동 저장
```

`credentials.json` 은 `mkt_edm_rpa` 의 OAuth Desktop 자격증명과 동일 파일을 사용할 수 있습니다.
단, Scopes 에 `spreadsheets` + `gmail.send` 가 모두 포함되어야 합니다.

---

## 실행

### 수동 실행

```powershell
# 어제 날짜 일별 업데이트
python -m src.main daily

# 특정 날짜 재실행
python -m src.main daily --date 2026-05-29

# 이미 완료된 날짜 강제 재실행
python -m src.main daily --date 2026-05-29 --force

# 화면 표시 + 쓰기 없이 (셀렉터 튜닝 시)
# .env 에서 HEADLESS=0, DRY_RUN=1 로 설정 후 실행
python -m src.main daily
```

### Windows Task Scheduler 설정

| 항목 | 값 |
|---|---|
| 프로그램 | `C:\Users\S9_User\projects\ai-RPA\run_rpa.bat` |
| 인수 | `daily` |
| 시작 위치 | `C:\Users\S9_User\projects\ai-RPA` |
| 트리거 | 매일 08:30 |
| 반복 | 30분마다, 3시간 동안 |
| 실행 시간 초과 | 10분 |

> **핵심**: Task Scheduler 가 30분마다 재실행하면 `state/daily_*.json` 이 이미 `success` 인지 확인해 중복 실행을 방지합니다.

---

## 재시도 전략

### 일별 업데이트

```
08:30 → 09:00 → 09:30 → 10:00 → 10:30 → 11:00 → 11:30
  첫 시도부터 3시간 경과 시 포기 → 실패 메일 → 수동 확인
```

- **데이터 미준비** (`DataNotReadyError`): OLAP 에 어제 날짜 데이터가 없으면
  알림 메일만 발송하고 종료 → 다음 30분 재시도
- **포기 조건**: 첫 시도 시각 기준 3시간 경과 후에도 실패 → 최종 실패 메일
- **재실행 방지**: `state/daily_YYYY-MM-DD.json` 의 `status: "success"` 확인

### 월별 업데이트 (구현 예정)

```
1일 08:30 ~ 5일 18:00 사이 30분마다 재시도
5일 초과 → 에스컬레이션 메일
```

---

## 셀렉터 튜닝 가이드

SASHBI 사이트의 실제 DOM 구조를 확인해야 합니다.

**1단계: 화면 표시 모드로 실행**

```
# .env 수정
HEADLESS=0
DRY_RUN=1
```

```powershell
python -m src.main daily
```

**2단계: 스크린샷 확인**

`logs/debug/` 폴더에 단계별 `.png` + `.html` 이 저장됩니다.

```
01_login_page.png       ← 로그인 폼 셀렉터 확인
02_after_login.png      ← 메인 화면 트리 구조 확인
03_report_loaded_*.png  ← 리포트 패널 확인
04_date_filter_set.png  ← 날짜 필터 입력 확인
05_after_run.png        ← 조회 결과 확인
06_downloaded_*.png     ← 다운로드 후 화면 확인
```

**3단계: `src/olap_scraper.py` 상단 `SEL_*` 상수 수정**

```python
# 예시 — 실제 DOM 확인 후 교체
SEL_LOGIN_ID  = "input#userId"          # 로그인 ID 필드
SEL_LOGIN_PW  = "input#password"        # 로그인 PW 필드
SEL_LOGIN_BTN = "button:has-text('로그인')"
SEL_TREE_NODE = "span[title='{label}']" # 리포트 트리 노드
SEL_DATE_START = "input[name='startDate']"
SEL_RUN_BTN   = "button:has-text('실행')"
SEL_EXPORT_BTN   = "button[title='내보내기']"
SEL_EXPORT_EXCEL = "li:has-text('Excel')"
```

> 브라우저 개발자 도구(F12) → Elements → 요소 우클릭 → Copy selector

---

## Excel 파서 튜닝 가이드

OLAP Excel 의 실제 컬럼명을 확인합니다.

```powershell
# 다운로드된 파일 구조 확인
python -m src.excel_parser downloads/<파일명>.xlsx
```

출력 예시:
```
=== member_metrics_xxx.xlsx ===
헤더: ['일자', '신규가입회원수', '해피앱 로그인 회원수', '해피앱 DAU', '해피오더 DAU', ...]
행: (datetime(2026, 5, 29), 1234, 567890, ...)
```

**`src/excel_parser.py` 상단 상수 수정:**

| 상수 | 기본값 | 설명 |
|---|---|---|
| `MEMBER_DATE_COL` | `"일자"` | 리포트 1 날짜 컬럼명 |
| `CHANNEL_KEY_COL` | `"채널"` | 리포트 2 채널명 컬럼 |
| `CHANNEL_VALUE_COL` | `"제시건수"` | 리포트 2 건수 컬럼 |
| `CLOSING_ROW_MAP` | (딕셔너리) | 리포트 3 메트릭 행 레이블 매핑 |

---

## Google Sheets 컬럼 매핑

### HPC 실적 (일별)

| 열 | 지표 | 출처 |
|---|---|---|
| A | 월 (YYYYMM) | 자동 계산 |
| B | 날짜 | 자동 계산 |
| C | 요일 | 자동 계산 |
| D | 신규회원수 | 리포트 1: 신규가입회원수 |
| E | 해피앱 로그인수 | 리포트 1: 해피앱 로그인 회원수 |
| F | 해피앱 DAU | 리포트 1 |
| G | 해피오더 DAU | 리포트 1 |

### 전사 매장실적 (일별)

| 열 | 지표 | 출처 |
|---|---|---|
| A | 월 (YYYYMM) | 자동 계산 |
| B | 날짜 | 자동 계산 |
| C | POS 총매출액 | 리포트 3: 0002. SPC전사(3사) |
| D | POS 영수증건수 | 리포트 3 |
| E | POS 거래점포수 | 리포트 3 |
| F | HPC 매출액 | 리포트 3 |
| G | HPC 거래점포수 | 리포트 3 |
| H | HPC 총적립액 | 리포트 3 |
| I | HPC 적립건수 | 리포트 3 |
| J | 객단가 | 리포트 3 |
| K | HPC 총사용액 | 리포트 3 |
| L | HPC 사용건수 | 리포트 3 |
| N | APP 제시건수 | 리포트 2: HPCAPP |

---

## Windows .exe 빌드

비기술 담당자 PC 에 Python 없이 배포할 때 사용합니다.

```powershell
.\build_windows.ps1
```

결과: `dist\고객지표_RPA\` 폴더

배포 시 함께 복사할 파일:
```
dist\고객지표_RPA\
.env
credentials.json
token.json          ← 최초 인증 후 생성됨
```

---

## 주의 사항

- `.env` / `credentials.json` / `token.json` 은 **절대 git 에 커밋하지 마세요** (`.gitignore` 에 포함)
- SASHBI 가 내부망 전용이므로 VPN 없이는 실행되지 않습니다
- `DRY_RUN=1` 상태에서는 Sheets 쓰기와 메일 발송이 생략되며 로그에만 기록됩니다
- 셀렉터 튜닝 중에는 반드시 `DRY_RUN=1` 로 설정하세요

---

## 미완성 항목

| 항목 | 상태 |
|---|---|
| OLAP 셀렉터 (`src/olap_scraper.py` `SEL_*` 상수) | 튜닝 필요 |
| Excel 파서 컬럼명 (`src/excel_parser.py`) | 튜닝 필요 |
| 월별 업데이트 (`src/monthly_runner.py`) | 구현 예정 (LOG REPORT + VISUAL REPORT) |
| iFrame 여부 (`REPORT_FRAME_SELECTOR`) | 확인 필요 |
