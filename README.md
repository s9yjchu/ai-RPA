# B2C 고객지표 관리 자동화 RPA

매일 08:30 SASHBI OLAP 시스템에서 3개 리포트를 자동 다운로드하고,
Google Sheets **"B2C사업본부 고객지표"** 의 시트를 업데이트합니다.
매월 1일에는 LOG REPORT / VISUAL REPORT 에서 월별 지표를 추가로 업데이트합니다.

> **운영 담당자용 사용 가이드** → [`사용법.md`](사용법.md)

---

## 업무 플로우

### 일별 (매일 08:30)

```
[Task Scheduler 기동]
        │
        ▼
state/daily_YYYY-MM-DD.json 확인
  └─ 이미 완료? → 종료
  └─ 재시도 시간 초과(>3시간)? → 실패 메일 → 종료
        │
        ▼
SPC Hub SSO 로그인 → sis.spc.co.kr → OLAP 팝업
        │
        ├─ 리포트 1: [HPC] 일별 회원관리지표      → Excel 다운로드
        ├─ 리포트 2: [HPC] 채널별 적립·사용건수    → Excel 다운로드
        └─ 리포트 3: [HPC, POS] HPC 일마감(브랜드) → HTML 저장
        │
        ▼
파싱 (openpyxl / HTML 파서)
        │
        ▼
Google Sheets 업데이트
  ├─ HPC 실적 (일별)     — D~G열 4개 지표
  └─ 전사 매장실적 (일별) — C~N열 11개 지표
        │
        ▼
완료 메일 발송 + state 기록

[데이터 미준비 시: 알림 메일 → 30분 후 재시도 → 최대 11:30까지]
```

### 월별 (매월 1일)

```
[Task Scheduler 기동]
        │
        ▼
state/monthly_YYYY-MM.json 확인
        │
        ├─ LOG REPORT (hplog.spc.co.kr)
        │    해피앱 > 종합 > 유저 접속 추이 > 월별 보기
        │    → 순 로그인 회원수 추출
        │
        └─ VISUAL REPORT (va.spc.co.kr)
             리포트 찾아보기 > 클라우드 > 프로모션 > 해피앱 GA 리포트
             → 필터 지우기 → MAU 당월 Excel 다운로드 → A2 셀 파싱
        │
        ▼
Google Sheets 업데이트
  └─ HPC 실적 (월별) — O열(로그인객수), P열(MAU)
        │
        ▼
완료 메일 발송 + state 기록
```

---

## 디렉토리 구조

```
ai-RPA/
├── .env.example               환경변수 템플릿
├── .gitignore
├── requirements.txt
├── 사용법.md                  운영 담당자용 사용 가이드
├── run_rpa.bat                Task Scheduler 등록용 진입점
├── run_rpa.ps1                PowerShell 래퍼 (로그 회전 포함)
├── build_windows.ps1          PyInstaller .exe 빌드
├── downloads/                 OLAP Excel 임시 파일 (gitignore)
├── state/                     일별/월별 실행 상태 JSON (gitignore)
├── logs/                      실행 로그 + 디버그 스크린샷 (gitignore)
│   └── debug/                 단계별 .png + .html (셀렉터 튜닝용)
├── reference/                 참고 RPA 프로젝트 (읽기 전용)
└── src/
    ├── main.py                CLI 진입점
    ├── config.py              .env 로딩 및 설정 객체
    ├── logger.py              파일+콘솔 로깅 (KST)
    ├── browser.py             Playwright 세션 헬퍼
    ├── state_manager.py       실행 상태 JSON 관리
    ├── hub_login.py           SPC Hub SSO 로그인 공통 모듈
    ├── olap_scraper.py        OLAP 자동화 (로그인·탐색·다운로드)
    ├── log_report_scraper.py  LOG REPORT 자동화 (월 로그인객수)
    ├── visual_report_scraper.py  VISUAL REPORT 자동화 (MAU Excel)
    ├── excel_parser.py        다운로드 Excel/HTML 파싱 (4개 파서)
    ├── sheets_writer.py       Google Sheets 쓰기 (일별·월별)
    ├── notifier.py            Gmail 알림
    ├── daily_runner.py        일별 업데이트 오케스트레이션
    ├── monthly_runner.py      월별 업데이트 오케스트레이션
    └── setup_gui.py           비개발자용 tkinter 설정 마법사
```

---

## 설치 및 초기 설정

### 일반 사용자 (비개발자)

1. **`setup.bat` 더블클릭**
   - 패키지 및 브라우저 자동 설치 (처음 실행 시 2~3분 소요)
   - 설정 마법사 창이 열립니다

2. **설정 마법사에서 4가지 항목 설정**

   | 항목 | 입력 내용 |
   |---|---|
   | SPC 로그인 정보 | SPC Hub 아이디·비밀번호 (`SPCHUB_ID` / `SPCHUB_PW`) |
   | Google 인증 | `credentials.json` 파일 선택 후 Google 로그인 |
   | 이메일 알림 설정 | 발송자·수신자 Gmail 주소 |
   | 자동 실행 등록 | 버튼 클릭 → Task Scheduler 자동 등록 |

3. **완료** — 이후 매일 08:30에 자동으로 실행됩니다

> **credentials.json** 은 `reference/hp_sett_rpa/` 폴더에 있는 파일을 그대로 사용하거나,
> IT 담당자에게 요청하세요.

---

<details>
<summary>개발자용 상세 설정</summary>

```powershell
# 수동 설치
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m playwright install chromium

# .env 설정
Copy-Item .env.example .env
notepad .env
```

| 환경변수 | 설명 |
|---|---|
| `SPCHUB_ID` / `SPCHUB_PW` | SPC Hub SSO 로그인 (모든 시스템 공통) |
| `OLAP_ID` / `OLAP_PW` | VISUAL REPORT 전용 (SSO 실패 시 fallback, 보통 불필요) |
| `SHEETS_SPREADSHEET_ID` | Google Sheets ID (기본값 내장) |
| `SHEETS_CREDENTIALS_PATH` | credentials.json 경로 |
| `GMAIL_CREDENTIALS_PATH` | Gmail credentials.json 경로 (Sheets와 동일 파일 가능) |
| `REPORT_SENDER` / `REPORT_RECIPIENTS` | 이메일 발송·수신 주소 |
| `HEADLESS` | `0` = 브라우저 표시 (셀렉터 튜닝 시) |
| `DRY_RUN` | `1` = Sheets 쓰기·메일 발송 생략, credentials.json 불필요 |

```powershell
# 수동 실행
python -m src.main daily
python -m src.main daily --date 2026-06-01 --force
python -m src.main monthly --year 2026 --month 5 --force
```

</details>

---

## 실행

### 수동 실행

`run_rpa.bat` 을 더블클릭하거나 설정 마법사의 **[지금 실행]** 버튼을 클릭합니다.

### 자동 실행 (Task Scheduler)

`setup.bat` → 설정 마법사 → **[자동 실행 등록]** 버튼으로 자동 등록됩니다.

| 스케줄 | 실행 조건 |
|---|---|
| 일별 | 매일 08:30, 30분 간격, 3시간 동안 |
| 월별 | 매월 1일 08:30, 30분 간격, 4일 동안 |

---

## 재시도 전략

### 일별 업데이트

```
08:30 → 09:00 → 09:30 → 10:00 → 10:30 → 11:00 → 11:30
  첫 시도부터 3시간 경과 시 포기 → 실패 메일 → 수동 확인
```

- **데이터 미준비** (`DataNotReadyError`): OLAP 에 어제 날짜 데이터가 없으면
  알림 메일만 발송하고 종료 → 다음 30분 재시도
- **포기 조건**: 첫 시도 시각 기준 3시간 경과 후에도 실패
- **재실행 방지**: `state/daily_YYYY-MM-DD.json` 의 `status: "success"` 확인

### 월별 업데이트

```
1일 08:30 → 30분 간격 → 최대 4일까지 재시도
  포기 조건: 첫 시도 기준 4일 경과
```

- **데이터 미준비** (`DataNotReadyError`): 알림 메일 발송, `failed` 전환 없음 → 재시도 유지
- **재실행 방지**: `state/monthly_YYYY-MM.json` 의 `status: "success"` 확인

---

## Google Sheets 컬럼 매핑

### HPC 실적 (월별)

| 열 | 지표 | 출처 |
|---|---|---|
| O (15) | 해피앱 월 로그인객수 | LOG REPORT: 순 로그인 회원수 |
| P (16) | 해피앱 MAU | VISUAL REPORT: MAU 당월 Excel A2 |

> 행 키: col A (YYYYMM 형식). 행은 staff 가 수동 생성 — RPA 는 기존 행만 덮어씀.

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
| N | APP 제시건수 | 리포트 2: HPCAPP 열 값 |

---

## 로그인 플로우 (확인 완료)

모든 내부 시스템은 **SPC Hub SSO** 를 통해 인증합니다.

```
1. SPC Hub 로그인 (hub.spc.co.kr) — SPCHUB_ID / SPCHUB_PW
2. 확인/동의 버튼 클릭 (있는 경우)
3. SSO 게이트 직접 접속 → sis.spc.co.kr 메뉴 페이지 도달
4. 메뉴 페이지에서 대상 시스템 버튼 클릭 (팝업 열림)
   - OLAP       → 별도 로그인 없음 (SSO 릴레이)
   - LOG REPORT → 별도 로그인 없음 (SSO 릴레이)
   - VISUAL REPORT → 로그인 페이지 표시 시 "OLAP 계정으로 로그인" 클릭
```

> OLAP 서버 (`dwweb.spc.co.kr:7980`) 는 **평일 업무시간(09:00–18:00 KST) 에만 접속 가능**합니다.

---

## 셀렉터 튜닝 가이드

셀렉터가 맞지 않을 경우 `HEADLESS=0, DRY_RUN=1` 로 실행해 `logs/debug/` 스크린샷을 확인합니다.

### Hub 로그인 / SIS 메뉴 (`src/hub_login.py` — `SEL_HUB_*`, `SEL_MENU_*`)

```powershell
# .env: HEADLESS=0, DRY_RUN=1
python -m src.main daily
# logs/debug/hub_*.png 확인
```

| 상수 | 역할 |
|---|---|
| `SEL_HUB_ID/PW/BTN` | Hub 로그인 폼 |
| `SEL_HUB_CONFIRM` | 로그인 후 확인/동의 버튼 |
| `SEL_MENU_OLAP/LOG_REPORT/VISUAL_REPORT` | sis.spc.co.kr 메뉴 버튼 |
| `SEL_VR_OLAP_LOGIN` | VISUAL REPORT "OLAP 계정으로 로그인" |

### OLAP 트리/필터/다운로드 (`src/olap_scraper.py` — `SEL_*`)

| 상수 | 역할 |
|---|---|
| `SEL_TREE_NODE` | 좌측 트리 탐색 (탭·아코디언·리포트 링크) |
| `SEL_DATE_START/END` | 날짜 필터 input |
| `SEL_RUN_BTN` | 조회 실행 버튼 |
| `SEL_EXPORT_BTN` | Excel 다운로드 버튼 (HBI 방식) |
| `REPORT_FRAME_SELECTOR` | 리포트 콘텐츠 iFrame 선택자 |

### LOG REPORT (`src/log_report_scraper.py` — `SEL_LR_*`)

```powershell
python -m src.main monthly --year 2026 --month 5 --force
# logs/debug/lr_*.png 확인
```

### VISUAL REPORT (`src/visual_report_scraper.py` — `SEL_VR_*`)

```powershell
python -m src.main monthly --year 2026 --month 5 --force
# logs/debug/vr_*.png 확인
```

---

## Excel 파서 튜닝 가이드

OLAP 이 내보내는 파일은 `.xls` 확장자의 HTML 파일입니다 (openpyxl 미지원 형식).
파서가 자동으로 감지하여 처리합니다.

```powershell
# 헤더 및 샘플 행 확인
python -m src.excel_parser downloads/<파일명>
```

`src/excel_parser.py` 상단 상수 수정이 필요한 경우:

| 상수 | 기본값 | 설명 |
|---|---|---|
| `MEMBER_DATE_COL` | `"일자"` | 리포트 1 날짜 컬럼명 |
| `MEMBER_COL_MAP` | (딕셔너리) | 리포트 1 컬럼명 → Sheets 필드명 매핑 |
| `CHANNEL_TARGET_LABEL` | `"HPCAPP"` | 리포트 2 채널 식별자 (헤더 또는 행 값) |
| `CLOSING_BRAND_LABEL` | `"0002. SPC전사(3사)"` | 리포트 3 브랜드 컬럼 헤더 |
| `CLOSING_ROW_MAP` | (딕셔너리) | 리포트 3 메트릭 행 레이블 → Sheets 필드명 매핑 |

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

- `.env` / `credentials.json` / `token.json` 은 **절대 git 에 커밋하지 마세요** (`.gitignore` 포함)
- 모든 시스템이 내부망 전용이므로 VPN 없이는 실행되지 않습니다
- OLAP 서버는 평일 업무시간(09:00–18:00 KST)에만 접속 가능합니다
- `DRY_RUN=1` 상태에서는 Sheets 쓰기와 메일 발송이 완전히 생략되며, `credentials.json` 없이도 실행됩니다
- VISUAL REPORT 리포트 로딩에 최대 5~6분 소요됩니다 (정상)
