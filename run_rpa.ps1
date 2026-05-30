# 고객지표 RPA — PowerShell 래퍼
# Task Scheduler: run_rpa.bat 를 직접 등록하는 것을 권장.
# 이 스크립트는 로그 회전 + 알림 등 추가 작업이 필요할 때 사용.

param(
    [string]$Mode = "daily",
    [string]$Date = "",
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptDir

# 가상환경 활성화
if (Test-Path ".venv\Scripts\Activate.ps1") {
    . .\.venv\Scripts\Activate.ps1
}

# 인수 조합
$args_list = @($Mode)
if ($Date)  { $args_list += "--date"; $args_list += $Date }
if ($Force) { $args_list += "--force" }

Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] RPA 실행: $($args_list -join ' ')"
python -m src.main @args_list
$exit = $LASTEXITCODE

# 180일 이상 된 로그 파일 자동 삭제
Get-ChildItem -Path "logs" -Filter "*.log" -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-180) } |
    Remove-Item -Force

exit $exit
