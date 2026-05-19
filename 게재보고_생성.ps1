# =============================================================
# 신한카드 검색광고 게재보고 PPT 자동 생성 Launcher
# - 더블클릭 또는 PowerShell 에서 실행
# - PC / MO 링크를 붙여넣으면 캡처 → PPT 생성
# =============================================================

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "  신한카드 멤버십 영업팀 - 검색광고 게재보고 자동 생성" -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "예시:" -ForegroundColor Gray
Write-Host "  PC URL:  https://card-search.naver.com/item?cardAdId=3257&query=신한카드%20Deep%20Once&cx=..." -ForegroundColor DarkGray
Write-Host "  MO URL:  https://m-card-search.naver.com/item?cardAdId=3257&query=신한카드%20Deep%20Once&cx=..." -ForegroundColor DarkGray
Write-Host ""

$pc = Read-Host "1) PC 게재지면 URL 을 붙여넣고 Enter"
if ([string]::IsNullOrWhiteSpace($pc)) { Write-Host "PC URL이 비었습니다. 종료." -ForegroundColor Red; exit 1 }

$mo = Read-Host "2) MO 게재지면 URL 을 붙여넣고 Enter"
if ([string]::IsNullOrWhiteSpace($mo)) { Write-Host "MO URL이 비었습니다. 종료." -ForegroundColor Red; exit 1 }

$card = Read-Host "3) 카드명 (예: 신한카드 Deep Once)"
if ([string]::IsNullOrWhiteSpace($card)) { $card = "신한카드" }

# 출력 파일 경로
$desktop = [Environment]::GetFolderPath("Desktop")
$outName = "(핀플로우) 신한카드 멤버십 영업팀 신용카드 검색광고 게재보고_$card.pptx"
$outPath = Join-Path $desktop $outName

# 같은 이름 파일이 있으면 _1, _2 ... 붙이기
$i = 1
while (Test-Path $outPath) {
  $outPath = Join-Path $desktop "(핀플로우) 신한카드 멤버십 영업팀 신용카드 검색광고 게재보고_${card}_$i.pptx"
  $i++
}

Write-Host ""
Write-Host "처리 시작..." -ForegroundColor Green
Write-Host "  카드: $card"
Write-Host "  출력: $outPath"
Write-Host ""

$script = Join-Path $here "generate_report.py"
python $script --pc "$pc" --mo "$mo" --card "$card" --out "$outPath"

if (Test-Path $outPath) {
  Write-Host ""
  Write-Host "완료!" -ForegroundColor Green
  Write-Host "  생성된 파일: $outPath" -ForegroundColor Green
  Write-Host ""
  $open = Read-Host "PPT 파일을 지금 열까요? (y/n)"
  if ($open -eq 'y' -or $open -eq 'Y') { Start-Process $outPath }
} else {
  Write-Host "PPT 파일이 생성되지 않았습니다." -ForegroundColor Red
}

Write-Host ""
Read-Host "Enter 를 눌러 종료"
