# BckPPT 실행 스크립트
# 소스 변경 시에만 빌드 후 실행, 변경 없으면 바로 실행
#
# 사용법:
#   run.ps1              # 변경 감지 → 빌드 + 실행
#   run.ps1 --build-only # 변경 감지 → 빌드만 (VS Code preLaunchTask용)

param([switch]$BuildOnly)

$ProjectDir = $PSScriptRoot
$ExePath = "$ProjectDir\bin\x64\Debug\net8.0-windows\BckPPT.exe"

# ── 소스 변경 감지 ────────────────────────────────────────────────
$SourceFiles = Get-ChildItem -Path $ProjectDir -Recurse -Include "*.cs","*.xaml","*.csproj" `
    | Where-Object { $_.FullName -notmatch "\\(bin|obj)\\" }

$NeedsBuild = $false

if (-not (Test-Path $ExePath)) {
    Write-Host "exe 없음 → 빌드 필요" -ForegroundColor Yellow
    $NeedsBuild = $true
} else {
    $ExeTime = (Get-Item $ExePath).LastWriteTime
    $NewerSource = $SourceFiles | Where-Object { $_.LastWriteTime -gt $ExeTime }

    if ($NewerSource) {
        Write-Host "변경된 파일 $($NewerSource.Count)개 감지 → 빌드" -ForegroundColor Cyan
        $NewerSource | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor DarkCyan }
        $NeedsBuild = $true
    } else {
        Write-Host "소스 변경 없음 → 바로 실행" -ForegroundColor Green
    }
}

# ── 빌드 ─────────────────────────────────────────────────────────
if ($NeedsBuild) {
    dotnet build "$ProjectDir\BckPPT.csproj" -c Debug -p:Platform=x64 --nologo -v:minimal
    if ($LASTEXITCODE -ne 0) {
        Write-Host "빌드 실패. 실행 중인 이전 개발 빌드가 파일을 점유했다면 창을 닫고 다시 시도하세요." -ForegroundColor Red
        exit 1
    }
}

# ── 실행 ─────────────────────────────────────────────────────────
if (-not $BuildOnly) {
    Start-Process $ExePath -WorkingDirectory $ProjectDir
}
