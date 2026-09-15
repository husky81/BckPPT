# BckPPT를 CellBox 플러그인 패키지로 묶는다.
#
#   scripts\package-plugin.ps1                 # publish → plugin.json → zip → sha256 → latest.json
#   scripts\package-plugin.ps1 -InstallLocal   # 위 + %LOCALAPPDATA%\CellBox\Plugins\BckPPT 에 전개
#   scripts\package-plugin.ps1 -SkipBuild      # 기존 publish 결과 재사용 (빠른 재패키징)
#
# 산출물 (artifacts/):
#   bckppt-<version>.zip          플러그인 폴더 통째 (plugin.json + self-contained 런타임)
#   bckppt-<version>.zip.sha256
#   latest.json                   릴리스 버킷의 plugins/bckppt/latest.json 로 올릴 매니페스트
#
# 참고: docs/plans/cellbox-tray-plugin.md

param(
    [string]$Configuration = "Release",
    [string]$BaseUrl = "https://releases.cellbox.ai/plugins/bckppt",
    [string]$ReleaseNotes = "",
    [switch]$InstallLocal,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$projectRoot   = Split-Path -Parent $PSScriptRoot
$projectPath   = Join-Path $projectRoot "BckPPT.csproj"
$publishPath   = Join-Path $projectRoot "bin\publish\win-x64"
$artifactsPath = Join-Path $projectRoot "artifacts"
$version       = ([xml](Get-Content -Raw -LiteralPath $projectPath)).Project.PropertyGroup.Version
if (-not $version) { throw "BckPPT.csproj에서 <Version>을 읽지 못했습니다." }

function Clear-StaleOutputs {
    if (Test-Path -LiteralPath $publishPath) { Remove-Item -LiteralPath $publishPath -Recurse -Force }
    $objRoot = Join-Path $projectRoot "obj"
    if (Test-Path -LiteralPath $objRoot) {
        Get-ChildItem -LiteralPath $objRoot -Recurse -Directory -Filter "R2R" -ErrorAction SilentlyContinue |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force }
    }
}

function Write-JsonFile([string]$Path, $Object) {
    $json = $Object | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($Path, $json, (New-Object Text.UTF8Encoding($false)))
}

# ── 1. publish ───────────────────────────────────────────────────────────────
if (-not $SkipBuild) {
    Write-Host "[1/5] publish (self-contained win-x64)..." -ForegroundColor Cyan
    # publish는 증분 복사라 이전 버전의 DLL이 남는다. 실제로 stale한 Microsoft.Data.Sqlite가
    # 패키지에 섞여 들어가 실행 즉시 FileNotFoundException이 났다. 매번 비우고 시작한다.
    # ReadyToRun 중간 산출물(obj\...\R2R)도 같은 함정이라 함께 지운다.
    Clear-StaleOutputs
    dotnet publish $projectPath -c $Configuration -r win-x64 --self-contained true `
        -p:PublishProfile=win-x64 --nologo
    if ($LASTEXITCODE -ne 0) { throw "dotnet publish 실패" }
} else {
    Write-Host "[1/5] publish 생략 (-SkipBuild)" -ForegroundColor DarkGray
}

$exePath = Join-Path $publishPath "BckPPT.exe"
if (-not (Test-Path -LiteralPath $exePath)) { throw "게시 결과에서 BckPPT.exe를 찾을 수 없습니다: $publishPath" }

# ── 2. 아이콘 + plugin.json ──────────────────────────────────────────────────
Write-Host "[2/5] plugin.json 생성..." -ForegroundColor Cyan

# icon.ico는 exe에 박혀 있지만 트레이 메뉴가 파일로 읽으므로 폴더에도 둔다.
$assetsOut = Join-Path $publishPath "Assets"
New-Item -ItemType Directory -Path $assetsOut -Force | Out-Null
foreach ($asset in @("icon.ico", "icon.png")) {
    $source = Join-Path $projectRoot "Assets\$asset"
    if (Test-Path -LiteralPath $source) { Copy-Item -LiteralPath $source -Destination $assetsOut -Force }
    else { Write-Warning "Assets\$asset 없음 — 건너뜁니다." }
}

$manifest = [ordered]@{
    id                = "bckppt"
    name              = "BckPPT"
    version           = $version
    description       = "PPT 텍스트 분석·치환 도구"
    executable        = "BckPPT.exe"
    arguments         = ""
    icon              = "Assets/icon.ico"
    processName       = "BckPPT"
    updateManifestUrl = "$($BaseUrl.TrimEnd('/'))/latest.json"
    dataDirs          = @("%LOCALAPPDATA%\BckSoft\BckPPT")
    minCellBoxVersion = "1.0.0"
}
# JSON은 BOM 없이 쓴다. 스토어 페이지의 JSON.parse가 BOM을 만나면 실패한다.
Write-JsonFile (Join-Path $publishPath "plugin.json") $manifest

# ── 3. zip ───────────────────────────────────────────────────────────────────
Write-Host "[3/5] zip 압축..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $artifactsPath -Force | Out-Null

$zipName = "bckppt-$version.zip"
$zipPath = Join-Path $artifactsPath $zipName
$resolvedArtifacts = [IO.Path]::GetFullPath($artifactsPath).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
if (-not [IO.Path]::GetFullPath($zipPath).StartsWith($resolvedArtifacts, [StringComparison]::OrdinalIgnoreCase)) {
    throw "패키지 경로가 artifacts 폴더 밖을 가리킵니다."
}
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path (Join-Path $publishPath "*") -DestinationPath $zipPath -CompressionLevel Optimal

# ── 4. sha256 + latest.json ──────────────────────────────────────────────────
Write-Host "[4/5] SHA-256 산출..." -ForegroundColor Cyan
$hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath "$zipPath.sha256" -Value "$hash  $zipName" -Encoding ascii

# BckPPT UpdateService가 그대로 읽는 형식 {version, packageUrl, sha256, releaseNotes}
$latest = [ordered]@{
    version      = $version
    packageUrl   = "$($BaseUrl.TrimEnd('/'))/$zipName"
    sha256       = $hash
    releaseNotes = $ReleaseNotes
}
$latestPath = Join-Path $artifactsPath "latest.json"
Write-JsonFile $latestPath $latest

# ── 5. 로컬 설치 (선택) ──────────────────────────────────────────────────────
if ($InstallLocal) {
    Write-Host "[5/5] 로컬 플러그인 폴더에 전개..." -ForegroundColor Cyan
    $target = Join-Path $env:LOCALAPPDATA "CellBox\Plugins\BckPPT"

    # 실행 중이면 파일이 잠긴다. plugin.json의 processName과 같은 규칙으로 멈춘다.
    Get-Process -Name $manifest.processName -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "  실행 중인 $($_.ProcessName) (PID $($_.Id)) 종료" -ForegroundColor Yellow
        try { $null = $_.CloseMainWindow(); $null = $_.WaitForExit(5000) } catch { }
        try { if (-not $_.HasExited) { $_.Kill(); $null = $_.WaitForExit(5000) } } catch { }
    }

    # 폴더 교체는 이름 변경 → 전개 → 되돌리기. 중간에 실패해도 이전 버전이 남는다.
    $backup = "$target.old"
    if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Recurse -Force }
    if (Test-Path -LiteralPath $target) { Rename-Item -LiteralPath $target -NewName (Split-Path -Leaf $backup) }
    try {
        New-Item -ItemType Directory -Path $target -Force | Out-Null
        Expand-Archive -LiteralPath $zipPath -DestinationPath $target -Force
        if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Recurse -Force }
    } catch {
        if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
        if (Test-Path -LiteralPath $backup) { Rename-Item -LiteralPath $backup -NewName (Split-Path -Leaf $target) }
        throw
    }
    Write-Host "  설치: $target" -ForegroundColor Green
} else {
    Write-Host "[5/5] 로컬 설치 생략 (-InstallLocal 로 활성화)" -ForegroundColor DarkGray
}

$sizeMb = [math]::Round((Get-Item -LiteralPath $zipPath).Length / 1MB, 1)
Write-Host ""
Write-Host "버전   : $version"
Write-Host "패키지 : $zipPath ($sizeMb MB)"
Write-Host "SHA256 : $hash"
Write-Host "매니페스트: $latestPath"
