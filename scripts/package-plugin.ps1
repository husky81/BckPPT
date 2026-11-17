# BckPPT를 CellBox 플러그인 패키지로 묶는다.
#
#   scripts\package-plugin.ps1                 # publish → plugin.json → zip → sha256 → latest.json
#   scripts\package-plugin.ps1 -InstallLocal   # 위 + %LOCALAPPDATA%\CellBox\Plugins\BckPPT 에 전개
#   scripts\package-plugin.ps1 -SkipBuild      # 기존 publish 결과 재사용 (빠른 재패키징)
#   scripts\package-plugin.ps1 -Upload         # 위 + 릴리스 버킷 releases/plugins/bckppt/ 에 게시
#
# 산출물 (artifacts/):
#   bckppt-<version>.zip          플러그인 폴더 통째 (plugin.json + self-contained 런타임)
#   bckppt-<version>.zip.sha256
#   latest.json                   릴리스 버킷의 plugins/bckppt/latest.json 로 올릴 매니페스트
#
# 참고: docs/plans/cellbox-tray-plugin.md

param(
    [string]$Configuration = "Release",
    # 릴리스 버킷 공개 주소. releases.cellbox.ai는 DNS에 없는 이름이었고(2026-10-07 확인),
    # 버킷은 releases/* 만 공개 읽기라 그 아래에 둔다.
    [string]$BaseUrl = "https://cellbox-release.s3.ap-northeast-2.amazonaws.com/releases/plugins/bckppt",
    [string]$S3Prefix = "s3://cellbox-release/releases/plugins/bckppt",
    [string]$AwsProfile = "cellbox-release",
    [switch]$Upload,
    [string]$ReleaseNotes = "",
    [switch]$InstallLocal,
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$projectRoot   = Split-Path -Parent $PSScriptRoot
$projectPath   = Join-Path $projectRoot "BckPPT.csproj"
$publishPath   = Join-Path $projectRoot "bin\publish\win-x64"
$artifactsPath = Join-Path $projectRoot "artifacts"
# 버전: 커밋 기반 (version.json + tools\Get-CommitVersion.ps1). 같은 커밋 = 같은 번호.
# 매니페스트 version은 트레이 PluginInstaller가 Version.TryParse로 비교하므로 숫자만 쓴다.
# 미커밋 변경이 있는 빌드는 게시하지 않는다(-InstallLocal 로컬 설치만 허용).
$commitVersion = & (Join-Path $projectRoot "tools\Get-CommitVersion.ps1") -ConfigPath (Join-Path $projectRoot "version.json")
$version       = $commitVersion.Version
if ($commitVersion.Dirty -and $Upload) { throw "미커밋 빌드는 게시할 수 없습니다." }
if ($commitVersion.Dirty -and -not $InstallLocal) {
    throw "미커밋 변경이 있습니다(v$($commitVersion.Label)). 커밋한 뒤 패키징하세요. 로컬 설치만이면 -InstallLocal."
}

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
        -p:PublishProfile=win-x64 --nologo `
        -p:Version=$version -p:InformationalVersion=$($commitVersion.Informational)
    if ($LASTEXITCODE -ne 0) { throw "dotnet publish 실패" }
    # Claude Code·Codex가 띄우는 stdio 프록시. 등록할 때 플러그인 폴더 밖으로 복사되므로 단일 파일이어야 한다.
    dotnet publish (Join-Path $projectRoot "McpProxy\BckPptMcp.csproj") -c $Configuration -r win-x64 --self-contained true --nologo `
        -p:PublishSingleFile=true -p:PublishReadyToRun=true -p:DebugType=none `
        -p:PublishDir="$publishPath\mcp\" `
        -p:Version=$version -p:InformationalVersion=$($commitVersion.Informational)
    if ($LASTEXITCODE -ne 0) { throw "MCP 프록시 publish 실패" }
} else {
    Write-Host "[1/5] publish 생략 (-SkipBuild)" -ForegroundColor DarkGray
}

$exePath = Join-Path $publishPath "BckPPT.exe"
if (-not (Test-Path -LiteralPath $exePath)) { throw "게시 결과에서 BckPPT.exe를 찾을 수 없습니다: $publishPath" }
if (-not (Test-Path -LiteralPath (Join-Path $publishPath "mcp\BckPptMcp.exe"))) { throw "게시 결과에 MCP 프록시(mcp\BckPptMcp.exe)가 없습니다." }
# -SkipBuild로 재사용하는 publish 결과가 지금 커밋의 빌드인지 확인한다.
$publishedStamp = (Get-Item -LiteralPath $exePath).VersionInfo.ProductVersion
if ($publishedStamp -ne $commitVersion.Informational) {
    throw "publish 결과($publishedStamp)가 현재 커밋 버전($($commitVersion.Informational))과 다릅니다. -SkipBuild 없이 다시 실행하세요."
}

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
    # 셀박스 트레이가 Claude Code·Codex 연결을 묻고 프록시의 --mcp-*로 등록한다 (bckhwp docs/mcp-client-registration.md).
    mcpRegistrar      = "mcp/BckPptMcp.exe"
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

    # 파일 단위로 바꾼다. 폴더 이름 바꾸기(Rename-Item)는 CellBox2 트레이가 플러그인 폴더를 쥐고 있으면
    # 「사용 중」으로 실패했다(2026-11-17). 새 판을 임시 폴더에 풀고 → 지금 판을 임시 백업으로 복사하고 →
    # 파일을 덮어쓴 뒤 → 새 판에 없는 옛 파일을 지운다. 덮어쓰기가 실패하면 백업에서 되돌린다.
    # 백업은 덮어쓸 파일만, 임시 폴더에 둔다.
    $stamp   = [Guid]::NewGuid().ToString("N").Substring(0, 8)
    $staging = Join-Path ([IO.Path]::GetTempPath()) "bckppt-install-$stamp"
    $backup  = Join-Path ([IO.Path]::GetTempPath()) "bckppt-backup-$stamp"
    Expand-Archive -LiteralPath $zipPath -DestinationPath $staging -Force
    New-Item -ItemType Directory -Path $target -Force | Out-Null

    function Copy-Tree([string]$From, [string]$To) {
        # robocopy: 같은 파일도 덮어쓰고(/IS /IT) 잠긴 파일은 몇 번 다시 시도한다. 8 이상이 실패.
        $null = robocopy $From $To /E /IS /IT /R:3 /W:1 /NFL /NDL /NJH /NJS /NP
        if ($LASTEXITCODE -ge 8) { throw "파일 복사 실패(robocopy $LASTEXITCODE): $From → $To" }
        $global:LASTEXITCODE = 0
    }

    try {
        # 덮어쓸 파일만 백업한다. 새 판에 없는 파일은 트레이가 잠가 둘 수 있어 읽지 않는다.
        # 덮어쓸 파일을 못 읽으면 아무것도 바꾸기 전에 멈춘다.
        $stagingFull = [IO.Path]::GetFullPath($staging).TrimEnd('\') + '\'
        Get-ChildItem -LiteralPath $staging -Recurse -File | ForEach-Object {
            $relative = $_.FullName.Substring($stagingFull.Length)
            $existing = Join-Path $target $relative
            if (Test-Path -LiteralPath $existing) {
                $dest = Join-Path $backup $relative
                New-Item -ItemType Directory -Path (Split-Path -Parent $dest) -Force | Out-Null
                Copy-Item -LiteralPath $existing -Destination $dest -Force -ErrorAction Stop
            }
        }
        try {
            Copy-Tree $staging $target
        } catch {
            Write-Warning "덮어쓰기 실패 — 이전 판으로 되돌립니다: $_"
            # 잠겨서 못 덮어쓴 파일은 옛 판 그대로라 되돌리기도 실패할 수 있다. 원래 오류를 알린다.
            if (Test-Path -LiteralPath $backup) { try { Copy-Tree $backup $target } catch { Write-Warning "되돌리기 일부 실패: $_" } }
            throw
        }
        # 새 판에 없는 옛 파일(이전 버전 DLL 등)을 지운다. 잠겨 있으면 남기고 알린다(실행에는 쓰이지 않는다).
        $stagingRoot = [IO.Path]::GetFullPath($staging).TrimEnd('\') + '\'
        $targetRoot  = [IO.Path]::GetFullPath($target).TrimEnd('\') + '\'
        $leftover = 0
        Get-ChildItem -LiteralPath $target -Recurse -File | ForEach-Object {
            $relative = $_.FullName.Substring($targetRoot.Length)
            if (-not (Test-Path -LiteralPath (Join-Path $stagingRoot $relative))) {
                try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop }
                catch { $leftover++; Write-Warning "  옛 파일을 지우지 못했습니다(남겨 둠): $relative" }
            }
        }
        Get-ChildItem -LiteralPath $target -Recurse -Directory | Sort-Object { $_.FullName.Length } -Descending | ForEach-Object {
            if (-not (Get-ChildItem -LiteralPath $_.FullName -Force | Select-Object -First 1)) {
                try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { }
            }
        }
        if ($leftover -gt 0) { Write-Host "  옛 파일 $($leftover)개는 잠겨 있어 남겼습니다." -ForegroundColor Yellow }
    } finally {
        foreach ($dir in @($staging, $backup)) {
            if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
    Write-Host "  설치: $target" -ForegroundColor Green
} else {
    Write-Host "[5/5] 로컬 설치 생략 (-InstallLocal 로 활성화)" -ForegroundColor DarkGray
}

if ($Upload) {
    Write-Host "[게시] 릴리스 버킷에 올리기..." -ForegroundColor Cyan
    # zip을 먼저 올리고 latest.json을 마지막에 바꾼다 — 매니페스트가 없는 zip을 가리키는 순간이 없게.
    aws s3 cp $zipPath "$S3Prefix/$zipName" --profile $AwsProfile --content-type application/zip
    if ($LASTEXITCODE -ne 0) { throw "zip 업로드 실패" }
    aws s3 cp $latestPath "$S3Prefix/latest.json" --profile $AwsProfile `
        --content-type application/json --cache-control "no-cache"
    if ($LASTEXITCODE -ne 0) { throw "latest.json 업로드 실패" }

    $published = Invoke-RestMethod -Uri "$($BaseUrl.TrimEnd('/'))/latest.json" -TimeoutSec 15
    if ($published.sha256 -ne $hash) { throw "게시된 latest.json이 방금 만든 패키지와 다릅니다." }
    Write-Host "  게시: $($latest.packageUrl)" -ForegroundColor Green
}

$sizeMb = [math]::Round((Get-Item -LiteralPath $zipPath).Length / 1MB, 1)
Write-Host ""
Write-Host "버전   : $version"
Write-Host "패키지 : $zipPath ($sizeMb MB)"
Write-Host "SHA256 : $hash"
Write-Host "매니페스트: $latestPath"
