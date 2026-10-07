param(
    [string]$Configuration = "Release",
    [switch]$Sign
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$projectPath = Join-Path $projectRoot "BckPPT.csproj"
$publishPath = Join-Path $projectRoot "bin\publish\win-x64"
$artifactsPath = Join-Path $projectRoot "artifacts"
# 버전: 커밋 기반 (version.json + tools\Get-CommitVersion.ps1). 미커밋 변경이 있으면 -dev.
$commitVersion = & (Join-Path $projectRoot "tools\Get-CommitVersion.ps1") -ConfigPath (Join-Path $projectRoot "version.json")
$version = $commitVersion.Label

# publish는 증분 복사라 이전 버전 DLL이 남는다(stale 어셈블리 → 실행 즉시 로드 실패).
# ReadyToRun 중간 산출물(obj\...\R2R)도 캐시되므로 함께 지운다.
if (Test-Path -LiteralPath $publishPath) { Remove-Item -LiteralPath $publishPath -Recurse -Force }
Get-ChildItem -LiteralPath (Join-Path $projectRoot "obj") -Recurse -Directory -Filter "R2R" -ErrorAction SilentlyContinue |
    ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force }

dotnet publish $projectPath -c $Configuration -r win-x64 --self-contained true `
    -p:PublishProfile=win-x64 --nologo `
    -p:Version=$($commitVersion.Version) -p:InformationalVersion=$($commitVersion.Informational)
if ($LASTEXITCODE -ne 0) { throw "dotnet publish 실패" }

$exePath = Join-Path $publishPath "BckPPT.exe"
if (-not (Test-Path -LiteralPath $exePath)) { throw "게시 결과에서 BckPPT.exe를 찾을 수 없습니다." }

if ($Sign) {
    $signTool = $env:BCKPPT_SIGNTOOL
    $certificateSha1 = $env:BCKPPT_CERT_SHA1
    if (-not $signTool -or -not (Test-Path -LiteralPath $signTool)) {
        throw "서명하려면 BCKPPT_SIGNTOOL에 signtool.exe 절대 경로를 설정하세요."
    }
    if (-not $certificateSha1) { throw "서명하려면 BCKPPT_CERT_SHA1을 설정하세요." }
    & $signTool sign /sha1 $certificateSha1 /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 $exePath
    if ($LASTEXITCODE -ne 0) { throw "코드 서명 실패" }
    & $signTool verify /pa /v $exePath
    if ($LASTEXITCODE -ne 0) { throw "코드 서명 검증 실패" }
}

New-Item -ItemType Directory -Path $artifactsPath -Force | Out-Null
$zipPath = Join-Path $artifactsPath "BckPPT-$version-win-x64.zip"
$resolvedArtifacts = [IO.Path]::GetFullPath($artifactsPath).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
$resolvedZip = [IO.Path]::GetFullPath($zipPath)
if (-not $resolvedZip.StartsWith($resolvedArtifacts, [StringComparison]::OrdinalIgnoreCase)) {
    throw "패키지 경로가 artifacts 폴더 밖을 가리킵니다."
}
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path (Join-Path $publishPath "*") -DestinationPath $zipPath -CompressionLevel Optimal
$hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash
$hashPath = "$zipPath.sha256"
Set-Content -LiteralPath $hashPath -Value "$hash  $(Split-Path -Leaf $zipPath)" -Encoding ascii

Write-Host "패키지: $zipPath"
Write-Host "SHA256: $hash"
