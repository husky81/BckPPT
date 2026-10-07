# Get-CommitVersion.ps1 — commit-based product version.
#
# Rules: cellbox2 docs/plans/unified-commit-versioning.md. The same file is
# copied verbatim into BckPPT and bckhwp; keep the copies identical.
#
#   <majorMinor>.<PATCH>[-dev]+<short commit>
#   PATCH = (commits touching `paths` reachable from HEAD) - baseCount
#
# version.json (edited by hand only, when bumping MAJOR.MINOR):
#   { "majorMinor": "2.1", "baseCount": 892, "paths": ["apps/engine", ...] }
# Omit "paths" (or leave it empty) to count the whole repository. When bumping
# majorMinor, set baseCount to (current count + 1) in the same commit so the
# bump commit itself is PATCH 0.
#
# A build never writes anything: the same commit always yields the same number.
#
# Usage:  $v = & installer/Get-CommitVersion.ps1 -ConfigPath installer/version.json
#         $v.Version  -> 2.1.5            (numeric: FileVersion / AssemblyVersion)
#         $v.Label    -> 2.1.5 or 2.1.5-dev (installer name, InformationalVersion prefix)
#         $v.Informational -> 2.1.5+abcdef12 / 2.1.5-dev+abcdef12
#         $v.Commit, $v.ShortCommit, $v.Dirty, $v.Count
param(
    [Parameter(Mandatory = $true)][string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

$ConfigPath = (Resolve-Path -LiteralPath $ConfigPath).Path
$config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$majorMinor = [string]$config.majorMinor
if ($majorMinor -notmatch '^\d+\.\d+$') {
    throw "${ConfigPath}: majorMinor must be MAJOR.MINOR, got '$majorMinor'."
}
$baseCount = [int]$config.baseCount
$paths = @()
if ($config.paths) { $paths = @($config.paths | ForEach-Object { [string]$_ }) }

$repoRoot = (& git -C (Split-Path -Parent $ConfigPath) rev-parse --show-toplevel 2>$null)
if ($LASTEXITCODE -ne 0 -or -not $repoRoot) { throw "Not inside a git repository: $ConfigPath" }
$repoRoot = $repoRoot.Trim()

# A shallow clone undercounts, which would silently ship a smaller number.
$shallow = (& git -C $repoRoot rev-parse --is-shallow-repository).Trim()
if ($shallow -eq 'true') { throw "Shallow clone: commit count would be wrong. Fetch full history (fetch-depth: 0)." }

$commit = (& git -C $repoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or -not $commit) { throw "git rev-parse HEAD failed in $repoRoot" }

$countArgs = @('-C', $repoRoot, 'rev-list', '--count', 'HEAD')
$statusArgs = @('-C', $repoRoot, 'status', '--porcelain', '--untracked-files=no')
if ($paths.Count -gt 0) {
    $countArgs += '--'; $countArgs += $paths
    $statusArgs += '--'; $statusArgs += $paths
}
$count = [int]((& git @countArgs) | Out-String).Trim()
if ($LASTEXITCODE -ne 0) { throw "git rev-list --count failed in $repoRoot" }
$dirty = -not [string]::IsNullOrWhiteSpace(((& git @statusArgs) | Out-String))

$patch = $count - $baseCount
if ($patch -lt 0) {
    throw "Commit count $count is below baseCount $baseCount ($ConfigPath). History rewritten, or baseCount set too high."
}

$version = "$majorMinor.$patch"
$label = if ($dirty) { "$version-dev" } else { $version }
$short = $commit.Substring(0, 8)

[pscustomobject]@{
    Version       = $version
    Label         = $label
    Informational = "$label+$short"
    Commit        = $commit
    ShortCommit   = $short
    Dirty         = $dirty
    Count         = $count
}
