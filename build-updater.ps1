[CmdletBinding()]
param(
    [string]$Version,
    [string]$OutputDirectory = 'artifacts/publish',
    [string]$DotNet = 'dotnet'
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# 유일한 배포 진입점: 현재 소스 CLI → 해시 manifest → GUI 내장 → 단일 exe → 스모크.
# caller의 현재 폴더와 무관하며, 이전 publish 결과를 입력으로 받지 않는다.
Push-Location $PSScriptRoot
try {
    if (-not $Version) {
        $tag = & git describe --tags --match 'v*' --abbrev=0 2>$null
        $Version = if ($LASTEXITCODE -eq 0 -and $tag) { "$tag" -replace '^v', '' } else { '0.1.0' }
    }
    if ($Version -notmatch '^\d+\.\d+\.\d+(-[0-9A-Za-z.-]+)?$') { throw 'Version must be semver, for example 1.2.3 or 1.2.3-rc.1' }
    $sdk = & $DotNet --version
    if ($LASTEXITCODE -ne 0 -or [int]("$sdk".Split('.')[0]) -lt 10) { throw '.NET SDK 10 or newer is required. Use -DotNet <path> if needed.' }

    $output = [IO.Path]::GetFullPath($OutputDirectory, $PSScriptRoot)
    $stage = Join-Path $PSScriptRoot ('artifacts/build/' + [guid]::NewGuid().ToString('N'))
    $cliDir = Join-Path $stage 'cli'
    $engineDir = Join-Path $stage 'engine'
    $guiDir = Join-Path $stage 'gui'
    New-Item -ItemType Directory -Force -Path $stage, $output | Out-Null
    $common = @('-c', 'Release', '-r', 'win-x64', "-p:Version=$Version", '-p:SelfContained=true',
        '-p:PublishSingleFile=true', '-p:IncludeNativeLibrariesForSelfExtract=true',
        '-p:EnableCompressionInSingleFile=true', '-p:PublishTrimmed=false', '-p:DebugType=None')

    & $DotNet publish src/VbaLoop/VbaLoop.csproj @common -o $cliDir
    if ($LASTEXITCODE -ne 0) { throw 'CLI publish failed' }
    $cli = Join-Path $cliDir 'vbaloop.exe'
    # GUI 안에는 런타임을 중복 내장하지 않는다. CLI는 외부 패키지 의존성이 없는 관리 DLL이다.
    & $DotNet publish src/VbaLoop/VbaLoop.csproj -c Release -r win-x64 "-p:Version=$Version" `
        -p:SelfContained=false -p:PublishSingleFile=false -p:UseAppHost=false -p:DebugType=None -o $engineDir
    if ($LASTEXITCODE -ne 0) { throw 'Embedded engine publish failed' }
    $engine = Join-Path $engineDir 'vbaloop.dll'
    $hash = (Get-FileHash -LiteralPath $engine -Algorithm SHA256).Hash
    $manifest = Join-Path $stage 'engine.json'
    @{ version = $Version; sha256 = $hash; size = (Get-Item -LiteralPath $engine).Length; protocol = 1 } |
        ConvertTo-Json | Set-Content -LiteralPath $manifest -Encoding utf8

    & $DotNet publish src/VbaLoop.Updater/VbaLoop.Updater.csproj @common `
        "-p:EmbeddedCliPath=$engine" "-p:EmbeddedCliManifestPath=$manifest" -o $guiDir
    if ($LASTEXITCODE -ne 0) { throw 'GUI publish failed' }
    $gui = Join-Path $guiDir 'vbaloop-updater.exe'

    # WinExe는 PowerShell에서 직접 호출하면 종료를 기다리지 않으므로 Process API 사용.
    $report = Join-Path $stage 'smoke.json'
    $start = [Diagnostics.ProcessStartInfo]::new($gui)
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.ArgumentList.Add('--self-test')
    $start.ArgumentList.Add($report)
    $process = [Diagnostics.Process]::Start($start)
    try {
        if (-not $process.WaitForExit(60000)) {
            $process.Kill($true)
            throw 'Updater smoke timed out (no Excel is launched by this check)'
        }
        if ($process.ExitCode -ne 0) {
            if (Test-Path -LiteralPath $report) { Write-Host (Get-Content -Raw -LiteralPath $report) }
            throw 'Updater smoke failed'
        }
    } finally { $process.Dispose() }
    $result = Get-Content -Raw -LiteralPath $report | ConvertFrom-Json
    if (-not $result.ok -or $result.version -ne $Version -or $result.cliSha256 -ne $hash) {
        throw 'Published updater does not contain the CLI just built'
    }

    Copy-Item -LiteralPath $cli -Destination (Join-Path $output 'vbaloop.exe') -Force
    Copy-Item -LiteralPath $gui -Destination (Join-Path $output 'vbaloop-updater.exe') -Force
    Get-FileHash -LiteralPath (Join-Path $output 'vbaloop.exe'), (Join-Path $output 'vbaloop-updater.exe') -Algorithm SHA256 |
        ForEach-Object { '{0}  {1}' -f $_.Hash.ToLowerInvariant(), [IO.Path]::GetFileName($_.Path) } |
        Set-Content -LiteralPath (Join-Path $output 'checksums.txt') -Encoding ascii
    Get-Item -LiteralPath (Join-Path $output 'vbaloop.exe'), (Join-Path $output 'vbaloop-updater.exe') |
        Select-Object Name, Length, FullName
} finally { Pop-Location }
