# Opt-in real Excel lane: CLI authoring -> push/compile -> saved form event -> regeneration.
param([string]$Dotnet = 'dotnet')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$repo = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$cli = Join-Path $repo 'src/VbaLoop/bin/Release/net10.0-windows/vbaloop.exe'
$hostExe = (Get-Command $Dotnet).Source
$env:DOTNET_ROOT = Split-Path $hostExe
$root = [IO.Directory]::CreateDirectory((Join-Path $repo ('_pilot_form/' + [guid]::NewGuid().ToString('N')))).FullName
$config = Join-Path $root 'Generated/vbaloop.toml'
$checks = [ordered]@{ output = $root }
function Invoke-Cli([string[]]$arguments, [bool]$expectSuccess = $true) {
    $raw = & $cli @arguments --json
    $exitCode = $LASTEXITCODE
    $result = ($raw -join "`n") | ConvertFrom-Json
    if ($result.ok -ne $expectSuccess -or ($expectSuccess -and $exitCode -ne 0)) {
        throw "CLI expectation failed: $($arguments -join ' ')`n$raw"
    }
    return $result
}
function Assert-Probe([bool]$condition, [string]$message) { if (!$condition) { throw $message } }
Push-Location $root
try {
    $null = Invoke-Cli @('setup', '--new', 'Generated', '--no-install', '--skip-verify', '--yes')
    $spec = Join-Path $root 'ProbeForm.form.json'
    $code = Join-Path $root 'ProbeForm.code.vba'
    Copy-Item (Join-Path $PSScriptRoot 'ProbeForm.form.json') $spec
    Copy-Item (Join-Path $PSScriptRoot 'ProbeForm.code.vba') $code
    $buildArgs = @('form', 'build', '--config', $config, '--spec', $spec, '--code', $code)
    $built = Invoke-Cli $buildArgs
    $frm = $built.payload.frm
    $frx = $built.payload.frx
    $checks.createdWithoutTemplate = Test-Path $frx
    Assert-Probe $checks.createdWithoutTemplate 'FRX missing'
    $src = Split-Path $frm
    @'
Attribute VB_Name = "modFormSmoke"
Option Explicit
Public Function Check() As String
    Load ProbeForm
    If ProbeForm.fraGroup.Controls("txtName").Text <> "한글 입력" Then Err.Raise 5
    ProbeForm.cmdProbe.Value = True
    Check = ProbeForm.lblStatus.Caption
    Unload ProbeForm
End Function
'@ | Set-Content (Join-Path $src 'modFormSmoke.bas') -Encoding utf8
    $pushed = Invoke-Cli @('push', '--config', $config)
    Assert-Probe ($pushed.payload.compile.result -eq 'compiled') 'Initial push did not compile'
    $ran = Invoke-Cli @('run', 'modFormSmoke.Check', '--config', $config)
    Assert-Probe ($ran.payload.run.returnValue -ceq '안녕하세요, 한글 입력') 'Saved form event failed'
    $checks.pushCompileAndEvent = $true
    $before = @((Get-FileHash $frm).Hash, (Get-FileHash $frx).Hash)
    $null = Invoke-Cli $buildArgs $false
    Assert-Probe (((Get-FileHash $frm).Hash -eq $before[0]) -and ((Get-FileHash $frx).Hash -eq $before[1])) 'Refused build changed source'
    $checks.overwriteRefusedWithoutForce = $true
    $invalid = (Get-Content $spec -Raw).Replace('"CommandButton"', '"Unknown.ActiveX"')
    $invalid | Set-Content $spec -Encoding utf8
    $null = Invoke-Cli ($buildArgs + '--force') $false
    Assert-Probe (((Get-FileHash $frm).Hash -eq $before[0]) -and ((Get-FileHash $frx).Hash -eq $before[1])) 'Invalid input changed source'
    $checks.invalidInputPreservesPair = $true
    Copy-Item (Join-Path $PSScriptRoot 'ProbeForm.form.json') $spec
    (Get-Content $code -Raw).Replace('안녕하세요, ', '반갑습니다, ') | Set-Content $code -Encoding utf8
    $null = Invoke-Cli ($buildArgs + '--force')
    $pushed = Invoke-Cli @('push', '--config', $config)
    Assert-Probe ($pushed.payload.compile.result -eq 'compiled') 'Regenerated form did not compile'
    $ran = Invoke-Cli @('run', 'modFormSmoke.Check', '--config', $config)
    Assert-Probe ($ran.payload.run.returnValue -ceq '반갑습니다, 한글 입력') 'Regenerated event did not persist'
    $checks.regeneratePushAndEvent = $true
    $checks.status = 'passed'
} catch {
    $checks.status = 'failed'
    $checks.error = $_.Exception.Message
    throw
} finally {
    Pop-Location
    $checks | ConvertTo-Json | Set-Content (Join-Path $root 'result.json') -Encoding utf8
    $checks | ConvertTo-Json
}
