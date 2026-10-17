<#
.SYNOPSIS
    vbaloop installer - downloads the latest release into %LOCALAPPDATA%\vbaloop\bin and adds it to PATH.

.DESCRIPTION
    irm https://raw.githubusercontent.com/rfastball/vbaloop/master/install.ps1 | iex

    One line from the folder you want to work in: run it next to an existing .xlsm to import
    that workbook, or in an empty folder to start a new project. In an interactive terminal the
    installer continues straight into `vbaloop setup` there. Shells without a console input
    (coding agents, CI, -NonInteractive) are never prompted; they get agent guidance instead.

    To install only (no setup), set VBALOOP_NO_SETUP=1 first, or pass -NoSetup:

    & ([scriptblock]::Create((irm https://raw.githubusercontent.com/rfastball/vbaloop/master/install.ps1))) -NoSetup

    Equivalent to downloading vbaloop.exe and double-clicking it, except for SmartScreen:
    an unsigned exe fetched by a browser goes through the "Windows protected your PC" prompt,
    this script path does not.

    The actual dev environment setup is done by `vbaloop setup`
    (environment check -> folder -> workbook -> harness).

    Encoding: this file must stay pure ASCII with no BOM. Windows PowerShell 5.1 reads a
    BOM-less local .ps1 in the system ANSI code page (ASCII is safe in every code page), and
    `irm | iex` keeps a UTF-8 BOM as U+FEFF, which breaks the first token of the script.

.PARAMETER Setup
    Run the vbaloop setup wizard right after installing, even when the shell does not look
    interactive.

.PARAMETER NoSetup
    Install only. Same as setting the VBALOOP_NO_SETUP environment variable.

.PARAMETER Version
    Install a specific release tag (e.g. v1.2.3). Defaults to the latest release.
#>
[CmdletBinding()]
param(
    [switch]$Setup,
    [switch]$NoSetup,
    [string]$Version
)

# Everything runs in a child scope: under `irm | iex` the script executes in the caller's
# session, so preferences set here must not leak into it, and `return` (never `exit`)
# ends the install without closing the user's shell.
& {
    $ErrorActionPreference = 'Stop'
    # Windows PowerShell 5.1 redraws the progress bar per chunk and slows downloads ~10x.
    $ProgressPreference = 'SilentlyContinue'
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $Repo       = 'rfastball/vbaloop'
    $InstallDir = Join-Path $env:LOCALAPPDATA 'vbaloop\bin'
    $ExePath    = Join-Path $InstallDir 'vbaloop.exe'

    function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
    function Write-Ok($msg)   { Write-Host "    $msg" -ForegroundColor Green }
    function Write-Warn2($msg){ Write-Host "    $msg" -ForegroundColor Yellow }

    # --- 0. Preconditions ---------------------------------------------------
    if ($env:OS -ne 'Windows_NT') {
        throw 'vbaloop is Windows-only (it automates Excel through COM).'
    }

    # --- 1. Resolve release -------------------------------------------------
    Write-Step 'Resolving release'
    $api = if ($Version) {
        "https://api.github.com/repos/$Repo/releases/tags/$Version"
    } else {
        "https://api.github.com/repos/$Repo/releases/latest"
    }
    try {
        $release = Invoke-RestMethod -Uri $api -Headers @{ 'User-Agent' = 'vbaloop-install' } -UseBasicParsing
    } catch {
        throw "Could not fetch release info: $($_.Exception.Message)`n     $api"
    }
    $tag = $release.tag_name
    Write-Ok "$tag"

    $exeAsset = $release.assets | Where-Object { $_.name -eq 'vbaloop.exe' } | Select-Object -First 1
    $sumAsset = $release.assets | Where-Object { $_.name -eq 'checksums.txt' } | Select-Object -First 1
    if (-not $exeAsset) { throw "Release $tag has no vbaloop.exe asset." }
    # Fail closed: an unverifiable download is never installed.
    if (-not $sumAsset) { throw "Release $tag has no checksums.txt asset - refusing to install an unverified exe." }

    # --- 2. Download + verify -----------------------------------------------
    Write-Step 'Downloading'
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("vbaloop-" + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    try {
        $tmpExe = Join-Path $tmp 'vbaloop.exe'
        Invoke-WebRequest -Uri $exeAsset.browser_download_url -OutFile $tmpExe -UseBasicParsing
        Write-Ok ("{0:N1} MB" -f ((Get-Item -LiteralPath $tmpExe).Length / 1MB))

        # Download to a file and read it back: Invoke-WebRequest .Content is byte[] for
        # application/octet-stream responses, which would make the text match silently fail.
        $tmpSums = Join-Path $tmp 'checksums.txt'
        Invoke-WebRequest -Uri $sumAsset.browser_download_url -OutFile $tmpSums -UseBasicParsing
        $expected = Get-Content -LiteralPath $tmpSums |
            ForEach-Object { if ($_ -match '^\s*([0-9a-fA-F]{64})\s+\*?vbaloop\.exe\s*$') { $Matches[1] } } |
            Select-Object -First 1
        if (-not $expected) {
            throw "checksums.txt in release $tag has no vbaloop.exe entry - refusing to install an unverified exe."
        }
        $actual = (Get-FileHash -LiteralPath $tmpExe -Algorithm SHA256).Hash
        if ($actual -ne $expected) {   # -ne is case-insensitive
            throw "Checksum mismatch - the download is corrupted or tampered with.`n     expected $($expected.ToLower())`n     actual   $($actual.ToLower())"
        }
        Write-Ok 'SHA256 verified'

        # --- 3. Install ------------------------------------------------------
        Write-Step 'Installing'
        New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
        try {
            Copy-Item -LiteralPath $tmpExe -Destination $ExePath -Force
        } catch {
            throw "Install failed (file probably locked) - close any running vbaloop and retry.`n     $($_.Exception.Message)"
        }
        Write-Ok $ExePath
    } finally {
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- 4. PATH (idempotent) -----------------------------------------------
    # setx truncates at 1024 chars, so it is not used - write the registry directly as
    # ExpandString and compare against the unexpanded text (keeps tokens like %USERPROFILE%).
    Write-Step 'Adding to PATH'
    $raw = (Get-Item 'HKCU:\Environment').GetValue('Path', '', 'DoNotExpandEnvironmentNames')
    $entries = @($raw -split ';' | Where-Object { $_ -ne '' })
    $already = $entries | Where-Object { $_.Trim().TrimEnd('\') -ieq $InstallDir.TrimEnd('\') }
    if (-not $already) {
        $newPath = (@($entries) + $InstallDir) -join ';'
        [Microsoft.Win32.Registry]::SetValue('HKEY_CURRENT_USER\Environment', 'Path', $newPath,
            [Microsoft.Win32.RegistryValueKind]::ExpandString)
        # Registry writes are not seen by Explorer until it is told to reload the environment;
        # without this, terminals opened from the Start menu/Explorer still miss vbaloop.
        # Same broadcast setx performs. Best-effort: a failure only delays it to the next logon.
        try {
            if (-not ('VbaLoopInstall.EnvBroadcast' -as [type])) {
                Add-Type -Namespace VbaLoopInstall -Name EnvBroadcast -MemberDefinition @'
[DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern IntPtr SendMessageTimeoutW(IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
    uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
'@
            }
            $res = [UIntPtr]::Zero
            # HWND_BROADCAST, WM_SETTINGCHANGE, SMTO_ABORTIFHUNG, 5 s per window
            [void][VbaLoopInstall.EnvBroadcast]::SendMessageTimeoutW([IntPtr]0xFFFF, 0x1A, [UIntPtr]::Zero,
                'Environment', 2, 5000, [ref]$res)
        } catch {
            Write-Warn2 "Could not broadcast the PATH change ($($_.Exception.Message)) - sign out and back in if new terminals miss vbaloop."
        }
        Write-Ok "$InstallDir added (applies to newly opened terminals)"
    } else {
        Write-Ok "$InstallDir already on PATH"
    }
    if (($env:Path -split ';') -notcontains $InstallDir) { $env:Path += ";$InstallDir" }

    # --- 5. Next step --------------------------------------------------------
    # A person in a terminal gets the project started right here: the wizard looks at the
    # current folder first (an .xlsm to import, an empty folder for a new project). Agent shells
    # have no console input (stdin redirected / -NonInteractive / CI): a prompt there would hang
    # or read EOF, so they never get the wizard - they get the agent procedure instead.
    $nonInteractiveArg = @([Environment]::GetCommandLineArgs() | Where-Object { $_ -match '^-noni' }).Count -gt 0
    $interactive = ($Host.Name -eq 'ConsoleHost') -and -not [Console]::IsInputRedirected -and
        -not $env:CI -and -not $nonInteractiveArg
    $runSetup = $Setup -or ($interactive -and -not $NoSetup -and -not $env:VBALOOP_NO_SETUP)

    Write-Host ''
    if ($runSetup) {
        Write-Step "Starting vbaloop setup in $((Get-Location).ProviderPath)"
        Write-Host '    Ctrl+C to stop here; run "vbaloop setup" again any time to continue.'
        Write-Host ''
        & $ExePath setup
        return
    }

    Write-Step 'Checking environment'
    & $ExePath doctor
    $doctorExit = $LASTEXITCODE

    Write-Host ''
    if ($interactive) {
        Write-Host 'Installed. Next step:' -ForegroundColor Cyan
        Write-Host '  vbaloop setup      dev environment wizard (workbook or new project -> harness -> verify)'
    } else {
        Write-Host 'Installed. Next steps for a coding agent:' -ForegroundColor Cyan
        Write-Host "  exe        $ExePath"
        Write-Host '             (use this absolute path - this shell''s PATH may predate the install)'
        Write-Host '  procedure  https://github.com/rfastball/vbaloop/blob/master/AGENT_SETUP.md'
        Write-Host '  preview    vbaloop setup --plan --json --workbook <path.xlsm> | --new <name> | --folder <dir>'
        Write-Host '  consent    ask the user before adding --enable-vbom (it turns on an Excel security setting)'
    }
    if ($doctorExit -ne 0) {
        Write-Host ''
        Write-Warn2 'doctor reported environment issues - setup reports the ones it can fix.'
    }
}
