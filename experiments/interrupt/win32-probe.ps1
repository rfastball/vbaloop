# Win32-only probe (no COM) against the ephemeral Excel of a running `vbaloop run modLoop.LongLoop`.
# 1) list workbook windows (XLMAIN > XLDESK > EXCEL7) without sending messages
# 2) try posted Ctrl+Break variants, watching whether LogShim ticks stop and what dialogs appear
param([string]$Project, [string]$Methods = 'post-xlmain,post-excel7,post-thread')
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class W {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int InternalGetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsHungAppWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool PostThreadMessage(uint t, uint m, IntPtr w, IntPtr l);
    public static string Cls(IntPtr h) { var s = new StringBuilder(256); GetClassName(h, s, 256); return s.ToString(); }
    public static string Txt(IntPtr h) { var s = new StringBuilder(512); InternalGetWindowText(h, s, 512); return s.ToString(); }
    public static List<IntPtr> Top(uint pid) {
        var r = new List<IntPtr>();
        EnumWindows((h, l) => { uint p; GetWindowThreadProcessId(h, out p); if (p == pid) r.Add(h); return true; }, IntPtr.Zero);
        return r;
    }
    public static List<IntPtr> Kids(IntPtr parent) {
        var r = new List<IntPtr>();
        EnumChildWindows(parent, (h, l) => { r.Add(h); return true; }, IntPtr.Zero);
        return r;
    }
}
'@
$dbg = Join-Path $Project '_dbg'
$jsonl = Join-Path $dbg 'Probe.jsonl'
$out = [ordered]@{}
function Ticks { if (Test-Path $jsonl) { @(Select-String -Path $jsonl -Pattern '"event":"tick"').Count } else { 0 } }

# wait for the ephemeral meta and the first ticks
$deadline = (Get-Date).AddSeconds(60)
$meta = $null
while (-not $meta -and (Get-Date) -lt $deadline) {
    $f = Get-ChildItem (Join-Path $dbg 'ephemeral') -Filter *.json -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($f) { $meta = Get-Content $f.FullName -Raw | ConvertFrom-Json } else { Start-Sleep -Milliseconds 200 }
}
if (-not $meta) { throw 'no ephemeral meta' }
$excelPid = [uint32]$meta.pid
$out.pid = $excelPid
$t0 = Ticks
while ((Ticks) -le $t0 + 2 -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }

function Snapshot {
    $tops = [W]::Top($excelPid)
    $xlmain = $tops | Where-Object { [W]::Cls($_) -eq 'XLMAIN' } | Select-Object -First 1
    $books = @()
    if ($xlmain) {
        $books = @([W]::Kids($xlmain) | Where-Object { [W]::Cls($_) -eq 'EXCEL7' } | ForEach-Object { [W]::Txt($_) })
    }
    $dialogs = @($tops | Where-Object { [W]::Cls($_) -eq '#32770' -and [W]::IsWindowVisible($_) } | ForEach-Object {
        $h = $_
        [ordered]@{ caption = [W]::Txt($h); texts = @([W]::Kids($h) | ForEach-Object { [W]::Txt($_) } | Where-Object { $_ }) }
    })
    [ordered]@{
        xlmain = if ($xlmain) { '0x{0:X}' -f [int64]$xlmain } else { $null }
        hung = if ($xlmain) { [W]::IsHungAppWindow($xlmain) } else { $null }
        xlmainEnabled = if ($xlmain) { [W]::IsWindowEnabled($xlmain) } else { $null }
        workbooks = $books
        dialogs = $dialogs
        ticks = Ticks
    }
}

$sw = [Diagnostics.Stopwatch]::StartNew()
$out.busySnapshot = Snapshot
$out.busySnapshotMs = $sw.ElapsedMilliseconds

$WM_KEYDOWN = 0x100; $WM_KEYUP = 0x101; $VK_CANCEL = 0x03; $VK_CONTROL = 0x11
# lParam: repeat 1, scan code of Break (0x46) with extended bit
$down = [IntPtr]0x01460001; $up = [IntPtr]0xC1460001
$out.attempts = @()
foreach ($m in $Methods.Split(',')) {
    $snapBefore = Snapshot
    if ($snapBefore.dialogs.Count -gt 0) { break }
    $xl = [IntPtr]([Convert]::ToInt64($snapBefore.xlmain, 16))
    switch ($m) {
        'post-xlmain' {
            [void][W]::PostMessage($xl, $WM_KEYDOWN, [IntPtr]$VK_CONTROL, [IntPtr]0x001D0001)
            [void][W]::PostMessage($xl, $WM_KEYDOWN, [IntPtr]$VK_CANCEL, $down)
            [void][W]::PostMessage($xl, $WM_KEYUP, [IntPtr]$VK_CANCEL, $up)
        }
        'post-excel7' {
            $e7 = [W]::Kids($xl) | Where-Object { [W]::Cls($_) -eq 'EXCEL7' } | Select-Object -First 1
            [void][W]::PostMessage($e7, $WM_KEYDOWN, [IntPtr]$VK_CANCEL, $down)
            [void][W]::PostMessage($e7, $WM_KEYUP, [IntPtr]$VK_CANCEL, $up)
        }
        'post-thread' {
            [uint32]$ignored = 0
            $tid = [W]::GetWindowThreadProcessId($xl, [ref]$ignored)
            [void][W]::PostThreadMessage($tid, $WM_KEYDOWN, [IntPtr]$VK_CANCEL, $down)
            [void][W]::PostThreadMessage($tid, $WM_KEYUP, [IntPtr]$VK_CANCEL, $up)
        }
    }
    $seen = @()
    for ($i = 0; $i -lt 10; $i++) { Start-Sleep -Milliseconds 400; $s = Snapshot; if ($s.dialogs.Count -gt 0) { $seen += , $s.dialogs } }
    $ticksA = Ticks; Start-Sleep -Seconds 2; $ticksB = Ticks
    $out.attempts += [ordered]@{ method = $m; ticksBefore = $snapBefore.ticks; ticksAfter = $ticksB; stillTicking = ($ticksB -gt $ticksA); dialogsSeen = $seen }
    if ($ticksB -eq $ticksA) { break }
}
$out.final = Snapshot
$out | ConvertTo-Json -Depth 8
