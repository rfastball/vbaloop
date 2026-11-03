# Win32-only (no COM): once the case macro logs case_ready, list every top-level window of the
# ephemeral Excel and every EXCEL7 descendant with its title and visibility.
param([string]$Project, [int]$ReadyBefore)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class W2 {
    public delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int InternalGetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    public static string Cls(IntPtr h) { var s = new StringBuilder(256); GetClassName(h, s, 256); return s.ToString(); }
    public static string Txt(IntPtr h) { var s = new StringBuilder(512); InternalGetWindowText(h, s, 512); return s.ToString(); }
    public static List<IntPtr> Top(uint pid) {
        var r = new List<IntPtr>();
        EnumWindows(delegate(IntPtr h, IntPtr l) { uint p; GetWindowThreadProcessId(h, out p); if (p == pid) r.Add(h); return true; }, IntPtr.Zero);
        return r;
    }
    public static List<IntPtr> Kids(IntPtr parent) {
        var r = new List<IntPtr>();
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr l) { r.Add(h); return true; }, IntPtr.Zero);
        return r;
    }
}
'@
$dbg = Join-Path $Project '_dbg'
$jsonl = Join-Path $dbg 'Probe.jsonl'
$deadline = (Get-Date).AddSeconds(90)
while ((Get-Date) -lt $deadline) {
    $n = if (Test-Path $jsonl) { @(Select-String -Path $jsonl -Pattern '"case_ready"').Count } else { 0 }
    if ($n -gt $ReadyBefore) { break }
    Start-Sleep -Milliseconds 200
}
Start-Sleep -Milliseconds 1500
$f = Get-ChildItem (Join-Path $dbg 'ephemeral') -Filter *.json | Select-Object -First 1
$excelPid = [uint32](Get-Content $f.FullName -Raw | ConvertFrom-Json).pid
$case = (Select-String -Path $jsonl -Pattern '"case_ready"' | Select-Object -Last 1).Line
$tops = foreach ($t in [W2]::Top($excelPid)) {
    $e7 = @([W2]::Kids($t) | Where-Object { [W2]::Cls($_) -eq 'EXCEL7' } | ForEach-Object {
        [ordered]@{ title = [W2]::Txt($_); visible = [W2]::IsWindowVisible($_) } })
    if ($e7.Count -gt 0 -or [W2]::Cls($t) -eq 'XLMAIN') {
        [ordered]@{ class = [W2]::Cls($t); title = [W2]::Txt($t); visible = [W2]::IsWindowVisible($t); excel7 = $e7 }
    }
}
[ordered]@{ case = $case; pid = $excelPid; tops = @($tops) } | ConvertTo-Json -Depth 6
