using System.Diagnostics;
using System.Runtime.InteropServices;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// 세션 Excel 종료 감시(크래시 판정, 설계 §4 test/run). 프로세스 핸들을 작업 전에 열어 쥐어야 종료 뒤에도 종료 코드를
/// 읽을 수 있다 — 크래시 핸들러를 거친 Excel은 0으로 끝나므로(실측) 종료 코드는 보조 징후이고, 판정은
/// ExcelCrash.Classify가 크래시 의심 오류·재기동 Excel과 함께 내린다. 신원(PID+StartTime)이 맞을 때만 핸들을 잡고,
/// 못 잡았으면 ProcessIdentity로 생존만 판정한다(종료 코드는 null = 판독 불가).
/// </summary>
public sealed class ExcelExitWatch : IDisposable
{
    /// <summary>크래시 의심 오류 뒤 프로세스 소멸을 기다리는 상한. 실측(2026-11-02): Application.Run이 6ms 만에
    /// DISP_E_EXCEPTION으로 돌아오고 프로세스는 1초 안에 사라졌다 — 상한은 덤프 기록이 느린 PC를 위한 여유.
    /// 일반 VBA 런타임 오류는 0x800A9C68이라 이 대기를 타지 않는다.</summary>
    public const int SettleMs = 5000;

    /// <summary>크래시 확인 뒤 Office가 Excel을 다시 띄우기를 기다리는 상한. 실측: 재기동 Excel은 크래시한
    /// 프로세스가 사라질 때 이미 떠 있었다(부모가 죽기 전에 띄운다).</summary>
    public const int RelaunchWaitMs = 3000;

    private const int DispException = unchecked((int)0x80020009); // DISP_E_EXCEPTION

    private readonly int _pid;
    private readonly string _startIso;
    private readonly Process? _proc;

    private ExcelExitWatch(int pid, string startIso, Process? proc)
    {
        _pid = pid;
        _startIso = startIso;
        _proc = proc;
    }

    public int Pid => _pid;

    public static ExcelExitWatch Open(int pid, string startIso)
    {
        Process? p = null;
        if (ProcessIdentity.Identify(pid, startIso) == IdentityState.Match)
        {
            try
            {
                p = Process.GetProcessById(pid);
                _ = p.SafeHandle; // 핸들을 지금 열어 둔다 — 종료 뒤엔 OpenProcess가 실패해 종료 코드를 못 읽는다
            }
            catch
            {
                p?.Dispose();
                p = null;
            }
        }
        return new ExcelExitWatch(pid, startIso, p);
    }

    /// <summary>크래시 의심 HRESULT: DISP_E_EXCEPTION(매크로 안 액세스 위반의 실측 반환) 또는 RPC 단절 계열.</summary>
    public static bool IsCrashSuspect(Exception ex)
        => ex is InvalidComObjectException
           || (ex is COMException c && (c.HResult == DispException || ComRetry.IsPoison(c.HResult)));

    /// <summary>프로세스가 종료됐는지 waitMs까지 기다린다. 0이면 즉시 판정.</summary>
    public bool WaitExit(int waitMs)
    {
        if (_proc != null)
        {
            try { return _proc.WaitForExit(waitMs); }
            catch { return ProcessIdentity.Identify(_pid, _startIso) == IdentityState.Gone; }
        }
        long deadline = Environment.TickCount64 + waitMs;
        while (true)
        {
            if (ProcessIdentity.Identify(_pid, _startIso) == IdentityState.Gone) return true;
            if (Environment.TickCount64 >= deadline) return false;
            Thread.Sleep(100);
        }
    }

    /// <summary>COM 실패 직후 판정: 크래시 의심 오류면 SettleMs까지 기다리고, 아니면 즉시 확인만 한다
    /// (일반 VBA 런타임 오류마다 수 초씩 기다리지 않도록).</summary>
    public ExcelExitKind ClassifyAfter(Exception ex)
    {
        bool suspect = IsCrashSuspect(ex);
        return Final(WaitExit(suspect ? SettleMs : 0), suspect);
    }

    /// <summary>오류 없이 끝난 뒤(또는 부모 백스톱)의 판정 — 크래시 의심 오류 징후 없이 본다.</summary>
    public ExcelExitKind Classify(int waitMs) => Final(WaitExit(waitMs), suspect: false);

    /// <summary>다른 징후가 없으면 재기동 Excel을 한 번 찾아 본다 — 재기동은 크래시한 프로세스가 사라질 때 이미
    /// 떠 있으므로(실측) 스냅샷 1회면 된다.</summary>
    private ExcelExitKind Final(bool exited, bool suspect)
    {
        if (!exited) return ExcelExitKind.Alive;
        var kind = ExcelCrash.Classify(true, ExitCode, suspect);
        if (kind == ExcelExitKind.Exited && FindRelaunched(_pid, _startIso, 0).Count > 0)
            kind = ExcelExitKind.Crashed;
        return kind;
    }

    /// <summary>종료됐고 핸들을 쥐고 있었으면 종료 코드, 아니면 null.</summary>
    public int? ExitCode
    {
        get
        {
            if (_proc == null) return null;
            try { return _proc.HasExited ? _proc.ExitCode : null; }
            catch { return null; }
        }
    }

    /// <summary>소멸 정보를 만든다 — 크래시면 재기동 Excel이 뜨기를 RelaunchWaitMs까지 기다리고, 정본의 `~$` 잠금
    /// 파일을 확인한다(재기동 Excel이 정본을 다시 열었는지의 COM 없는 단서).</summary>
    public ExcelCrashInfo Describe(ExcelExitKind kind, string workbookPath)
    {
        var relaunched = FindRelaunched(_pid, _startIso, kind == ExcelExitKind.Crashed ? RelaunchWaitMs : 0);
        bool locked = false;
        try
        {
            string? dir = Path.GetDirectoryName(workbookPath);
            locked = dir != null && File.Exists(Path.Combine(dir, "~$" + Path.GetFileName(workbookPath)));
        }
        catch { /* 판독 실패 = 모름 */ }
        return new(_pid, ExitCode, relaunched, locked);
    }

    /// <summary>크래시한 Excel을 부모로 하는 새 EXCEL.EXE를 찾는다. 하나라도 보이면 바로 반환한다.</summary>
    public static IReadOnlyList<int> FindRelaunched(int crashedPid, string crashedStartIso, int waitMs)
    {
        DateTime? crashedStart = DateTime.TryParse(crashedStartIso, null,
            System.Globalization.DateTimeStyles.RoundtripKind, out var t) ? t.ToUniversalTime() : null;
        long deadline = Environment.TickCount64 + waitMs;
        while (true)
        {
            var found = ExcelCrash.FindRelaunched(ProcessTree.Snapshot().Values, crashedPid, crashedStart, StartUtcOf);
            if (found.Count > 0 || Environment.TickCount64 >= deadline) return found;
            Thread.Sleep(250);
        }
    }

    private static DateTime? StartUtcOf(int pid)
    {
        try
        {
            using var p = Process.GetProcessById(pid);
            return p.StartTime.ToUniversalTime();
        }
        catch
        {
            return null;
        }
    }

    public void Dispose() => _proc?.Dispose();
}
