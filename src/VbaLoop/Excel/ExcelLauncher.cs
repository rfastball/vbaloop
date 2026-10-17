using System.Diagnostics;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// 도구 소유 세션 생성(설계 §4 session start): 파일 연결 실행 금지 —
/// COM으로 직접 프로세스를 생성하고, 생성 전 스냅샷에 없던 새 PID인지 검증한다.
/// </summary>
public static class ExcelLauncher
{
    public sealed record LaunchResult(dynamic App, int Pid, long Hwnd, string StartTimeIso);

    public static LaunchResult CreateNew(bool visible)
    {
        var before = new HashSet<int>(
            Process.GetProcessesByName("EXCEL").Select(p => p.Id));

        var progType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new CliError("Excel.Application ProgID not found — Excel not installed?",
                "Check the environment with vbaloop doctor", environment: true);

        dynamic app = Activator.CreateInstance(progType)
            ?? throw new CliError("Failed to create Excel COM instance", environment: true);

        try
        {
            app.Visible = visible;

            long hwnd = Convert.ToInt64(app.Hwnd);
            NativeMethods.GetWindowThreadProcessId(new IntPtr(hwnd), out uint pidU);
            int pid = (int)pidU;

            // 필수 단계: 자신이 만든 프로세스인지 검증 (기존 인스턴스에 붙었다면 실패)
            if (before.Contains(pid))
                throw new CliError(
                    $"Newly created session attached to an existing Excel process (PID {pid}) — cannot be a tool-owned session",
                    "Conflict with an existing Excel instance — have the user close their Excel or check DCOM settings",
                    environment: true);

            string startTime = ProcessIdentity.GetStartTimeIso(pid)
                ?? throw new CliError($"Could not read StartTime of PID {pid}", environment: true);

            // doctor --kill-orphans가 "도구가 띄운 PID"를 우선 식별하도록 기록(설계 §11, best-effort)
            ExcelPidRegistry.TryRecord(ExcelPidRegistry.DefaultPath, pid, startTime, ExcelPidRegistry.KindLaunched, null);

            return new LaunchResult(app, pid, hwnd, startTime);
        }
        catch
        {
            try { app.Quit(); } catch { /* 정리 실패는 무시 */ }
            throw;
        }
    }
}
