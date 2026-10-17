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
            ?? throw new CliError("Excel.Application ProgID를 찾지 못함 — Excel 미설치?",
                "vbaloop doctor로 환경 점검", environment: true);

        dynamic app = Activator.CreateInstance(progType)
            ?? throw new CliError("Excel COM 인스턴스 생성 실패", environment: true);

        try
        {
            app.Visible = visible;

            long hwnd = Convert.ToInt64(app.Hwnd);
            NativeMethods.GetWindowThreadProcessId(new IntPtr(hwnd), out uint pidU);
            int pid = (int)pidU;

            // 필수 단계: 자신이 만든 프로세스인지 검증 (기존 인스턴스에 붙었다면 실패)
            if (before.Contains(pid))
                throw new CliError(
                    $"새로 생성한 세션이 기존 Excel 프로세스(PID {pid})에 연결됨 — 도구 소유 세션 불가",
                    "기존 Excel 인스턴스와의 충돌 — 사람 Excel을 닫거나 DCOM 설정 확인",
                    environment: true);

            string startTime = ProcessIdentity.GetStartTimeIso(pid)
                ?? throw new CliError($"PID {pid}의 StartTime을 읽지 못함", environment: true);

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
