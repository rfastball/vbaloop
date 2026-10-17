namespace VbaLoop.Core;

/// <summary>
/// 종료 직전 RCW 정리(GC + 파이널라이저 대기)를 시간 제한 안에서만 수행한다.
/// 격리 모델(설계 §11)에서는 Excel을 kill하지 않으므로, VBA를 실행 중인(메시지 펌프 없는) Excel의
/// COM 프록시 Release가 파이널라이저 스레드를 무기한 붙잡을 수 있다. 무제한 대기하면 명령이
/// 종료되지 않는다(런타임 실측 2026-10-17: test 타임아웃 → 격리 후 부모가 행).
/// </summary>
public static class ComCleanup
{
    /// <summary>이 프로세스가 응답 없는 Excel(격리·busy)의 RCW를 쥐고 있을 수 있음 — 종료 시 RCW 정리를
    /// 건너뛰고 즉시 끝내야 한다. 정상 실패(read-only 거부 등)에서는 정리해야 Quit한 Excel이 실제로 종료된다.</summary>
    public static volatile bool SkipOnExit;

    /// <summary>정리가 제한 시간 안에 끝나면 true. 넘기면 정리를 포기하고 false(백그라운드 스레드는 버려짐).
    /// 주의: 호출 스레드가 STA면 Join이 COM 호출을 펌프하므로, 응답 없는 서버로의 Release를 받아 이 제한도
    /// 무력화될 수 있다 — 그런 상황(exit 3)은 호출부가 <see cref="HardExit"/>로 먼저 끝낸다.</summary>
    public static bool Run(TimeSpan timeout)
    {
        var worker = new Thread(() =>
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        })
        { IsBackground = true, Name = "vbaloop-rcw-cleanup" };
        worker.Start();
        return worker.Join(timeout);
    }

    /// <summary>
    /// 정리가 제한 시간을 넘겼을 때의 종료. 런타임 정상 종료(STA 해제·파이널라이저)는 같은 프록시
    /// Release에서 다시 막히므로(실측: 출력 후에도 프로세스가 미종료) 표준 출력을 비우고 종료 코드를
    /// 보존한 채 즉시 끝낸다.
    /// </summary>
    public static void HardExit(int exitCode)
    {
        try { Console.Out.Flush(); Console.Error.Flush(); } catch { /* 닫힌 파이프 */ }
        TerminateProcess(GetCurrentProcess(), unchecked((uint)exitCode));
    }

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    private static extern IntPtr GetCurrentProcess();

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    private static extern bool TerminateProcess(IntPtr hProcess, uint uExitCode);
}
