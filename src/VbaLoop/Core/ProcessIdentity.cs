using System.Diagnostics;

namespace VbaLoop.Core;

/// <summary>기록된 (PID, StartTime)의 현재 상태 — Gone(소멸/PID 재사용), Match(동일 프로세스), Unknown(살아있으나 확인 불가).</summary>
public enum IdentityState { Gone, Match, Unknown }

/// <summary>
/// (PID, StartTime) 기반 프로세스 생존 판정. vh-core.ps1 Test-AgentProcessAlive 이식:
/// StartTime이 비어 있으면 PID만 신뢰, 읽기 실패 시 살아있다고 가정(보수적).
/// </summary>
public static class ProcessIdentity
{
    public static bool IsAlive(int pid, string? startTimeIso)
    {
        if (pid <= 0) return false;
        Process p;
        try
        {
            p = Process.GetProcessById(pid);
            if (p.HasExited) return false;
        }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }

        if (string.IsNullOrWhiteSpace(startTimeIso)) return true;
        try
        {
            return p.StartTime.ToUniversalTime().ToString("o") == startTimeIso;
        }
        catch
        {
            return true;
        }
    }

    /// <summary>
    /// kill/해제 판정용 엄격 신원. IsAlive와 달리 "읽기 실패 = 살아있음"으로 뭉개지 않고
    /// Unknown으로 분리한다 — Unknown은 살아있다고 보되 신원 확인이 안 된 것이라 kill 근거가
    /// 될 수 없다(설계 §11 kill 규칙). startTimeIso가 비어 있으면 Match를 줄 수 없다.
    /// </summary>
    public static IdentityState Identify(int pid, string? startTimeIso)
    {
        if (pid <= 0) return IdentityState.Gone;
        Process p;
        try
        {
            p = Process.GetProcessById(pid);
            if (p.HasExited) return IdentityState.Gone;
        }
        catch (ArgumentException) { return IdentityState.Gone; }
        catch (InvalidOperationException) { return IdentityState.Gone; }

        using (p)
        {
            if (string.IsNullOrWhiteSpace(startTimeIso)) return IdentityState.Unknown;
            try
            {
                return p.StartTime.ToUniversalTime().ToString("o") == startTimeIso
                    ? IdentityState.Match
                    : IdentityState.Gone; // 같은 PID의 다른 프로세스 = 기록된 프로세스는 소멸
            }
            catch
            {
                return IdentityState.Unknown;
            }
        }
    }

    public static string? GetStartTimeIso(int pid)
    {
        try
        {
            return Process.GetProcessById(pid).StartTime.ToUniversalTime().ToString("o");
        }
        catch
        {
            return null;
        }
    }
}
