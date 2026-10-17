namespace VbaLoop.Core;

/// <summary>doctor 고아 스캔의 후보 한 건 — 판정에 필요한 사실만 담는다(수집은 DoctorCommand).</summary>
public sealed record OrphanCandidate(
    int Pid,
    string? StartTime,
    bool HasMainWindow,
    int WindowsSessionId,
    string? CommandLine,
    bool IsLiveProjectSession,
    bool Quarantined,
    bool ToolLaunched);

public sealed record OrphanVerdict(bool Killable, string Reason);

/// <summary>
/// doctor --kill-orphans 대상 판정(설계 §11 L4). 창 없는 EXCEL.EXE라는 것만으로는 부족하다 —
/// Word/PowerPoint의 OLE 임베디드 Excel, 다른 사용자의 RDS 세션 Excel도 창이 없다.
/// kill 가능 = 창 없음 AND 현재 Windows 세션 AND 명령줄에 /automation 토큰(COM 자동화 기동)
/// AND 격리 PID 아님 AND 현재 프로젝트의 살아있는 세션 아님. 그 밖은 사유와 함께 리포트만.
/// </summary>
public static class OrphanFilter
{
    public static OrphanVerdict Classify(OrphanCandidate c, int currentWindowsSessionId)
    {
        if (c.HasMainWindow)
            return new(false, "has windows (treated as user/session owned)");
        if (c.IsLiveProjectSession)
            return new(false, "session of the current project");
        if (c.Quarantined)
            return new(false, "quarantined Excel — needs user confirmation (vbaloop recovery status)");
        if (c.WindowsSessionId != currentWindowsSessionId)
            return new(false, $"other Windows session ({c.WindowsSessionId}) — possibly owned by another user/RDS");
        if (c.CommandLine == null)
            return new(false, "command line unreadable (permissions) — cannot determine owner");
        if (!IsAutomationLaunch(c.CommandLine))
            return new(false, "not launched with /automation — possibly OLE-embedded (Word/PowerPoint etc.) or user-launched");
        return new(true, c.ToolLaunched
            ? "windowless automation Excel launched by vbaloop"
            : "windowless automation Excel (may be owned by another automation tool)");
    }

    /// <summary>
    /// COM 자동화 기동(CreateObject/Activator)은 <c>EXCEL.EXE /automation -Embedding</c>.
    /// OLE 임베디드 활성화는 <c>-Embedding</c>만 붙는다. 경로 안의 부분 문자열을 오인하지 않도록
    /// 공백 토큰 단위로만 비교한다.
    /// </summary>
    public static bool IsAutomationLaunch(string commandLine)
        => commandLine.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Any(t => string.Equals(t, "/automation", StringComparison.OrdinalIgnoreCase));
}
