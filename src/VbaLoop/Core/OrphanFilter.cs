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
            return new(false, "창 있음(사람/세션 소유로 간주)");
        if (c.IsLiveProjectSession)
            return new(false, "현재 프로젝트의 세션");
        if (c.Quarantined)
            return new(false, "격리 중인 Excel — 사용자 확인 필요(vbaloop recovery status)");
        if (c.WindowsSessionId != currentWindowsSessionId)
            return new(false, $"다른 Windows 세션({c.WindowsSessionId}) — 다른 사용자/RDS 소유 가능성");
        if (c.CommandLine == null)
            return new(false, "명령줄 판독 불가(권한) — 소유 판별 불가");
        if (!IsAutomationLaunch(c.CommandLine))
            return new(false, "/automation 기동 아님 — OLE 임베디드(Word/PowerPoint 등) 또는 사용자 실행 가능성");
        return new(true, c.ToolLaunched
            ? "vbaloop가 띄운 창 없는 자동화 Excel"
            : "창 없는 자동화 Excel(다른 자동화 도구 소유일 수 있음)");
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
