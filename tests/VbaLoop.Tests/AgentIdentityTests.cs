using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class AgentIdentityTests
{
    private static readonly DateTime T0 = new(2026, 7, 8, 9, 0, 0, DateTimeKind.Utc);

    private static Dictionary<int, ProcessTree.ProcEntry> Table(params (int Pid, int PPid, string Name)[] rows)
        => rows.ToDictionary(r => r.Pid, r => new ProcessTree.ProcEntry(r.Pid, r.PPid, r.Name));

    // 조상일수록 이른 StartTime (정상 계보)
    private static Func<int, DateTime?> Starts(params (int Pid, int MinutesAfterT0)[] rows)
    {
        var map = rows.ToDictionary(r => r.Pid, r => T0.AddMinutes(r.MinutesAfterT0));
        return pid => map.TryGetValue(pid, out var t) ? t : null;
    }

    [Fact]
    public void FindsOutermostNonShellAncestor_ClaudeChain()
    {
        // vbaloop(100) <- powershell(90) <- claude(80) <- wt(70=shell) <- explorer(60=boundary)
        var table = Table((100, 90, "vbaloop"), (90, 80, "powershell"), (80, 70, "claude"),
                          (70, 60, "windowsterminal"), (60, 1, "explorer"));
        var starts = Starts((100, 50), (90, 40), (80, 10), (70, 5), (60, 0));

        var host = AgentIdentity.FindAgentHost(table, 100, starts);
        Assert.NotNull(host);
        Assert.Equal(80, host!.Pid);
        Assert.Equal("claude", host.Name);
        Assert.Equal(T0.AddMinutes(10).ToString("o"), host.StartTimeIso);
    }

    [Fact]
    public void Hop0SelfIsNeverCandidate()
    {
        // 조상이 전부 shell/boundary면 자기 자신(non-shell인 vbaloop)이 후보가 되면 안 됨
        var table = Table((100, 90, "vbaloop"), (90, 60, "pwsh"), (60, 1, "explorer"));
        var starts = Starts((100, 50), (90, 40), (60, 0));
        Assert.Null(AgentIdentity.FindAgentHost(table, 100, starts));
    }

    [Fact]
    public void PidReuse_ParentNewerThanChild_StopsChain()
    {
        // claude(80)의 부모 pid 70이 재사용됨(자식보다 최신 StartTime) → 체인 중단, claude까지만
        var table = Table((100, 90, "vbaloop"), (90, 80, "powershell"), (80, 70, "claude"),
                          (70, 60, "code"), (60, 1, "explorer"));
        var starts = Starts((100, 50), (90, 40), (80, 10), (70, 999), (60, 0)); // 70이 훨씬 최신

        var host = AgentIdentity.FindAgentHost(table, 100, starts);
        Assert.Equal(80, host!.Pid); // code(70)는 가짜 부모라 후보 안 됨
    }

    [Fact]
    public void UnreadableAncestorStartTime_StopsChain_KeepsVerifiedCandidate()
    {
        var table = Table((100, 90, "vbaloop"), (90, 80, "claude"), (80, 70, "code"));
        var starts = Starts((100, 50), (90, 10)); // 80의 StartTime 읽기 실패
        var host = AgentIdentity.FindAgentHost(table, 100, starts);
        Assert.Equal(90, host!.Pid); // 검증된 claude까지만
    }

    [Fact]
    public void CycleGuard_Terminates()
    {
        var table = Table((100, 90, "vbaloop"), (90, 100, "node")); // 90 -> 100 사이클
        var starts = Starts((100, 50), (90, 10));
        var host = AgentIdentity.FindAgentHost(table, 100, starts);
        Assert.Equal(90, host!.Pid); // 사이클에서 무한루프 없이 종료
    }

    [Fact]
    public void Resolve_EnvIdTakesPriorityOverProcess()
    {
        var table = Table((100, 90, "vbaloop"), (90, 1, "claude"));
        var starts = Starts((100, 50), (90, 10));
        var id = AgentIdentity.Resolve(@"C:\proj", n => n == "CLAUDE_SESSION_ID" ? "sess-42" : null,
            table, 100, starts);
        Assert.Equal("env", id.Source);
        Assert.Equal("env:CLAUDE_SESSION_ID:sess-42", id.Thread);
        Assert.Null(id.AgentPid);
    }

    private static Func<string, string?> Env(params (string Name, string Value)[] vars)
        => n => vars.FirstOrDefault(v => v.Name == n).Value;

    [Fact]
    public void Resolve_ClaudeCodeSessionId_BindsClaudePid()
    {
        // 고아 백그라운드 체인: vbaloop(100) <- bash(90) <- (죽은 부모) — 계보는 끊겨도 env는 남는다
        var table = Table((100, 90, "vbaloop"), (90, 77, "bash"));
        var starts = Starts((100, 50), (90, 40), (55, 10));
        var id = AgentIdentity.Resolve(@"C:\proj",
            Env(("CLAUDE_CODE_SESSION_ID", "872762ef"), ("CLAUDE_PID", "55")), table, 100, starts);
        Assert.Equal("env", id.Source);
        Assert.Equal("env:CLAUDE_CODE_SESSION_ID:872762ef", id.Thread);
        Assert.Equal(55, id.AgentPid);
        Assert.Equal(T0.AddMinutes(10).ToString("o"), id.AgentStartTime);
    }

    [Fact]
    public void Resolve_ClaudePidUnusable_KeepsEnvIdWithoutPid()
    {
        var table = Table((100, 90, "vbaloop"), (90, 1, "bash"));
        var starts = Starts((100, 50), (90, 40), (56, 60));
        foreach (var pid in new[] { "abc", "0", "100", "99", "56" }) // 비정수/0/자기 자신/죽음/자기보다 최신(PID 재사용)
        {
            var id = AgentIdentity.Resolve(@"C:\proj",
                Env(("CLAUDE_CODE_SESSION_ID", "s"), ("CLAUDE_PID", pid)), table, 100, starts);
            Assert.Equal("env:CLAUDE_CODE_SESSION_ID:s", id.Thread);
            Assert.Null(id.AgentPid);
            Assert.Null(id.AgentStartTime);
        }
    }

    [Fact]
    public void Resolve_ClaudePidAlone_DoesNotBindOtherEnvIds()
    {
        // 사용자 명시 VBALOOP_SESSION_ID는 Claude 세션 수명에 묶지 않는다
        var table = Table((100, 1, "vbaloop"));
        var starts = Starts((100, 50), (55, 10));
        var id = AgentIdentity.Resolve(@"C:\proj",
            Env(("VBALOOP_SESSION_ID", "mine"), ("CLAUDE_CODE_SESSION_ID", "s"), ("CLAUDE_PID", "55")),
            table, 100, starts);
        Assert.Equal("env:VBALOOP_SESSION_ID:mine", id.Thread);
        Assert.Null(id.AgentPid);
    }

    [Fact]
    public void ClaudeDesktopChain_StopsAtSihost()
    {
        // 실측(2026-10-31): powershell <- cmd <- claude(CLI) <- claude(데스크톱) <- sihost <- svchost.
        // sihost가 경계가 아니면 로그아웃까지 사는 sihost가 모든 세션 공용 소유자가 된다.
        var table = Table((100, 90, "vbaloop"), (90, 85, "powershell"), (85, 80, "cmd"),
                          (80, 70, "claude"), (70, 60, "claude"), (60, 50, "sihost"), (50, 1, "svchost"));
        var starts = Starts((100, 50), (90, 45), (85, 44), (80, 40), (70, 10), (60, 1), (50, 0));

        var host = AgentIdentity.FindAgentHost(table, 100, starts);
        Assert.NotNull(host);
        Assert.Equal(70, host!.Pid);
    }

    [Fact]
    public void Resolve_FallsBackToWeak_WhenNoHostFound()
    {
        var table = Table((100, 60, "vbaloop"), (60, 1, "explorer"));
        var starts = Starts((100, 50), (60, 0));
        var id = AgentIdentity.Resolve(@"C:\proj", _ => null, table, 100, starts);
        Assert.Equal("weak", id.Source);
        Assert.StartsWith("local:", id.Thread);
        Assert.EndsWith(AgentIdentity.ProjectKey(@"C:\proj"), id.Thread);
    }

    [Fact]
    public void Resolve_ProcessSource_CarriesPidAndStartTime()
    {
        var table = Table((100, 90, "vbaloop"), (90, 1, "claude"));
        var starts = Starts((100, 50), (90, 10));
        var id = AgentIdentity.Resolve(@"C:\proj", _ => null, table, 100, starts);
        Assert.Equal("process", id.Source);
        Assert.Equal(90, id.AgentPid);
        Assert.Equal(T0.AddMinutes(10).ToString("o"), id.AgentStartTime);
        Assert.Contains(":90:", id.Thread);
    }

    [Fact]
    public void ProjectKey_StableAndCaseInsensitive()
    {
        Assert.Equal(AgentIdentity.ProjectKey(@"C:\Proj\A"), AgentIdentity.ProjectKey(@"c:\proj\a"));
        Assert.True(AgentIdentity.ProjectKey(@"C:\아주\긴\프로젝트\경로\이름").Length <= 12);
    }
}
