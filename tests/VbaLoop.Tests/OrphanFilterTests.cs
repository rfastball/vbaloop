using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>doctor --kill-orphans 대상 판정(설계 §11 L4): 현재 세션 + /automation + 비격리만 kill.</summary>
public class OrphanFilterTests
{
    private const int MySession = 2;
    private const string AutomationCmd = "\"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE\" /automation -Embedding";

    private static OrphanCandidate Candidate(
        string? cmd = AutomationCmd, int session = MySession, bool window = false,
        bool live = false, bool quarantined = false, bool tool = false)
        => new(1234, "2026-10-17T00:00:00.0000000Z", window, session, cmd, live, quarantined, tool);

    [Fact]
    public void AutomationInCurrentSession_IsKillable()
    {
        var v = OrphanFilter.Classify(Candidate(), MySession);
        Assert.True(v.Killable);
    }

    [Fact]
    public void ToolLaunched_IsKillable_AndSaysSo()
    {
        var v = OrphanFilter.Classify(Candidate(tool: true), MySession);
        Assert.True(v.Killable);
        Assert.Contains("vbaloop", v.Reason);
    }

    [Fact]
    public void OleEmbedded_WithoutAutomationToken_IsReportedOnly()
    {
        // Word/PowerPoint에 임베드된 Excel 개체 활성화는 -Embedding만 붙는다
        var v = OrphanFilter.Classify(Candidate(cmd: "\"C:\\Office16\\EXCEL.EXE\" -Embedding"), MySession);
        Assert.False(v.Killable);
        Assert.Contains("OLE", v.Reason);
    }

    [Fact]
    public void OtherWindowsSession_IsReportedOnly()
    {
        var v = OrphanFilter.Classify(Candidate(session: 5), MySession);
        Assert.False(v.Killable);
        Assert.Contains("RDS", v.Reason);
    }

    [Fact]
    public void UnreadableCommandLine_IsReportedOnly()
    {
        Assert.False(OrphanFilter.Classify(Candidate(cmd: null), MySession).Killable);
    }

    [Fact]
    public void Quarantined_NeverKillable_EvenIfToolLaunched()
    {
        var v = OrphanFilter.Classify(Candidate(quarantined: true, tool: true), MySession);
        Assert.False(v.Killable);
        Assert.Contains("격리", v.Reason);
    }

    [Fact]
    public void WindowedOrLiveSession_NotKillable()
    {
        Assert.False(OrphanFilter.Classify(Candidate(window: true), MySession).Killable);
        Assert.False(OrphanFilter.Classify(Candidate(live: true), MySession).Killable);
    }

    [Theory]
    [InlineData("EXCEL.EXE /automation -Embedding", true)]
    [InlineData("EXCEL.EXE /AUTOMATION", true)]
    [InlineData("EXCEL.EXE -Embedding", false)]
    [InlineData("\"C:\\tools\\/automation\\EXCEL.EXE\" -Embedding", false)] // 경로 안 부분 문자열은 토큰이 아니다
    [InlineData("EXCEL.EXE /automationx", false)]
    [InlineData("", false)]
    public void IsAutomationLaunch_TokenMatch(string cmd, bool expected)
    {
        Assert.Equal(expected, OrphanFilter.IsAutomationLaunch(cmd));
    }

    [Fact]
    public void CommandLine_OfCurrentProcess_IsReadable()
    {
        string? cmd = ProcessCommandLine.Get(Environment.ProcessId);
        Assert.False(string.IsNullOrWhiteSpace(cmd));
    }

    [Fact]
    public void CommandLine_OfMissingProcess_IsNull()
    {
        Assert.Null(ProcessCommandLine.Get(0));
    }
}
