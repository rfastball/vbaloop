using System.Diagnostics;
using VbaLoop.Core;
using VbaLoop.Excel;
using Xunit;

namespace VbaLoop.Tests;

public class RecoveryMarkerTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "vbaloop-recovery-" + Guid.NewGuid().ToString("N"));
    private string MarkerPath => Path.Combine(_dir, "_dbg", "recovery.json");

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); } catch { /* best effort */ }
    }

    private ProjectConfig Cfg()
    {
        Directory.CreateDirectory(_dir);
        return new ProjectConfig
        {
            RootDir = _dir,
            WorkbookPath = Path.Combine(_dir, "Book.xlsm"),
            SrcRoot = Path.Combine(_dir, "src"),
            ProjectName = "Book",
        };
    }

    private static SessionMeta Meta(int pid, string start, string? owner = null) => new()
    {
        Pid = pid,
        ExcelStartTime = start,
        Hwnd = 0x1234,
        WorkbookFullName = @"C:\Work\Book.xlsm",
        Owner = owner,
    };

    [Fact]
    public void Missing_LoadsNull()
    {
        Assert.Null(RecoveryMarker.Load(MarkerPath));
    }

    [Fact]
    public void RoundTrip_PreservesFields()
    {
        var m = RecoveryMarker.For(Meta(4242, "2026-10-17T01:02:03.0000000Z", SessionMeta.OwnerExternal),
            "timeout", "test 120s", "test");
        m.Save(MarkerPath);

        var back = RecoveryMarker.Load(MarkerPath)!;
        Assert.False(back.Corrupt);
        Assert.Equal(4242, back.Pid);
        Assert.Equal("2026-10-17T01:02:03.0000000Z", back.ExcelStartTime);
        Assert.True(back.IsExternal);
        Assert.Equal("timeout", back.Reason);
        Assert.Equal("test", back.Command);
        Assert.Equal(@"C:\Work\Book.xlsm", back.Workbook);
        Assert.False(string.IsNullOrEmpty(back.CreatedAt));
    }

    [Fact]
    public void LegacyMetaWithoutOwner_IsTool()
    {
        var m = RecoveryMarker.For(Meta(1, "x"), "poison", null, "push");
        Assert.Equal(SessionMeta.OwnerTool, m.Owner);
        Assert.False(m.IsExternal);
    }

    [Fact]
    public void Corrupt_FileStillQuarantines()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(MarkerPath)!);
        File.WriteAllText(MarkerPath, "{ not json");
        var m = RecoveryMarker.Load(MarkerPath);
        Assert.NotNull(m);
        Assert.True(m!.Corrupt);
    }

    [Fact]
    public void Gate_NoMarker_Passes()
    {
        RecoveryGate.Enforce(Cfg(), null);
    }

    [Fact]
    public void Gate_LiveProcess_ThrowsEnvironmentError_WithHint()
    {
        var cfg = Cfg();
        using var self = Process.GetCurrentProcess();
        string start = self.StartTime.ToUniversalTime().ToString("o");
        RecoveryMarker.For(Meta(self.Id, start), "timeout", null, "test").Save(cfg.RecoveryMarkerPath);

        var e = Assert.Throws<CliError>(() => RecoveryGate.Enforce(cfg, null));
        Assert.True(e.Environment); // exit 3
        Assert.Contains("recovery clear", e.Hint);
        Assert.Contains("--discard", e.Hint);
        Assert.True(File.Exists(cfg.RecoveryMarkerPath)); // 격리 유지
    }

    [Fact]
    public void Gate_ExternalLiveProcess_HintSaysDetachOnly()
    {
        var cfg = Cfg();
        using var self = Process.GetCurrentProcess();
        string start = self.StartTime.ToUniversalTime().ToString("o");
        RecoveryMarker.For(Meta(self.Id, start, SessionMeta.OwnerExternal), "poison", null, "push")
            .Save(cfg.RecoveryMarkerPath);

        var e = Assert.Throws<CliError>(() => RecoveryGate.Enforce(cfg, null));
        Assert.Contains("detach only", e.Hint);
    }

    [Fact]
    public void Gate_ProcessGone_AutoClears()
    {
        var cfg = Cfg();
        using var self = Process.GetCurrentProcess();
        // 같은 PID, 다른 시작 시각 = 기록된 프로세스는 소멸(PID 재사용)
        RecoveryMarker.For(Meta(self.Id, "2000-01-01T00:00:00.0000000Z"), "timeout", null, "test")
            .Save(cfg.RecoveryMarkerPath);

        RecoveryGate.Enforce(cfg, null);
        Assert.False(File.Exists(cfg.RecoveryMarkerPath));
    }

    [Fact]
    public void Gate_CorruptMarker_NeverAutoClears()
    {
        var cfg = Cfg();
        Directory.CreateDirectory(cfg.DbgDir);
        File.WriteAllText(cfg.RecoveryMarkerPath, "garbage");

        var e = Assert.Throws<CliError>(() => RecoveryGate.Enforce(cfg, null));
        Assert.True(e.Environment);
        Assert.Contains("--force", e.Hint);
        Assert.True(File.Exists(cfg.RecoveryMarkerPath));
    }

    [Fact]
    public void Identify_CurrentProcess_Match_AndStates()
    {
        using var self = Process.GetCurrentProcess();
        string start = self.StartTime.ToUniversalTime().ToString("o");
        Assert.Equal(IdentityState.Match, ProcessIdentity.Identify(self.Id, start));
        Assert.Equal(IdentityState.Gone, ProcessIdentity.Identify(self.Id, "2000-01-01T00:00:00.0000000Z"));
        Assert.Equal(IdentityState.Unknown, ProcessIdentity.Identify(self.Id, ""));
        Assert.Equal(IdentityState.Gone, ProcessIdentity.Identify(0, start));
    }

    [Fact]
    public void ProbeParse_RoundTrip_AndGarbageIsUnresponsive()
    {
        var r = SessionProbe.Parse("noise {\"kind\":\"Responsive\",\"workbooks\":[\"C:\\\\a.xlsm\"],\"detail\":null}");
        Assert.Equal(ProbeKind.Responsive, r.Kind);
        Assert.Equal(new[] { @"C:\a.xlsm" }, r.Workbooks);

        Assert.Equal(ProbeKind.Unresponsive, SessionProbe.Parse("").Kind);
        Assert.Equal(ProbeKind.Unresponsive, SessionProbe.Parse("{\"kind\":\"???\"}").Kind);
        Assert.Equal(ProbeKind.Unresponsive, SessionProbe.Parse("{oops").Kind);
    }
}
