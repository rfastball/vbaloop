using System.Text;
using System.Text.Json;
using VbaLoop.Cli;
using VbaLoop.Commands;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>실행 취소(설계 §14): 실행 잠금, 취소 표시, 취소 레코드 판독, 막는 대화상자 판정, 대기 한도.</summary>
public class RunCancelTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "vbaloop-runcancel-" + Guid.NewGuid().ToString("N"));

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

    private ActiveRun Record(string runId, int runnerPid, string runnerStart) => new()
    {
        RunId = runId, Command = "run", Target = "modA.Go", RunnerPid = runnerPid, RunnerStart = runnerStart,
        ExcelPid = 4242, StartedUtc = DateTime.UtcNow.ToString("O"),
        CancelFlag = RunCancel.PathFor(Path.Combine(_dir, "Book.xlsm"), runId),
    };

    private static string SelfStart => ProcessIdentity.GetStartTimeIso(Environment.ProcessId) ?? "";

    // ---- 취소 표시 ----

    [Fact]
    public void CancelFlag_IsPerRunUnderWorkbookDbg()
    {
        Assert.Equal(@"C:\Work\_dbg\cancel.r1", RunCancel.PathFor(@"C:\Work\Book.xlsm", "r1"));
    }

    [Fact]
    public void CancelFlag_RequestThenClear()
    {
        string flag = RunCancel.PathFor(Path.Combine(_dir, "Book.xlsm"), "r1");
        Assert.True(RunCancel.Request(flag));
        Assert.True(File.Exists(flag));
        Assert.True(RunCancel.Request(flag)); // 멱등
        RunCancel.Clear(flag);
        Assert.False(File.Exists(flag));
    }

    [Fact]
    public void SweepOld_RemovesOnlyStaleFlags()
    {
        string dbg = Path.Combine(_dir, "_dbg");
        string oldFlag = RunCancel.PathFor(Path.Combine(_dir, "Book.xlsm"), "old");
        string newFlag = RunCancel.PathFor(Path.Combine(_dir, "Book.xlsm"), "new");
        RunCancel.Request(oldFlag);
        RunCancel.Request(newFlag);
        File.SetLastWriteTimeUtc(oldFlag, DateTime.UtcNow.AddDays(-2));
        RunCancel.SweepOld(dbg, TimeSpan.FromDays(1));
        Assert.False(File.Exists(oldFlag));
        Assert.True(File.Exists(newFlag));
    }

    // ---- 실행 잠금 ----

    [Fact]
    public void RunLock_HeldByAnotherThread_RefusesWithDescription()
    {
        var cfg = Cfg();
        using var held = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        bool gotLock = false;
        var holder = new Thread(() =>
        {
            using var rl = RunLock.Hold(cfg, Record("r1", Environment.ProcessId, SelfStart));
            gotLock = rl != null;
            held.Set();
            release.Wait();
        });
        holder.Start();
        held.Wait();
        try
        {
            Assert.True(gotLock);
            var ex = Assert.Throws<CliError>(() => RunLock.ThrowIfActive(cfg));
            Assert.Contains("still executing a macro", ex.Message);
            Assert.Contains("modA.Go", ex.Message);
        }
        finally
        {
            release.Set();
            holder.Join();
        }
        // 러너가 끝나면 잠금과 기록이 함께 사라진다
        Assert.False(File.Exists(RunLock.RecordPath(cfg)));
        RunLock.ThrowIfActive(cfg);
    }

    [Fact]
    public void RunLock_AbandonedByDeadThread_CountsAsFree()
    {
        var cfg = Cfg();
        var t = new Thread(() => RunLock.Hold(cfg, Record("r1", Environment.ProcessId, SelfStart))); // Dispose 없이 종료
        t.Start();
        t.Join();
        RunLock.ThrowIfActive(cfg); // 버려진 뮤텍스 = 비어 있음
    }

    [Fact]
    public void Orphan_DeadRunnerAndDeadExcel_RecordAndFlagRemoved()
    {
        var cfg = Cfg();
        Directory.CreateDirectory(cfg.DbgDir);
        var rec = Record("r9", int.MaxValue - 7, "2020-01-01T00:00:00.0000000Z");
        rec.ExcelPid = int.MaxValue - 9;
        File.WriteAllText(RunLock.RecordPath(cfg), JsonSerializer.Serialize(rec));
        RunLock.ThrowIfActive(cfg); // Excel도 없다 - 기다릴 것 없이 통과
        Assert.False(File.Exists(rec.CancelFlag)); // 멈출 매크로가 없으니 표시도 남기지 않는다
        Assert.False(File.Exists(RunLock.RecordPath(cfg)));
    }

    // 트리 kill: 러너는 죽었고 Excel(여기선 이 테스트 프로세스)은 살아 있다
    private ActiveRun TreeKilled(string runId)
    {
        var rec = Record(runId, int.MaxValue - 7, "2020-01-01T00:00:00.0000000Z");
        rec.ExcelPid = Environment.ProcessId;
        rec.Jsonl = Path.Combine(_dir, "_dbg", "Book.jsonl");
        return rec;
    }

    [Fact]
    public void Orphan_LiveExcel_RefusesUntilMacroStops()
    {
        var cfg = Cfg();
        Directory.CreateDirectory(cfg.DbgDir);
        var rec = TreeKilled("r7");
        File.WriteAllText(RunLock.RecordPath(cfg), JsonSerializer.Serialize(rec));

        var ex = Assert.Throws<CliError>(() => RunLock.ThrowIfActive(cfg));
        Assert.Contains("was killed while its macro was still running", ex.Message);
        Assert.True(File.Exists(rec.CancelFlag));
        Assert.NotNull(RunLock.Load(RunLock.RecordPath(cfg))!.OrphanCancelUtc);
        Assert.Throws<CliError>(() => RunLock.ThrowIfActive(cfg)); // 아직 멈춤 근거 없음

        // LogShim이 이 runId의 취소 기록을 남겼다 = 매크로가 멈췄다
        File.WriteAllText(rec.Jsonl, "{\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r7\"}\r\n");
        RunLock.ThrowIfActive(cfg);
        Assert.False(File.Exists(RunLock.RecordPath(cfg)));
    }

    [Fact]
    public void Orphan_LiveExcel_GivesUpAfterWait()
    {
        var cfg = Cfg();
        Directory.CreateDirectory(cfg.DbgDir);
        var rec = TreeKilled("r8");
        rec.OrphanCancelUtc = DateTime.UtcNow.AddMilliseconds(-RunLock.OrphanWaitMs - 1000).ToString("O");
        File.WriteAllText(RunLock.RecordPath(cfg), JsonSerializer.Serialize(rec));
        RunLock.ThrowIfActive(cfg); // 상한이 지났다 - 평소의 세션 복구에 맡긴다
        Assert.False(File.Exists(RunLock.RecordPath(cfg)));
    }

    [Fact]
    public void ClearAll_RemovesFlagsUnlessARunIsRecorded()
    {
        var cfg = Cfg();
        string flag = RunCancel.PathFor(cfg.WorkbookPath, "old");
        RunCancel.Request(flag);
        File.WriteAllText(RunLock.RecordPath(cfg), JsonSerializer.Serialize(Record("r1", 1, "")));
        RunCancel.ClearAll(cfg);
        Assert.True(File.Exists(flag)); // 실행 중(또는 고아 취소 대기) - 그 표시가 필요하다
        File.Delete(RunLock.RecordPath(cfg));
        RunCancel.ClearAll(cfg);
        Assert.False(File.Exists(flag));
    }

    [Fact]
    public void Orphan_RecordOfLiveRunner_IsLeftAlone()
    {
        var cfg = Cfg();
        Directory.CreateDirectory(cfg.DbgDir);
        var rec = Record("r9", Environment.ProcessId, SelfStart);
        File.WriteAllText(RunLock.RecordPath(cfg), JsonSerializer.Serialize(rec));
        RunLock.ThrowIfActive(cfg);
        Assert.False(File.Exists(rec.CancelFlag));
        Assert.True(File.Exists(RunLock.RecordPath(cfg)));
    }

    [Fact]
    public void Orphan_CorruptRecord_IsRemoved()
    {
        var cfg = Cfg();
        Directory.CreateDirectory(cfg.DbgDir);
        File.WriteAllText(RunLock.RecordPath(cfg), "{not json");
        RunLock.ThrowIfActive(cfg);
        Assert.False(File.Exists(RunLock.RecordPath(cfg)));
    }

    [Fact]
    public void RunLock_DisposeKeepsAnotherRunsRecord()
    {
        var cfg = Cfg();
        var rl = RunLock.Hold(cfg, Record("r1", Environment.ProcessId, SelfStart))!;
        File.WriteAllText(RunLock.RecordPath(cfg), JsonSerializer.Serialize(Record("r2", 1, "")));
        rl.Dispose();
        Assert.Equal("r2", RunLock.Load(RunLock.RecordPath(cfg))!.RunId);
    }

    [Fact]
    public void Forget_RemovesOnlyThatRunsRecord()
    {
        var cfg = Cfg();
        Directory.CreateDirectory(cfg.DbgDir);
        File.WriteAllText(RunLock.RecordPath(cfg), JsonSerializer.Serialize(Record("r1", 1, "")));
        RunLock.Forget(cfg, "r2");
        Assert.True(File.Exists(RunLock.RecordPath(cfg)));
        RunLock.Forget(cfg, "r1");
        Assert.False(File.Exists(RunLock.RecordPath(cfg)));
    }

    [Fact]
    public void ActiveRun_RecordFromOlderBuildStillLoads()
    {
        var cfg = Cfg();
        Directory.CreateDirectory(cfg.DbgDir);
        File.WriteAllText(RunLock.RecordPath(cfg), "{\"runId\":\"r1\",\"cancelSupported\":false}");
        Assert.Equal("r1", RunLock.Load(RunLock.RecordPath(cfg))!.RunId);
    }

    // ---- LogShim 취소 지원 판정 ----

    [Theory]
    [InlineData("Public Sub BeginRun(runId As String)\r\n    mRunId = runId\r\nEnd Sub", true)]
    [InlineData("Option Explicit\r\n  public sub beginrun (ByVal runId As String)\r\nEnd Sub", true)]
    [InlineData("Public Sub LogD(level As String, msg As String)\r\nEnd Sub", false)]
    [InlineData("' Public Sub BeginRun(runId As String)", false)]
    [InlineData("Private Sub BeginRun(runId As String)", false)]
    [InlineData("Public Sub BeginRunner(x As String)", false)]
    public void ShimSupportsCancel_NeedsPublicBeginRun(string body, bool expected)
        => Assert.Equal(expected, RunCancel.ShimSupportsCancel(body));

    [Fact]
    public void ShimSupportsCancel_CanonicalLogShim()
    {
        var asm = typeof(RunCancel).Assembly;
        using var s = asm.GetManifestResourceStream("VbaLoop.assets.LogShim.bas")!;
        Assert.True(RunCancel.ShimSupportsCancel(new StreamReader(s).ReadToEnd()));
    }

    // ---- 중간에 끊긴 매크로의 뒷정리 ----

    private static readonly AppSettings Baseline = new(EnableEvents: true, Calculation: -4105, StatusBar: null,
        Interactive: true, Cursor: -4143);

    [Fact]
    public void AppSettings_UnchangedReportsNothing()
        => Assert.Empty(AppSettings.Changed(Baseline, Baseline with { }));

    [Fact]
    public void AppSettings_ReportsEachChangedSetting()
    {
        var after = Baseline with { EnableEvents = false, Calculation = -4135, StatusBar = "Working 3/10", Interactive = false, Cursor = 2 };
        Assert.Equal(new[] { "EnableEvents", "Calculation", "StatusBar", "Interactive", "Cursor" }, AppSettings.Changed(Baseline, after));
        Assert.Equal(new[] { "StatusBar" }, AppSettings.Changed(Baseline with { StatusBar = "x" }, Baseline));
    }

    [Theory]
    [InlineData(false, null)]
    [InlineData("FALSE", null)] // 문구를 쓴 뒤 False로 돌려준 Excel이 돌려주는 값
    [InlineData("false", null)]
    [InlineData("Working 3/10", "Working 3/10")]
    [InlineData("", "")]
    public void StatusBarText_TreatsFalseStringAsExcelInControl(object? value, string? expected)
        => Assert.Equal(expected, AppSettings.StatusBarText(value));

    [Fact]
    public void NewWorkbooks_OnlyThoseOpenedDuringTheRun()
    {
        var before = new[] { @"C:\p\Book.xlsm", "PERSONAL.XLSB" };
        var after = new[] { @"c:\P\BOOK.xlsm", "PERSONAL.XLSB", "Book1", @"C:\data\in.xlsx" };
        Assert.Equal(new[] { "Book1", @"C:\data\in.xlsx" }, RunAftermath.NewWorkbooks(before, after));
        Assert.Empty(RunAftermath.NewWorkbooks(before, new[] { @"C:\p\Book.xlsm" })); // 닫힌 기존 문서는 다루지 않는다
    }

    [Fact]
    public void Aftermath_EmptyAndSummary()
    {
        var none = Array.Empty<string>();
        Assert.True(new RunAftermath(none, none, none, none).IsEmpty);
        var a = new RunAftermath(new[] { "EnableEvents", "Calculation" }, new[] { "Book1" }, none, none);
        Assert.False(a.IsEmpty);
        Assert.Equal("restored Application.EnableEvents, Application.Calculation to the value before the run; "
                     + "closed without saving 1 workbook(s) the macro left open (Book1)", a.Summary);
        var kept = new RunAftermath(none, none, new[] { "Book1" }, new[] { "Application.Cursor" });
        Assert.Contains("left 1 workbook(s) open (Book1) - not closed (a kept session, or a saved file with unsaved changes)", kept.Summary);
        Assert.Contains("could not undo: Application.Cursor", kept.Summary);
    }

    [Fact]
    public void Aftermath_SerializesWithoutIsEmpty()
    {
        var none = Array.Empty<string>();
        string json = JsonSerializer.Serialize(new RunAftermath(new[] { "Cursor" }, none, none, none), Output.JsonOptions);
        Assert.DoesNotContain("isEmpty", json);
        Assert.Contains("\"summary\"", json);
        Assert.Contains("\"restored\"", json);
    }

    // ---- LogShim 취소 레코드 ----

    [Theory]
    [InlineData("{\"ts\":\"2026-11-03T10:00:00\",\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r1\"}", true)]
    [InlineData("  {\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r1\"}  ", true)]
    [InlineData("{\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r2\"}", false)]   // 다른 실행
    [InlineData("{\"kind\":\"J\",\"event\":\"cancelled\",\"runId\":\"r1\"}", false)]   // 사용자 LogEvent
    [InlineData("{\"kind\":\"D\",\"msg\":\"cancelled\"}", false)]
    [InlineData("cancelled", false)]
    [InlineData("{\"kind\":\"X\",\"event\":\"cancelled\",", false)]                   // 잘린 줄
    public void IsCancelRecord_MatchesOnlyThisRunsMarker(string line, bool expected)
    {
        Assert.Equal(expected, RunnerMain.IsCancelRecord(line, "r1"));
    }

    [Fact]
    public void SpanCancelled_LooksOnlyAfterOffset()
    {
        Directory.CreateDirectory(_dir);
        string jsonl = Path.Combine(_dir, "Book.jsonl");
        File.WriteAllText(jsonl, "{\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r1\"}\r\n", new UTF8Encoding(false));
        long offset = new FileInfo(jsonl).Length;
        Assert.False(RunnerMain.SpanCancelled(jsonl, offset, "r1"));
        File.AppendAllText(jsonl, "{\"kind\":\"D\",\"msg\":\"x\"}\r\n{\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r1\"}\r\n");
        Assert.True(RunnerMain.SpanCancelled(jsonl, offset, "r1"));
        Assert.False(RunnerMain.SpanCancelled(Path.Combine(_dir, "missing.jsonl"), 0, "r1"));
    }

    [Fact]
    public void SpanCancelled_LogShrunkByReset_ScansFromStart()
    {
        Directory.CreateDirectory(_dir);
        string jsonl = Path.Combine(_dir, "Book.jsonl");
        File.WriteAllText(jsonl, "{\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r1\"}\r\n");
        Assert.True(RunnerMain.SpanCancelled(jsonl, 1_000_000, "r1"));
    }

    [Fact]
    public void ScanForCancel_ReadsOnlyCompleteLinesAndAdvances()
    {
        Directory.CreateDirectory(_dir);
        string jsonl = Path.Combine(_dir, "Book.jsonl");
        File.WriteAllText(jsonl, "{\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r1\"}\r\n", new UTF8Encoding(false));
        long pos = new FileInfo(jsonl).Length; // 실행 전부터 있던 기록은 보지 않는다
        Assert.False(RunCancel.ScanForCancel(jsonl, ref pos, "r1"));
        File.AppendAllText(jsonl, "{\"kind\":\"D\",\"msg\":\"x\"}\r\n{\"kind\":\"X\",\"event\":\"cance");
        long afterFirstLine = pos + "{\"kind\":\"D\",\"msg\":\"x\"}\r\n".Length;
        Assert.False(RunCancel.ScanForCancel(jsonl, ref pos, "r1"));
        Assert.Equal(afterFirstLine, pos); // 쓰는 중인 줄은 다음에 다시 읽는다
        File.AppendAllText(jsonl, "lled\",\"runId\":\"r2\"}\r\n");
        Assert.False(RunCancel.ScanForCancel(jsonl, ref pos, "r1")); // 다른 실행의 기록
        File.AppendAllText(jsonl, "{\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r1\"}\r\n");
        Assert.True(RunCancel.ScanForCancel(jsonl, ref pos, "r1"));
        Assert.Equal(new FileInfo(jsonl).Length, pos);
    }

    [Fact]
    public void ScanForCancel_LogShrunkOrMissing()
    {
        Directory.CreateDirectory(_dir);
        string jsonl = Path.Combine(_dir, "Book.jsonl");
        long pos = 1_000_000;
        Assert.False(RunCancel.ScanForCancel(jsonl, ref pos, "r1"));
        File.WriteAllText(jsonl, "{\"kind\":\"X\",\"event\":\"cancelled\",\"runId\":\"r1\"}\n");
        Assert.True(RunCancel.ScanForCancel(jsonl, ref pos, "r1")); // Reset이 잘라냈다 - 처음부터
    }

    // ---- 막는 대화상자 ----

    [Fact]
    public void BlockingDialog_ReportedOnlyAfterPersisting()
    {
        var t = new BlockingDialogTracker(Array.Empty<IntPtr>(), persistMs: 5000);
        var d = (IntPtr)0x100;
        Assert.Null(t.Observe(new[] { d }, 0));
        Assert.Null(t.Observe(new[] { d }, 4999));
        Assert.Equal(d, t.Observe(new[] { d }, 5000));
    }

    [Fact]
    public void BlockingDialog_TransientWindowResets()
    {
        var t = new BlockingDialogTracker(Array.Empty<IntPtr>(), persistMs: 5000);
        var d = (IntPtr)0x100;
        Assert.Null(t.Observe(new[] { d }, 0));
        Assert.Null(t.Observe(Array.Empty<IntPtr>(), 3000)); // 워치독이 닫았다
        Assert.Null(t.Observe(new[] { d }, 4000));          // 재사용된 핸들 - 새로 센다
        Assert.Null(t.Observe(new[] { d }, 8999));
        Assert.Equal(d, t.Observe(new[] { d }, 9000));
    }

    [Fact]
    public void BlockingDialog_PreexistingIgnored()
    {
        var pre = (IntPtr)0x200;
        var t = new BlockingDialogTracker(new[] { pre }, persistMs: 1000);
        Assert.Null(t.Observe(new[] { pre }, 0));
        Assert.Null(t.Observe(new[] { pre }, 60_000));
    }

    // ---- 대기 한도·진행 표시 ----

    [Fact]
    public void RunLimits_DefaultIsSafetyCapOnly()
    {
        var l = new RunLimits(null, 3600);
        Assert.Equal(3_600_000, l.DeadlineMs);
        Assert.Equal("safety cap reached ([run] maxSeconds = 3600s)", l.Describe());
    }

    [Fact]
    public void RunLimits_TimeoutBelowCapWins()
    {
        var l = new RunLimits(180, 3600);
        Assert.Equal(180_000, l.DeadlineMs);
        Assert.Equal("timeout (180s)", l.Describe());
    }

    [Fact]
    public void RunLimits_TimeoutAboveCapIsCapped()
    {
        var l = new RunLimits(7200, 3600);
        Assert.Equal(3_600_000, l.DeadlineMs);
        Assert.StartsWith("safety cap", l.Describe());
    }

    [Theory]
    [InlineData(0, "0s")]
    [InlineData(59, "59s")]
    [InlineData(61, "1m01s")]
    [InlineData(3725, "1h02m05s")]
    public void FormatElapsed(int seconds, string expected)
    {
        Assert.Equal(expected, RunProgress.FormatElapsed(TimeSpan.FromSeconds(seconds)));
    }

    [Fact]
    public void ProgressLine_IncludesTargetRecordsAndQuietTime()
    {
        Assert.Equal("vbaloop: run modA.Go still running - 1m30s elapsed, 12 LogShim record(s), last 45s ago",
            RunProgress.Line("run", "modA.Go", TimeSpan.FromSeconds(90), 12, TimeSpan.FromSeconds(45)));
        Assert.Equal("vbaloop: test Test_* still running - 30s elapsed, 0 LogShim record(s)",
            RunProgress.Line("test", "Test_*", TimeSpan.FromSeconds(30), 0, null));
    }

    // ---- 설정·인자 ----

    private ProjectConfig LoadToml(string toml)
    {
        Directory.CreateDirectory(_dir);
        string path = Path.Combine(_dir, "vbaloop.toml");
        File.WriteAllText(path, toml);
        return ProjectConfig.Load(path, _dir);
    }

    [Fact]
    public void RunMaxSeconds_DefaultsAndParses()
    {
        Assert.Equal(ProjectConfig.DefaultRunMaxSeconds, LoadToml("[excel]\nworkbook = \"B.xlsm\"").RunMaxSeconds);
        Assert.Equal(900, LoadToml("[excel]\nworkbook = \"B.xlsm\"\n[run]\nmaxSeconds = 900").RunMaxSeconds);
    }

    [Theory]
    [InlineData("maxSeconds = 0")]
    [InlineData("maxSeconds = -5")]
    [InlineData("maxSeconds = \"600\"")]
    [InlineData("timeout = 60")]
    public void RunMaxSeconds_Invalid_FailsLoudly(string line)
    {
        Assert.Throws<CliError>(() => LoadToml("[excel]\nworkbook = \"B.xlsm\"\n[run]\n" + line));
    }

    [Fact]
    public void ParentPid_ParsedAndConsumesValue()
    {
        var a = CliArgs.Parse(new[] { "__runner", "--parent-pid", "1234", "--json" });
        Assert.Equal(1234, a.ParentPid);
        Assert.Equal(new[] { "__runner" }, a.Positional);
        Assert.Throws<CliError>(() => CliArgs.Parse(new[] { "__runner", "--parent-pid" }));
        Assert.Throws<CliError>(() => CliArgs.Parse(new[] { "__runner", "--parent-pid", "0" }));
        Assert.Throws<CliError>(() => CliArgs.Parse(new[] { "__runner", "--parent-pid", "abc" }));
    }

    // ---- LogShim 자산 ----

    [Fact]
    public void LogShimAsset_HasCooperativeCancelEntryPoints()
    {
        using var stream = typeof(SampleProject).Assembly.GetManifestResourceStream("VbaLoop.assets.LogShim.bas")!;
        string text = new StreamReader(stream, new UTF8Encoding(false)).ReadToEnd();
        Assert.Contains("Public Sub BeginRun(runId As String)", text);
        Assert.Contains("Public Sub Poll()", text);
        Assert.Contains("\"\"cancelled\"\"", text); // VBA 문자열 리터럴 안의 JSON 키
        // 확인 지점: 모든 기록 경로가 취소를 본다
        foreach (string proc in new[] { "Public Sub LogD(", "Public Sub LogJ(", "Public Sub AssertEq(" })
        {
            int at = text.IndexOf(proc, StringComparison.Ordinal);
            Assert.True(at >= 0, proc);
            int end = text.IndexOf("End Sub", at, StringComparison.Ordinal);
            Assert.Contains("CheckCancel", text[at..end]);
        }
    }
}
