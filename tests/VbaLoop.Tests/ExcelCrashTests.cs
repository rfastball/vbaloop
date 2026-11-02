using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ExcelCrashTests
{
    private static readonly DateTime CrashedStart = new(2026, 11, 2, 11, 37, 46, DateTimeKind.Utc);

    [Fact]
    public void Classify_AliveProcess_IsAlive_RegardlessOfSigns()
    {
        Assert.Equal(ExcelExitKind.Alive, ExcelCrash.Classify(false, null, crashSuspectError: true, relaunched: true));
    }

    [Fact]
    public void Classify_ExitCodeZeroAfterDispException_IsCrash()
    {
        // 실측: Office 크래시 핸들러를 거친 Excel은 종료 코드 0 — 오류 징후가 크래시를 가른다
        Assert.Equal(ExcelExitKind.Crashed, ExcelCrash.Classify(true, 0, crashSuspectError: true));
    }

    [Fact]
    public void Classify_RelaunchedChild_IsCrash()
    {
        Assert.Equal(ExcelExitKind.Crashed, ExcelCrash.Classify(true, 0, crashSuspectError: false, relaunched: true));
    }

    [Fact]
    public void Classify_NonZeroExitCode_IsCrash()
    {
        Assert.Equal(ExcelExitKind.Crashed, ExcelCrash.Classify(true, unchecked((int)0xC0000005), crashSuspectError: false));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(null)]
    public void Classify_NoCrashSigns_IsExited(int? exitCode)
    {
        Assert.Equal(ExcelExitKind.Exited, ExcelCrash.Classify(true, exitCode, crashSuspectError: false));
    }

    [Fact]
    public void FindRelaunched_PicksExcelChildrenOfCrashedPid()
    {
        var procs = new[]
        {
            new ProcessTree.ProcEntry(100, 4, "excel"),          // 크래시한 Excel 자신
            new ProcessTree.ProcEntry(200, 100, "excel"),        // 재기동 Excel
            new ProcessTree.ProcEntry(300, 100, "werfault"),     // 같은 부모의 다른 프로세스
            new ProcessTree.ProcEntry(400, 1892, "excel"),       // 무관한 Excel(다른 세션의 일회용 인스턴스)
        };
        var found = ExcelCrash.FindRelaunched(procs, 100, CrashedStart, _ => CrashedStart.AddSeconds(1));
        Assert.Equal(new[] { 200 }, found);
    }

    [Fact]
    public void FindRelaunched_IgnoresChildOlderThanCrashedProcess()
    {
        // 부모 PID가 재사용된 경우: 크래시한 프로세스보다 먼저 시작한 "자식"은 무관한 프로세스다
        var procs = new[] { new ProcessTree.ProcEntry(200, 100, "excel") };
        var found = ExcelCrash.FindRelaunched(procs, 100, CrashedStart, _ => CrashedStart.AddMinutes(-5));
        Assert.Empty(found);
    }

    [Fact]
    public void FindRelaunched_IgnoresUnknownStartTime()
    {
        var procs = new[] { new ProcessTree.ProcEntry(200, 100, "excel") };
        Assert.Empty(ExcelCrash.FindRelaunched(procs, 100, CrashedStart, _ => null));
    }

    [Fact]
    public void Detail_OmitsZeroExitCode_ShowsNonZero()
    {
        Assert.Equal("Excel crashed while running modCrash.Crash (PID 7 exited)", ExcelCrash.Detail("modCrash.Crash", 7, 0));
        Assert.Equal("Excel crashed while running m.P (PID 7 exited with code 0xC0000005)",
            ExcelCrash.Detail("m.P", 7, unchecked((int)0xC0000005)));
    }

    [Fact]
    public void RelaunchNote_EmptyWithoutRelaunch_AndMentionsLockWhenPresent()
    {
        Assert.Equal("", ExcelCrash.RelaunchNote(Array.Empty<int>(), @"C:\p\Book.xlsm", workbookLocked: true));

        string locked = ExcelCrash.RelaunchNote(new[] { 81956 }, @"C:\p\Book.xlsm", workbookLocked: true);
        Assert.Contains("PID 81956", locked);
        Assert.Contains("~$Book.xlsm", locked);
        Assert.Contains("push is refused", locked);

        string unlocked = ExcelCrash.RelaunchNote(new[] { 81956 }, @"C:\p\Book.xlsm", workbookLocked: false);
        Assert.Contains("may reopen Book.xlsm", unlocked);
    }

    [Fact]
    public void Hint_WithRelaunch_ForbidsKill_AndKeepsRestartStep()
    {
        string hint = ExcelCrash.Hint(new[] { 81956 }, "rerun the same command");
        Assert.StartsWith("Do NOT kill the relaunched Excel (PID 81956)", hint);
        Assert.EndsWith("rerun the same command", hint);

        string plain = ExcelCrash.Hint(Array.Empty<int>(), "session start");
        Assert.DoesNotContain("Do NOT kill", plain);
        Assert.EndsWith("session start", plain);
    }
}
