using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class WorkbookRollbackTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "vbaloop-rollbacktest", Guid.NewGuid().ToString("N"));

    public WorkbookRollbackTests() => Directory.CreateDirectory(_dir);

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); } catch { }
    }

    private string WriteFile(string name, string content)
    {
        string p = Path.Combine(_dir, name);
        File.WriteAllText(p, content);
        return p;
    }

    // ---- 세션 차단 규칙 ----

    private static SessionMeta Meta(bool poisoned = false)
        => new() { Pid = 4242, ExcelStartTime = "2026-10-17T03:00:00.0000000Z", Poisoned = poisoned };

    [Fact]
    public void Block_NoMeta_Allows()
        => Assert.Null(WorkbookRollback.SessionBlockReason(null, (_, _) => throw new InvalidOperationException()));

    [Fact]
    public void Block_ActiveProcess_Refuses()
    {
        string? why = WorkbookRollback.SessionBlockReason(Meta(), (pid, start) => pid == 4242 && start != null);
        Assert.NotNull(why);
        Assert.Contains("4242", why);
        Assert.Contains("active", why);
    }

    [Fact]
    public void Block_PoisonedButAlive_StillRefuses()
    {
        // poisoned 세션의 Excel도 살아 있으면 파일을 물고 있다 — session stop(kill)이 먼저
        string? why = WorkbookRollback.SessionBlockReason(Meta(poisoned: true), (_, _) => true);
        Assert.NotNull(why);
        Assert.Contains("poisoned", why);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Block_DeadProcess_Allows(bool poisoned)
        => Assert.Null(WorkbookRollback.SessionBlockReason(Meta(poisoned), (_, _) => false));

    // ---- 원자 교체 ----

    [Fact]
    public void Replace_SwapsContentAndLeavesNoTemp()
    {
        string backup = WriteFile("backup.xlsm", "OLD-GOOD");
        string wb = WriteFile("Book.xlsm", "NEW-BAD");
        WorkbookRollback.ReplaceWorkbook(backup, BackupStore.HashFile(backup), wb);
        Assert.Equal("OLD-GOOD", File.ReadAllText(wb));
        Assert.Equal("OLD-GOOD", File.ReadAllText(backup)); // 백업은 복사만, 소비되지 않음
        Assert.Equal(new[] { "Book.xlsm", "backup.xlsm" },
            Directory.GetFiles(_dir).Select(Path.GetFileName).OrderBy(n => n, StringComparer.Ordinal));
    }

    [Fact]
    public void Replace_MissingWorkbook_Moves()
    {
        string backup = WriteFile("backup.xlsm", "OLD");
        string wb = Path.Combine(_dir, "Book.xlsm");
        WorkbookRollback.ReplaceWorkbook(backup, null, wb);
        Assert.Equal("OLD", File.ReadAllText(wb));
    }

    [Fact]
    public void Replace_HashMismatch_RefusesAndKeepsWorkbook()
    {
        string backup = WriteFile("backup.xlsm", "CORRUPTED");
        string wb = WriteFile("Book.xlsm", "CURRENT");
        Assert.Throws<CliError>(() => WorkbookRollback.ReplaceWorkbook(backup, new string('0', 64), wb));
        Assert.Equal("CURRENT", File.ReadAllText(wb));
    }

    [Fact]
    public void Replace_LockedWorkbook_EnvironmentErrorAndWorkbookIntact()
    {
        string backup = WriteFile("backup.xlsm", "OLD");
        string wb = WriteFile("Book.xlsm", "CURRENT");
        using (new FileStream(wb, FileMode.Open, FileAccess.Read, FileShare.Read)) // Excel처럼 공유 쓰기 불가 오픈
        {
            var ex = Assert.Throws<CliError>(() => WorkbookRollback.ReplaceWorkbook(backup, null, wb));
            Assert.True(ex.Environment);
        }
        Assert.Equal("CURRENT", File.ReadAllText(wb));
        Assert.Equal(2, Directory.GetFiles(_dir).Length); // 임시 파일 정리됨
    }

    [Fact]
    public void Replace_MissingBackup_Throws()
        => Assert.Throws<CliError>(() =>
            WorkbookRollback.ReplaceWorkbook(Path.Combine(_dir, "nope.xlsm"), null, WriteFile("Book.xlsm", "x")));

    [Fact]
    public void IsLocked_DetectsExclusiveConflict()
    {
        string wb = WriteFile("Book.xlsm", "x");
        Assert.False(WorkbookRollback.IsLocked(wb));
        using (new FileStream(wb, FileMode.Open, FileAccess.Read, FileShare.Read))
            Assert.True(WorkbookRollback.IsLocked(wb));
        Assert.False(WorkbookRollback.IsLocked(Path.Combine(_dir, "absent.xlsm")));
    }

    // ---- sync 무효화 ----

    [Fact]
    public void InvalidateSync_ForcesFullPushAndBlocksPlainPull()
    {
        var sync = new SyncCache
        {
            LastPushFingerprint = "AAA",
            LastPushAt = "2026-10-17T03:00:00Z",
            LastPullFingerprint = "AAA",
        };
        sync.Forms["frmMain"] = new SyncCache.FormSync { FrmHash = "h", FrxHash = "x" };

        WorkbookRollback.InvalidateSync(sync, "20261017-031500-123-push");

        Assert.Null(sync.LastPushFingerprint); // push 캐시 히트 불가 → 전량 재주입
        Assert.Null(sync.LastPushAt);
        Assert.Empty(sync.Forms);              // 폼 레이아웃도 재Import 경로로
        // src가 이전 push/pull과 같아도 plain pull은 차단 — 복원으로 사라진 편집이 src에만 있을 수 있음
        Assert.NotNull(PullGuard.FindBlockReason("AAA", sync.LastPushFingerprint, sync.LastPullFingerprint));
        Assert.StartsWith(WorkbookRollback.PullBaselinePrefix, sync.LastPullFingerprint);
    }
}
