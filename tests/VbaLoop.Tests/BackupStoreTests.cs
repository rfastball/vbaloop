using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class BackupStoreTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "vbaloop-backuptest", Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    private static readonly DateTime T0 = new(2026, 10, 17, 3, 15, 0, 123, DateTimeKind.Utc);

    /// <summary>완결 백업 한 건 생성(사본 + metadata).</summary>
    private BackupStore.Slot MakeBackup(DateTime utc, string kind = BackupStore.KindPush, string content = "wb")
    {
        var slot = BackupStore.Allocate(_root, kind, @"C:\proj\Book.xlsm", utc);
        File.WriteAllText(slot.CopyPath, content);
        BackupStore.Complete(slot, new BackupStore.Metadata { Command = "push", WorkbookPath = @"C:\proj\Book.xlsm" });
        return slot;
    }

    private static BackupStore.Entry E(string id, bool complete = true)
    {
        Assert.True(BackupStore.TryParseId(id, out var utc, out var kind));
        return new BackupStore.Entry(id, @"C:\x\" + id, utc, kind,
            complete ? new BackupStore.Metadata { Id = id, WorkbookFile = "Book.xlsm" } : null);
    }

    [Fact]
    public void MakeId_UtcStampThenKind()
    {
        Assert.Equal("20261017-031500-123-push", BackupStore.MakeId(T0, BackupStore.KindPush));
        Assert.Equal("20261017-031500-123-pre-rollback", BackupStore.MakeId(T0, BackupStore.KindPreRollback));
    }

    [Fact]
    public void TryParseId_RoundTrips()
    {
        Assert.True(BackupStore.TryParseId("20261017-031500-123-pre-rollback", out var utc, out var kind));
        Assert.Equal(T0, utc);
        Assert.Equal(DateTimeKind.Utc, utc.Kind);
        Assert.Equal("pre-rollback", kind);
    }

    [Theory]
    [InlineData("notes")]
    [InlineData("20261017-031500-123")]          // kind 없음
    [InlineData("20261017-031500-12-push")]      // ms 2자리
    [InlineData("20261317-031500-123-push")]     // 13월
    [InlineData("20261017-031500-123-PUSH")]     // 대문자 kind
    [InlineData("..\\20261017-031500-123-push")]
    public void TryParseId_RejectsForeignNames(string id)
        => Assert.False(BackupStore.TryParseId(id, out _, out _));

    [Fact]
    public void Allocate_CollisionShiftsByOneMillisecond()
    {
        var a = BackupStore.Allocate(_root, BackupStore.KindPush, @"C:\p\Book.xlsm", T0);
        var b = BackupStore.Allocate(_root, BackupStore.KindPush, @"C:\p\Book.xlsm", T0);
        Assert.Equal("20261017-031500-123-push", a.Id);
        Assert.Equal("20261017-031500-124-push", b.Id);
        Assert.Equal(Path.Combine(_root, a.Id, "Book.xlsm"), a.CopyPath);
    }

    [Fact]
    public void Complete_WritesMetadataWithHashAndSize()
    {
        var slot = MakeBackup(T0, content: "hello");
        var list = BackupStore.List(_root);
        var e = Assert.Single(list);
        Assert.True(e.Complete);
        Assert.Equal(slot.Id, e.Meta!.Id);
        Assert.Equal("push", e.Meta.Kind);
        Assert.Equal("Book.xlsm", e.Meta.WorkbookFile);
        Assert.Equal(5, e.Meta.SizeBytes);
        Assert.Equal(BackupStore.HashFile(slot.CopyPath), e.Meta.Sha256);
        Assert.Equal(64, e.Meta.Sha256.Length);
        Assert.Equal(slot.CopyPath, e.WorkbookCopyPath);
    }

    [Fact]
    public void Complete_WithoutCopy_Throws()
    {
        var slot = BackupStore.Allocate(_root, BackupStore.KindPush, @"C:\p\Book.xlsm", T0);
        Assert.Throws<CliError>(() => BackupStore.Complete(slot, new BackupStore.Metadata()));
    }

    [Fact]
    public void List_NewestFirst_IncompleteFlagged_ForeignIgnored()
    {
        MakeBackup(T0);
        MakeBackup(T0.AddMinutes(5), BackupStore.KindPreRollback);
        BackupStore.Allocate(_root, BackupStore.KindPush, @"C:\p\Book.xlsm", T0.AddMinutes(1)); // 미완(메타 없음)
        Directory.CreateDirectory(Path.Combine(_root, "my-notes"));                              // 사용자 폴더

        var list = BackupStore.List(_root);
        Assert.Equal(new[]
        {
            "20261017-032000-123-pre-rollback",
            "20261017-031600-123-push",
            "20261017-031500-123-push",
        }, list.Select(e => e.Id));
        Assert.Equal(new[] { true, false, true }, list.Select(e => e.Complete));
    }

    [Fact]
    public void List_MissingCopyOrTamperedMetadata_IsIncomplete()
    {
        var a = MakeBackup(T0);
        File.Delete(a.CopyPath);
        var b = MakeBackup(T0.AddMinutes(1));
        string meta = Path.Combine(b.Dir, BackupStore.MetadataFileName);
        File.WriteAllText(meta, File.ReadAllText(meta).Replace("\"Book.xlsm\"", "\"..\\\\Book.xlsm\""));
        var c = MakeBackup(T0.AddMinutes(2));
        File.WriteAllText(Path.Combine(c.Dir, BackupStore.MetadataFileName), "{ not json");

        Assert.All(BackupStore.List(_root), e => Assert.False(e.Complete));
    }

    [Fact]
    public void List_NoRoot_Empty() => Assert.Empty(BackupStore.List(Path.Combine(_root, "absent")));

    // ---- Select (rollback 대상 선택 규칙) ----

    [Fact]
    public void Select_Latest_PicksNewestComplete_AnyKind()
    {
        var entries = new[]
        {
            E("20261017-040000-000-push", complete: false),
            E("20261017-035000-000-pre-rollback"),
            E("20261017-030000-000-push"),
        };
        Assert.Equal("20261017-035000-000-pre-rollback", BackupStore.Select(entries, latest: true, id: null).Id);
    }

    [Fact]
    public void Select_ById_ExactMatch()
    {
        var entries = new[] { E("20261017-040000-000-push"), E("20261017-030000-000-push") };
        Assert.Equal("20261017-030000-000-push", BackupStore.Select(entries, false, "20261017-030000-000-push").Id);
    }

    [Fact]
    public void Select_ById_Incomplete_Refused()
    {
        var entries = new[] { E("20261017-040000-000-push", complete: false) };
        Assert.Throws<CliError>(() => BackupStore.Select(entries, false, "20261017-040000-000-push"));
    }

    [Fact]
    public void Select_ById_Unknown_HintListsRecent()
    {
        var entries = new[] { E("20261017-040000-000-push") };
        var ex = Assert.Throws<CliError>(() => BackupStore.Select(entries, false, "20261017-040000"));
        Assert.Contains("20261017-040000-000-push", ex.Hint);
    }

    [Theory]
    [InlineData(true, "20261017-040000-000-push")] // 둘 다
    [InlineData(false, null)]                      // 둘 다 없음
    public void Select_RequiresExactlyOneForm(bool latest, string? id)
        => Assert.Throws<CliError>(() => BackupStore.Select(new[] { E("20261017-040000-000-push") }, latest, id));

    [Fact]
    public void Select_Latest_NothingComplete_Throws()
        => Assert.Throws<CliError>(() => BackupStore.Select(new[] { E("20261017-040000-000-push", false) }, true, null));

    // ---- 보존 정책 ----

    private static List<BackupStore.Entry> Series(int n)
        => Enumerable.Range(0, n)
            .Select(i => E(BackupStore.MakeId(T0.AddMinutes(i), BackupStore.KindPush)))
            .OrderByDescending(e => e.Id, StringComparer.Ordinal).ToList();

    [Fact]
    public void PlanPrune_KeepsNewestN()
    {
        var entries = Series(12);
        var del = BackupStore.PlanPrune(entries, keep: 10);
        Assert.Equal(new[] { entries[10].Id, entries[11].Id }, del);
    }

    [Fact]
    public void PlanPrune_UnderLimit_DeletesNothing()
        => Assert.Empty(BackupStore.PlanPrune(Series(3), keep: 10));

    [Fact]
    public void PlanPrune_KeepZero_NeverDeletes()
        => Assert.Empty(BackupStore.PlanPrune(Series(30), keep: 0));

    [Fact]
    public void PlanPrune_IncompleteAlwaysDeleted_AndNotCounted()
    {
        var entries = new List<BackupStore.Entry>
        {
            E("20261017-050000-000-push", complete: false),
            E("20261017-040000-000-push"),
            E("20261017-030000-000-push"),
        };
        Assert.Equal(new[] { "20261017-050000-000-push" }, BackupStore.PlanPrune(entries, keep: 2));
    }

    [Fact]
    public void PlanPrune_CountProtected_KeepsExactlyN()
    {
        // push 경로: 방금 만든(가장 최신) 백업은 보호하되 keep 개수에 포함한다 — 런타임 실측에서
        // keep = 2인데 push 백업이 3개 남던 N+1 결함의 회귀 방지.
        var entries = Series(4);
        string newest = entries.OrderByDescending(e => e.Id, StringComparer.Ordinal).First().Id;
        var del = BackupStore.PlanPrune(entries, keep: 2, protect: new[] { newest }, countProtected: true);
        Assert.Equal(2, del.Count);
        Assert.DoesNotContain(newest, del);
    }

    [Fact]
    public void PlanPrune_ProtectedSurvivesAndDoesNotCount()
    {
        var entries = Series(4);
        string oldest = entries[^1].Id;
        var del = BackupStore.PlanPrune(entries, keep: 2, protect: new[] { oldest });
        Assert.DoesNotContain(oldest, del);
        Assert.Equal(new[] { entries[2].Id }, del);
    }

    [Fact]
    public void Prune_DeletesOnDisk_LeavesForeignDirs()
    {
        for (int i = 0; i < 4; i++) MakeBackup(T0.AddMinutes(i));
        Directory.CreateDirectory(Path.Combine(_root, "keep-me"));
        var (deleted, failed) = BackupStore.Prune(_root, keep: 2);
        Assert.Equal(2, deleted.Count);
        Assert.Empty(failed);
        Assert.Equal(2, BackupStore.List(_root).Count);
        Assert.True(Directory.Exists(Path.Combine(_root, "keep-me")));
    }

    [Fact]
    public void Discard_RemovesSlot()
    {
        var slot = BackupStore.Allocate(_root, BackupStore.KindPush, @"C:\p\Book.xlsm", T0);
        BackupStore.Discard(slot);
        Assert.False(Directory.Exists(slot.Dir));
    }
}
