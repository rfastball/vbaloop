using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class WorkbookIntakeTests
{
    private sealed class TempDir : IDisposable
    {
        public string Path { get; } =
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        public TempDir() => Directory.CreateDirectory(Path);
        public void Dispose() { try { Directory.Delete(Path, recursive: true); } catch { } }
        public string File(string rel, string content = "wb")
        {
            var p = System.IO.Path.Combine(Path, rel);
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(p)!);
            System.IO.File.WriteAllText(p, content);
            return p;
        }
        public string Dir(string rel) => Directory.CreateDirectory(System.IO.Path.Combine(Path, rel)).FullName;
    }

    private static string Toml(string workbook) =>
        $"[project]\nname = \"P\"\n\n[excel]\nworkbook = \"{workbook}\"\n";

    [Theory]
    [InlineData(@"C:\Users\me\Desktop\Book.xlsm", true)]
    [InlineData("Book.xlsm", true)]
    [InlineData("setup", false)]
    [InlineData("--version", false)]
    [InlineData("push", false)]
    public void LooksLikePath(string arg, bool expected)
        => Assert.Equal(expected, WorkbookIntake.LooksLikePath(arg));

    [Fact]
    public void Validate_StripsQuotes_ReturnsFullPath()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");
        Assert.Equal(wb, WorkbookIntake.Validate($"  \"{wb}\" "));
    }

    [Theory]
    [InlineData("Book.xlsx")]
    [InlineData("Book.xls")]
    [InlineData("Book.xlam")]
    [InlineData("Book.txt")]
    [InlineData("~$Book.xlsm")]
    public void Validate_RejectsNonDevWorkbook(string name)
    {
        using var t = new TempDir();
        var p = t.File(name);
        Assert.Throws<CliError>(() => WorkbookIntake.Validate(p));
    }

    [Fact]
    public void Validate_Missing_Throws()
    {
        using var t = new TempDir();
        Assert.Throws<CliError>(() => WorkbookIntake.Validate(System.IO.Path.Combine(t.Path, "None.xlsm")));
    }

    [Fact]
    public void Decide_Fresh_CopiesIntoSiblingFolderNamedAfterWorkbook()
    {
        using var t = new TempDir();
        var wb = t.File("월별 정산.xlsm");
        t.File("other.xlsm"); // 원본 폴더에 다른 워크북이 있어도 사본 폴더는 깨끗하다

        var plan = WorkbookIntake.Decide(wb, null);

        Assert.Equal(WorkbookIntake.Kind.Copy, plan.Kind);
        Assert.Equal(System.IO.Path.Combine(t.Path, "월별 정산"), plan.Folder);
        Assert.Equal("월별 정산.xlsm", plan.WorkbookName);
    }

    [Fact]
    public void Execute_Copy_LeavesOriginal_ClearsReadOnly()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm", "payload");
        File.SetAttributes(wb, FileAttributes.ReadOnly);
        try
        {
            var plan = WorkbookIntake.Decide(wb, null);
            WorkbookIntake.Execute(plan);

            var copy = System.IO.Path.Combine(plan.Folder, "Book.xlsm");
            Assert.Equal("payload", File.ReadAllText(copy));
            Assert.Equal(0, (int)(File.GetAttributes(copy) & FileAttributes.ReadOnly));
            Assert.True(File.Exists(wb));
            File.SetAttributes(copy, FileAttributes.Normal);
        }
        finally
        {
            File.SetAttributes(wb, FileAttributes.Normal);
        }
    }

    [Fact]
    public void Decide_SecondRun_ReusesOnboardedFolder()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm", "original");
        t.File(@"Book\Book.xlsm", "edited by push"); // 사본은 이후 원본과 달라진다
        t.File(@"Book\vbaloop.toml", Toml("Book.xlsm"));

        var plan = WorkbookIntake.Decide(wb, null);

        Assert.Equal(WorkbookIntake.Kind.Reuse, plan.Kind);
        Assert.Equal(System.IO.Path.Combine(t.Path, "Book"), plan.Folder);
    }

    [Fact]
    public void Decide_RetryAfterCopyOnlyFailure_ReusesIdenticalCopy()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm", "same");
        t.File(@"Book\Book.xlsm", "same");
        Assert.Equal(WorkbookIntake.Kind.Reuse, WorkbookIntake.Decide(wb, null).Kind);
    }

    [Theory]
    [InlineData(@"Book\notes.txt")]          // 관계없는 파일이 든 남의 폴더
    [InlineData(@"Book\Book.xlsm")]          // 같은 이름이지만 내용이 다른 워크북(toml 없음)
    [InlineData(@"Book\Other.xlsm")]         // 다른 워크북
    public void Decide_OccupiedFolder_FallsBackToNumberedSibling(string occupant)
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm", "original");
        t.File(occupant, "someone else's");

        var plan = WorkbookIntake.Decide(wb, null);

        Assert.Equal(WorkbookIntake.Kind.Copy, plan.Kind);
        Assert.Equal(System.IO.Path.Combine(t.Path, "Book (2)"), plan.Folder);
    }

    [Fact]
    public void Decide_SourceInsideProject_InPlace()
    {
        using var t = new TempDir();
        var wb = t.File(@"Proj\Book.xlsm");
        t.File(@"Proj\vbaloop.toml", Toml("Book.xlsm"));

        var plan = WorkbookIntake.Decide(wb, null);

        Assert.Equal(WorkbookIntake.Kind.InPlace, plan.Kind);
        Assert.Equal(System.IO.Path.Combine(t.Path, "Proj"), plan.Folder);
    }

    [Fact]
    public void Decide_SourceInsideProjectOfAnotherWorkbook_Throws()
    {
        using var t = new TempDir();
        var wb = t.File(@"Proj\Scratch.xlsm");
        t.File(@"Proj\Main.xlsm");
        t.File(@"Proj\vbaloop.toml", Toml("Main.xlsm"));
        Assert.Throws<CliError>(() => WorkbookIntake.Decide(wb, null));
    }

    [Fact]
    public void Decide_FolderOverride_EmptyOrUnrelated_Copies()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");
        var target = t.Dir("repo");
        t.File(@"repo\README.md");

        var plan = WorkbookIntake.Decide(wb, target);

        Assert.Equal(WorkbookIntake.Kind.Copy, plan.Kind);
        Assert.Equal(target, plan.Folder);
    }

    [Theory]
    [InlineData(@"dest\Other.xlsm")]
    [InlineData(@"dest\Book.xlsm")]
    public void Decide_FolderOverride_Conflict_Throws(string occupant)
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm", "original");
        t.File(occupant, "different");
        Assert.Throws<CliError>(() => WorkbookIntake.Decide(wb, System.IO.Path.Combine(t.Path, "dest")));
    }

    [Fact]
    public void Decide_FolderOverrideIsSourceDir_WithOtherWorkbooks_Throws()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");
        t.File("Other.xlsm");
        Assert.Throws<CliError>(() => WorkbookIntake.Decide(wb, t.Path));
    }

    [Fact]
    public void Decide_FolderOverrideIsSourceDir_Alone_InPlace()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");
        Assert.Equal(WorkbookIntake.Kind.InPlace, WorkbookIntake.Decide(wb, t.Path + @"\").Kind);
    }

    [Fact]
    public void Decide_FileAtCopyFolderPath_FallsBackToNumberedSibling()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");
        t.File("Book", "확장자 없는 파일"); // 폴더를 만들 수 없는 자리

        var plan = WorkbookIntake.Decide(wb, null);

        Assert.Equal(WorkbookIntake.Kind.Copy, plan.Kind);
        Assert.Equal(System.IO.Path.Combine(t.Path, "Book (2)"), plan.Folder);
    }

    [Fact]
    public void Decide_FolderOverrideIsAFile_Throws()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");
        var file = t.File("dest.txt");
        Assert.Throws<CliError>(() => WorkbookIntake.Decide(wb, file));
    }
}
