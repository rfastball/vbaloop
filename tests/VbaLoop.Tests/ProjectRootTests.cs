using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ProjectRootTests
{
    private sealed class TempDir : IDisposable
    {
        public string Path { get; } =
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        public TempDir() => Directory.CreateDirectory(Path);
        public void Dispose() { try { Directory.Delete(Path, recursive: true); } catch { } }
        public string File(string rel, string content = "x")
        {
            var p = System.IO.Path.Combine(Path, rel);
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(p)!);
            System.IO.File.WriteAllText(p, content);
            return p;
        }
    }

    private static string Toml(string workbook) => Onboarding.BuildTomlScaffold("P", workbook);

    [Fact]
    public void Enclosing_FindsAncestorProject_ForFolderWithoutOwnToml()
    {
        using var t = new TempDir();
        t.File("vbaloop.toml", Toml("Outer.xlsm"));
        string inner = System.IO.Path.Combine(t.Path, "a", "b");

        Assert.True(PathDrift.SamePath(t.Path, ProjectRoot.Enclosing(inner)!));
        Assert.Throws<CliError>(() => ProjectRoot.EnsureNotNested(inner));
    }

    [Fact]
    public void Enclosing_NullForOwnRootOrNoProject()
    {
        using var t = new TempDir();
        Assert.Null(ProjectRoot.Enclosing(t.Path));                 // 상위 어디에도 없음
        t.File(@"proj\vbaloop.toml", Toml("Main.xlsm"));
        Assert.Null(ProjectRoot.Enclosing(System.IO.Path.Combine(t.Path, "proj"))); // 자기 toml = 자기 루트
        ProjectRoot.EnsureNotNested(System.IO.Path.Combine(t.Path, "proj"));
    }

    [Fact]
    public void ConfiguredWorkbook_RelativePathFromToml()
    {
        using var t = new TempDir();
        t.File(@"wb\Main.xlsm");
        t.File("vbaloop.toml", Toml("wb/Main.xlsm"));
        Assert.Equal(System.IO.Path.Combine("wb", "Main.xlsm"), ProjectRoot.ConfiguredWorkbook(t.Path));
    }

    [Fact]
    public void ConfiguredWorkbook_NullWithoutOwnToml()
    {
        using var t = new TempDir();
        t.File("vbaloop.toml", Toml("Outer.xlsm"));
        Assert.Null(ProjectRoot.ConfiguredWorkbook(System.IO.Path.Combine(t.Path, "child")));
    }

    [Fact]
    public void ConfiguredWorkbook_MissingFile_Throws()
    {
        using var t = new TempDir();
        t.File("vbaloop.toml", Toml("Gone.xlsm"));
        Assert.Throws<CliError>(() => ProjectRoot.ConfiguredWorkbook(t.Path));
    }
}
