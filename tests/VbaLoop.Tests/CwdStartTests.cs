using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class CwdStartTests
{
    private sealed class TempDir : IDisposable
    {
        public string Path { get; } =
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        public TempDir() => Directory.CreateDirectory(Path);
        public void Dispose() { try { Directory.Delete(Path, recursive: true); } catch { } }
        public void File(string rel)
        {
            var p = System.IO.Path.Combine(Path, rel);
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(p)!);
            System.IO.File.WriteAllText(p, "x");
        }
    }

    [Fact]
    public void Empty_IsNewProjectCandidate()
    {
        using var t = new TempDir();
        var s = CwdStart.Inspect(t.Path);
        Assert.True(s.IsEmpty);
        Assert.Empty(s.Workbooks);
        Assert.False(s.IsProject);
    }

    [Fact]
    public void DotEntriesOnly_StillEmpty()
    {
        using var t = new TempDir();
        t.File(".gitignore");
        t.File(@".git\HEAD");
        Assert.True(CwdStart.Inspect(t.Path).IsEmpty);
    }

    [Fact]
    public void Workbooks_SortedWithoutLockStubs()
    {
        using var t = new TempDir();
        t.File("b.xlsm");
        t.File("A.xlsm");
        t.File("~$A.xlsm");
        t.File("data.xlsx");
        t.File(@"sub\deep.xlsm");

        var s = CwdStart.Inspect(t.Path);

        Assert.Equal(new[] { "A.xlsm", "b.xlsm" }, s.Workbooks);
        Assert.False(s.IsEmpty);
    }

    [Theory]
    [InlineData("vbaloop.toml")]
    [InlineData("vbaflow.toml")]
    public void Toml_IsProject(string toml)
    {
        using var t = new TempDir();
        t.File(toml);
        Assert.True(CwdStart.Inspect(t.Path).IsProject);
    }

    [Fact]
    public void MissingDir_TreatedAsEmpty()
        => Assert.True(CwdStart.Inspect(System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString("N"))).IsEmpty);
}
