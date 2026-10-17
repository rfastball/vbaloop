using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class RepoLocatorTests
{
    private sealed class TempDir : IDisposable
    {
        public string Path { get; } =
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        public TempDir() => Directory.CreateDirectory(Path);
        public void Dispose() { try { Directory.Delete(Path, recursive: true); } catch { } }
        public string File(string name, string content = "")
        {
            var p = System.IO.Path.Combine(Path, name);
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(p)!);
            System.IO.File.WriteAllText(p, content);
            return p;
        }
    }

    [Fact]
    public void FindRepoRoot_FromNestedPublishDir_WalksUpToSln()
    {
        using var t = new TempDir();
        t.File(System.IO.Path.Combine("vbaloop", "VbaLoop.sln"));
        var publishDir = System.IO.Path.Combine(
            t.Path, "vbaloop", "src", "VbaLoop", "bin", "Release", "net10.0-windows", "win-x64", "publish");
        Directory.CreateDirectory(publishDir);
        Assert.Equal(System.IO.Path.Combine(t.Path, "vbaloop"), RepoLocator.FindRepoRoot(publishDir));
    }

    [Fact]
    public void FindRepoRoot_MarkerAbsent_ReturnsNull()
    {
        using var t = new TempDir();
        var deep = System.IO.Path.Combine(t.Path, "a", "b");
        Directory.CreateDirectory(deep);
        // 임시 폴더 상위 어딘가에 우연히 sln이 있어도 t.Path 밖이므로 테스트가 흔들릴 수 있으나,
        // %TEMP% 경로에 VbaLoop.sln이 있을 가능성은 없다고 본다
        Assert.Null(RepoLocator.FindRepoRoot(deep));
    }

    [Fact]
    public void EnumerateOnboardedProjects_TomlSiblings_InclKoreanNames_ExclOthers()
    {
        using var t = new TempDir();
        t.File(System.IO.Path.Combine("vbaloop", "VbaLoop.sln"));
        t.File(System.IO.Path.Combine("SampleRpa", "vbaloop.toml"));
        t.File(System.IO.Path.Combine("한글 프로젝트", "vbaloop.toml"));
        Directory.CreateDirectory(System.IO.Path.Combine(t.Path, "docs")); // toml 없음 → 제외
        t.File(System.IO.Path.Combine("_VbaHarness", "readme.txt"));       // toml 없음 → 제외

        var repo = System.IO.Path.Combine(t.Path, "vbaloop");
        var projects = RepoLocator.EnumerateOnboardedProjects(t.Path, repo);
        Assert.Equal(new[] { "SampleRpa", "한글 프로젝트" }, projects.Select(p => p.Name).ToArray());
        Assert.All(projects, p => Assert.True(Directory.Exists(p.Root)));
    }

    [Fact]
    public void EnumerateOnboardedProjects_RepoItselfExcluded_EvenWithToml()
    {
        using var t = new TempDir();
        t.File(System.IO.Path.Combine("vbaloop", "VbaLoop.sln"));
        t.File(System.IO.Path.Combine("vbaloop", "vbaloop.toml")); // 레포 자신은 무조건 제외
        var repo = System.IO.Path.Combine(t.Path, "vbaloop");
        Assert.Empty(RepoLocator.EnumerateOnboardedProjects(t.Path, repo));
    }

    [Fact]
    public void EnumerateOnboardedProjects_PilotForm_IncludedIffHasToml()
    {
        using var t = new TempDir();
        t.File(System.IO.Path.Combine("vbaloop", "VbaLoop.sln"));
        var repo = System.IO.Path.Combine(t.Path, "vbaloop");

        Directory.CreateDirectory(System.IO.Path.Combine(repo, "_pilot_form"));
        Assert.Empty(RepoLocator.EnumerateOnboardedProjects(t.Path, repo)); // toml 없음 → 제외

        t.File(System.IO.Path.Combine("vbaloop", "_pilot_form", "vbaloop.toml"));
        var projects = RepoLocator.EnumerateOnboardedProjects(t.Path, repo);
        Assert.Equal(new[] { "_pilot_form" }, projects.Select(p => p.Name).ToArray());
    }
}
