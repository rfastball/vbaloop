using System.Text.Json;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ProjectRegistryTests
{
    private static readonly DateTimeOffset Now = new(2026, 10, 17, 12, 0, 0, TimeSpan.Zero);

    private static ProjectRegistry.Entry E(string root) => new() { Root = root, AddedAt = "t" };

    [Fact]
    public void Upsert_AddsMissing_IgnoresSameRootCaseAndTrailingSlash()
    {
        var (one, added) = ProjectRegistry.Upsert(Array.Empty<ProjectRegistry.Entry>(), @"C:\Work\UnivDB\", Now);
        Assert.True(added);
        Assert.Equal(@"C:\Work\UnivDB", Assert.Single(one).Root); // 정규화(끝 구분자 제거)
        Assert.Equal(Now.ToString("o"), one[0].AddedAt);

        var (same, again) = ProjectRegistry.Upsert(one, @"c:\work\univdb", Now);
        Assert.False(again);
        Assert.Single(same);

        var (two, added2) = ProjectRegistry.Upsert(same, @"C:\Work\UnivCore", Now);
        Assert.True(added2);
        Assert.Equal(2, two.Count);
    }

    [Fact]
    public void Prune_RemovesGoneAndBlank_DedupesRoots()
    {
        var entries = new[] { E(@"C:\A"), E(@"C:\Gone"), E(""), E(@"c:\a\"), E(@"C:\B") };
        var (kept, removed) = ProjectRegistry.Prune(entries, r => !r.Contains("Gone"));
        Assert.Equal(new[] { @"C:\A", @"C:\B" }, kept.Select(e => e.Root).ToArray());
        Assert.Equal(new[] { @"C:\Gone", "" }, removed.Select(e => e.Root).ToArray());
    }

    [Fact]
    public void Union_SiblingsFirst_ThenRegistryOnlyByName()
    {
        var siblings = new List<(string Name, string Root)> { ("UnivDB", @"C:\P\UnivDB"), ("_pilot_form", @"C:\P\vbaloop\_pilot_form") };
        var registry = new[] { @"C:\Other\Zeta", @"c:\p\univdb", @"C:\Other\Alpha\", @"C:\Other\alpha" };
        var all = ProjectRegistry.Union(siblings, registry);
        Assert.Equal(new[] { "UnivDB", "_pilot_form", "Alpha", "Zeta" }, all.Select(p => p.Name).ToArray());
        Assert.Equal(@"C:\Other\Alpha", all[2].Root);
    }

    [Fact]
    public void Union_NoRepo_IsRegistryOnly()
    {
        var all = ProjectRegistry.Union(new List<(string, string)>(), new[] { @"C:\X\B", @"C:\X\A" });
        Assert.Equal(new[] { "A", "B" }, all.Select(p => p.Name).ToArray());
    }

    [Fact]
    public void File_TryAddThenPrune_RoundTrips()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        try
        {
            var reg = Path.Combine(dir, "projects.json");
            var live = Path.Combine(dir, "Live");
            var dead = Path.Combine(dir, "Dead");
            Directory.CreateDirectory(live);
            File.WriteAllText(Path.Combine(live, "vbaloop.toml"), "");
            Directory.CreateDirectory(dead);

            ProjectRegistry.TryAdd(reg, live);
            ProjectRegistry.TryAdd(reg, live + "\\");  // 중복
            ProjectRegistry.TryAdd(reg, dead);
            Assert.Equal(2, ProjectRegistry.Load(reg).Count);
            using (var doc = JsonDocument.Parse(File.ReadAllText(reg)))
                Assert.Equal(2, doc.RootElement.GetProperty("projects").GetArrayLength());

            var (kept, removed) = ProjectRegistry.PruneFile(reg);
            Assert.Equal(new[] { ProjectRegistry.Normalize(live) }, kept.ToArray());
            Assert.Equal(new[] { ProjectRegistry.Normalize(dead) }, removed.ToArray());
            Assert.Single(ProjectRegistry.Load(reg)); // 파일에도 반영

            File.WriteAllText(reg, "garbage");
            Assert.Empty(ProjectRegistry.Load(reg));  // 깨진 파일 = 빈 목록
            ProjectRegistry.TryAdd(reg, live);         // 그리고 다시 기록 가능
            Assert.Single(ProjectRegistry.Load(reg));
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }
}
