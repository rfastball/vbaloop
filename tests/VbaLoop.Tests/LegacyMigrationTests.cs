using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>개명(vbaflow → vbaloop) 이전: 기존 온보딩 프로젝트가 깨지지 않고 init 한 번으로 수렴하는지.</summary>
public class LegacyMigrationTests
{
    private static readonly string[] Shim = { "Attribute VB_Name = \"LogShim\"", "Option Explicit" };
    private static readonly string[] DocsBody = { "## 개발 루프 (vbaloop)", "", "새 루프 설명" };
    private static readonly string[] Skill = { "---", "name: vbaloop", "---", "# 스킬" };
    private static readonly string[] Pitfalls = { "# pitfalls" };
    private static readonly string[] Dist = { "# dist" };

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

    private const string LegacyToml =
        "[project]\r\nname = \"P\"\r\n\r\n[excel]\r\nworkbook = \"P.xlsm\"\r\nvisible = false\r\n\r\n[src]\r\nroot = \"src\"\r\n";

    private static List<Onboarding.Item> Analyze(string root)
        => Onboarding.Analyze(root, ProjectConfig.FindTomlPath(null, root), Shim, DocsBody, Skill, Pitfalls, Dist);

    [Fact]
    public void FindTomlPath_FallsBackToLegacy_AndPrefersNewName()
    {
        using var t = new TempDir();
        var legacy = t.File(LegacyNames.TomlFile, LegacyToml);
        Assert.Equal(legacy, ProjectConfig.FindTomlPath(null, t.Path));
        Assert.True(ProjectConfig.IsLegacyToml(legacy));

        var modern = t.File("vbaloop.toml", LegacyToml);
        Assert.Equal(modern, ProjectConfig.FindTomlPath(null, t.Path));
    }

    [Fact]
    public void Init_MigratesLegacyProject_AndIsIdempotent()
    {
        using var t = new TempDir();
        t.File("P.xlsm");
        t.File(LegacyNames.TomlFile, LegacyToml);
        t.File("CLAUDE.md", "# P\r\n\r\n" + LegacyNames.ManagedStart + "\r\n## 개발 루프 (_VbaFlow)\r\n옛 절차\r\n"
            + LegacyNames.ManagedEnd + "\r\n\r\n## 내 섹션\r\n유지\r\n");
        t.File(".gitignore", "*.bak\r\n" + LegacyNames.GitignoreStart + "\r\n_dbg/\r\n" + LegacyNames.GitignoreEnd + "\r\n");
        var oldSkill = System.IO.Path.Combine(".claude", "skills", "vbaflow");
        t.File(System.IO.Path.Combine(oldSkill, "SKILL.md"), "old");
        t.File(System.IO.Path.Combine(oldSkill, "references", "vba-pitfalls.md"), "old");
        t.File(System.IO.Path.Combine(oldSkill, "references", "dist.md"), "old");
        var userNote = t.File(System.IO.Path.Combine(oldSkill, "my-notes.md"), "mine");

        var results = Onboarding.Apply(Analyze(t.Path));

        // toml: 내용 보존 이동
        Assert.False(File.Exists(System.IO.Path.Combine(t.Path, LegacyNames.TomlFile)));
        Assert.Equal(LegacyToml, File.ReadAllText(System.IO.Path.Combine(t.Path, "vbaloop.toml")));
        Assert.Contains(results, r => r.Item.Name == "vbaloop.toml" && r.Action == ArtifactAction.Migrated);

        // 관리 블록: 옛 마커/본문 교체, 사용자 섹션 보존
        var claude = File.ReadAllText(System.IO.Path.Combine(t.Path, "CLAUDE.md"));
        Assert.DoesNotContain("vbaflow:managed", claude);
        Assert.Contains(ManagedBlock.StartMarker, claude);
        Assert.Contains("새 루프 설명", claude);
        Assert.DoesNotContain("옛 절차", claude);
        Assert.Contains("## 내 섹션", claude);

        var gi = File.ReadAllText(System.IO.Path.Combine(t.Path, ".gitignore"));
        Assert.DoesNotContain("vbaflow:managed", gi);
        Assert.Contains(GitignoreBlock.StartMarker, gi);
        Assert.StartsWith("*.bak", gi);

        // 옛 스킬: 관리 파일만 삭제, 사용자 파일과 그 폴더는 보존
        Assert.False(File.Exists(System.IO.Path.Combine(t.Path, oldSkill, "SKILL.md")));
        Assert.False(Directory.Exists(System.IO.Path.Combine(t.Path, oldSkill, "references")));
        Assert.True(File.Exists(userNote));
        Assert.True(File.Exists(System.IO.Path.Combine(t.Path, ".claude", "skills", "vbaloop", "SKILL.md")));

        // 2회차: 전부 정상(남은 사용자 파일만 있는 옛 폴더는 이전 대상 아님)
        Assert.All(Analyze(t.Path), i => Assert.Equal(ArtifactStatus.Ok, i.Status));
    }

    [Fact]
    public void Init_RemovesLegacySkillFolder_WhenOnlyManagedFiles()
    {
        using var t = new TempDir();
        t.File("P.xlsm");
        t.File("vbaloop.toml", LegacyToml);
        var oldSkill = System.IO.Path.Combine(".claude", "skills", "vbaflow");
        t.File(System.IO.Path.Combine(oldSkill, "SKILL.md"), "old");

        Onboarding.Apply(Analyze(t.Path));

        Assert.False(Directory.Exists(System.IO.Path.Combine(t.Path, oldSkill)));
        Assert.True(Directory.Exists(System.IO.Path.Combine(t.Path, ".claude", "skills")));
    }

    [Fact]
    public void ManagedBlock_PreMarkerLegacyHeading_StillMigrates()
    {
        var (text, changed) = ManagedBlock.Upsert("# H\r\n\r\n" + LegacyNames.DocsHeading + "\r\n옛 절차\r\n\r\n## 다음\r\n유지\r\n",
            DocsBody);
        Assert.True(changed);
        Assert.Contains(ManagedBlock.StartMarker, text);
        Assert.DoesNotContain(LegacyNames.DocsHeading, text);
        Assert.Contains("## 다음", text);
    }

    [Fact]
    public void AgentIdentity_LegacySessionEnv_StillHonoured()
    {
        Assert.Contains(LegacyNames.SessionEnv, AgentIdentity.EnvIdNames);
        Assert.True(Array.IndexOf(AgentIdentity.EnvIdNames, "VBALOOP_SESSION_ID")
                    < Array.IndexOf(AgentIdentity.EnvIdNames, LegacyNames.SessionEnv));
    }

    [Fact]
    public void LoadedConfig_ReportsTheFileActuallyRead()
    {
        // 러너 자식/status가 경로를 조립하지 않고 이 값을 써야 이전 전 프로젝트에서도 동작한다
        using var t = new TempDir();
        var legacy = t.File(LegacyNames.TomlFile, LegacyToml);
        var cfg = ProjectConfig.Load(null, t.Path);
        Assert.Equal(legacy, cfg.TomlPath);
    }

    [Fact]
    public void InitAll_EnumeratesLegacyOnlyProjects()
    {
        using var t = new TempDir();
        t.File(System.IO.Path.Combine("vbaloop", "VbaLoop.sln"));
        t.File(System.IO.Path.Combine("Old", LegacyNames.TomlFile), LegacyToml);
        t.File(System.IO.Path.Combine("New", "vbaloop.toml"), LegacyToml);
        var names = RepoLocator.EnumerateOnboardedProjects(t.Path, System.IO.Path.Combine(t.Path, "vbaloop"))
            .Select(p => p.Name).ToArray();
        Assert.Equal(new[] { "New", "Old" }, names);
    }
}
