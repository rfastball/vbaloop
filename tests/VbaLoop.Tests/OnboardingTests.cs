using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class OnboardingTests
{
    private static readonly string[] Shim =
        { "Attribute VB_Name = \"LogShim\"", "' 정본 v1", "Option Explicit", "Sub A()", "End Sub" };
    private static readonly string[] DocsBody = { "## 개발 루프 (vbaloop)", "", "루프 설명" };
    private static readonly string[] Skill = { "---", "name: vbaloop", "---", "# 스킬 v1" };
    private static readonly string[] Pitfalls = { "# VBA pitfall 참조 v1", "", "- 항목 1" };
    private static readonly string[] Dist = { "# vbaloop dist v1", "", "배포 절차" };
    private static readonly string SkillRelPath =
        System.IO.Path.Combine(".claude", "skills", "vbaloop", "SKILL.md");
    private static readonly string PitfallsRelPath =
        System.IO.Path.Combine(".claude", "skills", "vbaloop", "references", "vba-pitfalls.md");
    private static readonly string DistRelPath =
        System.IO.Path.Combine(".claude", "skills", "vbaloop", "references", "dist.md");

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
    public void InferWorkbook_Single_Returns()
    {
        using var t = new TempDir();
        t.File("업무 정산 목록.xlsm");
        Assert.Equal("업무 정산 목록.xlsm", Onboarding.InferWorkbook(t.Path));
    }

    [Fact]
    public void InferWorkbook_ExcludesLockFiles()
    {
        using var t = new TempDir();
        t.File("SampleRpa.xlsm");
        t.File("~$SampleRpa.xlsm");
        Assert.Equal("SampleRpa.xlsm", Onboarding.InferWorkbook(t.Path));
    }

    [Fact]
    public void InferWorkbook_None_Throws()
    {
        using var t = new TempDir();
        var e = Assert.Throws<CliError>(() => Onboarding.InferWorkbook(t.Path));
        Assert.Contains("찾지 못함", e.Message);
    }

    [Fact]
    public void InferWorkbook_Multiple_Throws()
    {
        using var t = new TempDir();
        t.File("SampleFilter(v2.6).xlsm");
        t.File("백업 사본.xlsm");
        var e = Assert.Throws<CliError>(() => Onboarding.InferWorkbook(t.Path));
        Assert.Contains("SampleFilter(v2.6).xlsm", e.Message);
        Assert.Contains("백업 사본.xlsm", e.Message);
    }

    [Fact]
    public void TomlScaffold_RoundTripsThroughProjectConfigLoad()
    {
        using var t = new TempDir();
        var toml = Onboarding.BuildTomlScaffold("My Proj", "My Book(v2.6).xlsm");
        var tomlPath = t.File("vbaloop.toml", toml);
        var cfg = ProjectConfig.Load(tomlPath, t.Path);
        Assert.Equal("My Proj", cfg.ProjectName);
        Assert.Equal(System.IO.Path.Combine(t.Path, "My Book(v2.6).xlsm"), cfg.WorkbookPath);
        Assert.True(cfg.Visible);
        Assert.Equal(SessionMode.Ephemeral, cfg.SessionMode);
        Assert.Equal(System.IO.Path.Combine(t.Path, "src"), cfg.SrcRoot);
    }

    [Fact]
    public void ShimCompare_BomLfTrailingWs_NotDrift()
    {
        using var t = new TempDir();
        // BOM + LF + 후행 공백 + 끝 빈 줄 - 전부 정규화 대상, 드리프트 아님
        var text = "﻿" + string.Join("\n", Shim.Select(l => l + "  ")) + "\n\n\n";
        var p = System.IO.Path.Combine(t.Path, "LogShim.bas");
        File.WriteAllText(p, text, new UTF8Encoding(true));
        Assert.False(Onboarding.IsWholeFileDrifted(p, Shim));
    }

    [Fact]
    public void ShimCompare_LegacyAddinVersion_Drift()
    {
        using var t = new TempDir();
        var p = t.File("LogShim.bas",
            "Attribute VB_Name = \"LogShim\"\r\nPrivate Const ADDIN_FILENAME As String = \"_DebugAddin.xlam\"\r\n");
        Assert.True(Onboarding.IsWholeFileDrifted(p, Shim));
    }

    [Fact]
    public void Analyze_FreshDir_AllMissing()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var items = Onboarding.Analyze(t.Path, tomlPath: null, Shim, DocsBody, Skill, Pitfalls, Dist);
        Assert.Equal(new[] { "vbaloop.toml", System.IO.Path.Combine("src", "LogShim.bas"), "CLAUDE.md", "AGENTS.md", SkillRelPath, PitfallsRelPath, DistRelPath, ".gitignore" },
            items.Select(i => i.Name).ToArray());
        Assert.All(items, i => Assert.Equal(ArtifactStatus.Missing, i.Status));
        Assert.Contains("workbook = \"Book.xlsm\"", items[0].NewContent);
        Assert.Contains($"name = \"{new DirectoryInfo(t.Path).Name}\"", items[0].NewContent);
    }

    [Fact]
    public void Apply_Then_Analyze_AllOk_And_SecondApply_AllUnchanged()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var first = Onboarding.Apply(Onboarding.Analyze(t.Path, null, Shim, DocsBody, Skill, Pitfalls, Dist));
        Assert.All(first, r => Assert.Equal(ArtifactAction.Created, r.Action));
        Assert.True(File.Exists(System.IO.Path.Combine(t.Path, SkillRelPath))); // 중첩 디렉토리 생성
        Assert.True(File.Exists(System.IO.Path.Combine(t.Path, PitfallsRelPath))); // references/ 하위

        var tomlPath = ProjectConfig.FindTomlPath(null, t.Path);
        Assert.NotNull(tomlPath);
        var again = Onboarding.Analyze(t.Path, tomlPath, Shim, DocsBody, Skill, Pitfalls, Dist);
        Assert.All(again, i => Assert.Equal(ArtifactStatus.Ok, i.Status));

        var second = Onboarding.Apply(again);
        Assert.All(second, r => Assert.Equal(ArtifactAction.Unchanged, r.Action));
    }

    [Fact]
    public void Analyze_DriftedShimAndUnmarkedDoc_ReportsDrift()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var tomlPath = t.File("vbaloop.toml", Onboarding.BuildTomlScaffold("P", "Book.xlsm"));
        t.File(System.IO.Path.Combine("src", "LogShim.bas"), "옛 버전\r\n");
        t.File("CLAUDE.md", "# P\n\n## 개발 루프 (vbaloop)\n옛 절차\n");
        var items = Onboarding.Analyze(t.Path, tomlPath, Shim, DocsBody, Skill, Pitfalls, Dist);
        Assert.Equal(ArtifactStatus.Ok, items[0].Status);       // toml
        Assert.Equal(ArtifactStatus.Drift, items[1].Status);    // shim
        Assert.Equal(ArtifactStatus.Drift, items[2].Status);    // CLAUDE.md (마커 없는 블록 헤딩)
        Assert.Equal(ArtifactStatus.Missing, items[3].Status);  // AGENTS.md
        Assert.Equal(ArtifactStatus.Missing, items[4].Status);  // 스킬
        Assert.Equal(ArtifactStatus.Missing, items[5].Status);  // pitfall 참조
    }

    [Fact]
    public void Analyze_DriftedSkill_ReportsDriftAndOverwrites()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var tomlPath = t.File("vbaloop.toml", Onboarding.BuildTomlScaffold("P", "Book.xlsm"));
        t.File(SkillRelPath, "# 손댄 스킬\r\n");
        var items = Onboarding.Analyze(t.Path, tomlPath, Shim, DocsBody, Skill, Pitfalls, Dist);
        var skillItem = items.Single(i => i.Name == SkillRelPath);
        Assert.Equal(ArtifactStatus.Drift, skillItem.Status);

        var results = Onboarding.Apply(items);
        Assert.Equal(ArtifactAction.Overwritten,
            results.Single(r => r.Item.Name == SkillRelPath).Action);
        Assert.False(Onboarding.IsWholeFileDrifted(
            System.IO.Path.Combine(t.Path, SkillRelPath), Skill)); // 정본 수렴
    }

    [Fact]
    public void Analyze_DriftedPitfalls_ReportsDriftAndOverwrites()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var tomlPath = t.File("vbaloop.toml", Onboarding.BuildTomlScaffold("P", "Book.xlsm"));
        t.File(PitfallsRelPath, "# 프로젝트에서 직접 고친 사본\r\n");
        var items = Onboarding.Analyze(t.Path, tomlPath, Shim, DocsBody, Skill, Pitfalls, Dist);
        var item = items.Single(i => i.Name == PitfallsRelPath);
        Assert.Equal(ArtifactStatus.Drift, item.Status);

        var results = Onboarding.Apply(items);
        Assert.Equal(ArtifactAction.Overwritten,
            results.Single(r => r.Item.Name == PitfallsRelPath).Action);
        Assert.False(Onboarding.IsWholeFileDrifted(
            System.IO.Path.Combine(t.Path, PitfallsRelPath), Pitfalls)); // 정본 수렴
    }

    [Fact]
    public void Analyze_ExistingGitignore_UpsertsBlockPreservingUserLines()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var tomlPath = t.File("vbaloop.toml", Onboarding.BuildTomlScaffold("P", "Book.xlsm"));
        t.File(".gitignore", "*.bak\r\nsecret.txt\r\n");
        var items = Onboarding.Analyze(t.Path, tomlPath, Shim, DocsBody, Skill, Pitfalls, Dist);
        var gi = items.Single(i => i.Name == ".gitignore");
        Assert.Equal(ArtifactStatus.Drift, gi.Status); // 파일은 있으나 관리 블록 부재

        Onboarding.Apply(items);
        var text = System.IO.File.ReadAllText(System.IO.Path.Combine(t.Path, ".gitignore"));
        Assert.Contains("*.bak", text);          // 사용자 라인 보존
        Assert.Contains("secret.txt", text);
        Assert.Contains(GitignoreBlock.StartMarker, text);
        Assert.Contains("dist/", text);
        Assert.Contains("_dbg/", text);

        // 멱등: 재분석 시 Ok
        var again = Onboarding.Analyze(t.Path, tomlPath, Shim, DocsBody, Skill, Pitfalls, Dist);
        Assert.Equal(ArtifactStatus.Ok, again.Single(i => i.Name == ".gitignore").Status);
    }

    [Fact]
    public void EnumerateThenAnalyze_BrokenToml_DoesNotBlockOthers()
    {
        // init --all의 continue-and-summarize 계약을 공개 조각(열거 + Analyze)으로 검증:
        // 깨진 toml 프로젝트는 CliError, 정상 프로젝트는 그대로 수렴 가능해야 한다
        using var t = new TempDir();
        t.File(System.IO.Path.Combine("vbaloop", "VbaLoop.sln"));
        t.File(System.IO.Path.Combine("GoodProj", "vbaloop.toml"),
            Onboarding.BuildTomlScaffold("GoodProj", "Book.xlsm"));
        t.File(System.IO.Path.Combine("BadProj", "vbaloop.toml"), "[excel]\nworkbook = \n");

        var repo = System.IO.Path.Combine(t.Path, "vbaloop");
        var projects = RepoLocator.EnumerateOnboardedProjects(t.Path, repo);
        Assert.Equal(new[] { "BadProj", "GoodProj" }, projects.Select(p => p.Name).ToArray());

        var outcomes = new List<(string Name, bool Ok)>();
        foreach (var (name, root) in projects)
        {
            try
            {
                var items = Onboarding.Analyze(root, System.IO.Path.Combine(root, "vbaloop.toml"),
                    Shim, DocsBody, Skill, Pitfalls, Dist);
                Onboarding.Apply(items);
                outcomes.Add((name, true));
            }
            catch (CliError) { outcomes.Add((name, false)); }
        }
        Assert.Equal(new[] { ("BadProj", false), ("GoodProj", true) }, outcomes.ToArray());
        Assert.True(File.Exists(System.IO.Path.Combine(t.Path, "GoodProj", "src", "LogShim.bas")));
    }

    [Fact]
    public void Analyze_InvalidToml_Throws()
    {
        using var t = new TempDir();
        var tomlPath = t.File("vbaloop.toml", "[excel]\nworkbook = \n");
        Assert.Throws<CliError>(() => Onboarding.Analyze(t.Path, tomlPath, Shim, DocsBody, Skill, Pitfalls, Dist));
    }
}
