using System.Text.Json;
using VbaLoop.Cli;
using VbaLoop.Commands;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ManagedStampTests
{
    private static SemVer V(string s) => SemVer.ParseOrNull(s)!;

    // ---- 판독 ----

    [Fact]
    public void Read_FindsKeyOnlyInVbaloopSection()
    {
        var toml = "[project]\nmanaged_by = \"9.9.9\"\n\n[vbaloop]  # 도구 소유\nmanaged_by = \"0.3.0\"  # 주석\n";
        Assert.Equal("0.3.0", ManagedStamp.Read(toml));
    }

    [Theory]
    [InlineData("[project]\nname = \"P\"\n")]
    [InlineData("[vbaloop]\n")]
    [InlineData("")]
    public void Read_Absent_ReturnsNull(string toml) => Assert.Null(ManagedStamp.Read(toml));

    [Fact]
    public void Read_ToleratesSyntaxThisExeDoesNotKnow()
    {
        // 미래 exe가 쓴 toml(이 exe의 TomlMini가 거부할 문법)에서도 스탬프는 읽혀야 한다
        var toml = "[excel]\nworkbook = \"B.xlsm\"\nratio = 1.5\n[vbaloop]\nmanaged_by = \"2.0.0\"\nchannel = { a = 1 }\n";
        Assert.Equal("2.0.0", ManagedStamp.Read(toml));
    }

    // ---- 판정 매트릭스 ----

    [Theory]
    [InlineData("0.3.0", null, StampState.Unstamped)]
    [InlineData("0.3.0", "0.3.0", StampState.Current)]
    [InlineData("0.4.0", "0.3.0", StampState.ExeNewer)]
    [InlineData("0.3.0", "0.4.0", StampState.ExeOlder)]
    [InlineData("0.4.0-rc.1", "0.4.0", StampState.ExeOlder)]   // prerelease는 같은 번호 정식보다 낮다
    [InlineData("0.4.0", "0.4.0-rc.1", StampState.ExeNewer)]
    [InlineData("0.4.0-rc.2", "0.4.0-rc.1", StampState.ExeNewer)]
    [InlineData("0.3.0", "v0.3.0", StampState.Current)]
    [InlineData("0.3.0", "garbage", StampState.Invalid)]
    [InlineData("0.3.0", "", StampState.Invalid)]
    [InlineData(null, "0.3.0", StampState.ExeOlder)]          // exe 버전 판독 불가 -> 쓰지 않음
    [InlineData(null, null, StampState.Unstamped)]
    public void Decide_Matrix(string? exe, string? stamp, StampState expected)
        => Assert.Equal(expected, ManagedStamp.Decide(exe == null ? null : V(exe), stamp));

    [Theory]
    [InlineData(StampState.Unstamped, true)]
    [InlineData(StampState.Current, true)]
    [InlineData(StampState.ExeNewer, true)]
    [InlineData(StampState.ExeOlder, false)]
    [InlineData(StampState.Invalid, false)]
    public void AllowsWrite_OnlyWhenExeNotOlder(StampState s, bool allowed)
        => Assert.Equal(allowed, ManagedStamp.AllowsWrite(s));

    [Fact]
    public void StampToWrite_NeverForRefusedStatesOrUnknownExe()
    {
        Assert.Equal("0.4.0", ManagedStamp.StampToWrite(V("0.4.0"), StampState.ExeNewer));
        Assert.Equal("0.4.0", ManagedStamp.StampToWrite(V("v0.4.0"), StampState.Unstamped));
        Assert.Null(ManagedStamp.StampToWrite(V("0.4.0"), StampState.ExeOlder));
        Assert.Null(ManagedStamp.StampToWrite(V("0.4.0"), StampState.Invalid));
        Assert.Null(ManagedStamp.StampToWrite(null, StampState.Unstamped));
    }

    // ---- upsert ----

    [Fact]
    public void Upsert_NoSection_AppendsAtEndPreservingContent()
    {
        var toml = Onboarding.BuildTomlScaffold("P", "Book.xlsm");
        var updated = ManagedStamp.Upsert(toml, "0.3.0");
        Assert.StartsWith(toml.TrimEnd(), updated);
        Assert.EndsWith("[vbaloop]\r\n# vbaloop version that last wrote the managed files. Recorded by vbaloop init - do not edit directly\r\nmanaged_by = \"0.3.0\"\r\n", updated);
        Assert.Equal("0.3.0", ManagedStamp.Read(updated));
    }

    [Fact]
    public void Upsert_ExistingKey_ReplacesInPlace_AndIsIdempotent()
    {
        var toml = "[vbaloop]\nmanaged_by = \"0.2.0\"\n\n[excel]\nworkbook = \"B.xlsm\" # 주석\n";
        var updated = ManagedStamp.Upsert(toml, "0.3.0");
        Assert.Equal("[vbaloop]\nmanaged_by = \"0.3.0\"\n\n[excel]\nworkbook = \"B.xlsm\" # 주석\n", updated); // LF 보존
        Assert.Equal(updated, ManagedStamp.Upsert(updated, "0.3.0"));
    }

    [Fact]
    public void Upsert_SectionWithoutKey_InsertsUnderHeader()
    {
        var toml = "[excel]\r\nworkbook = \"B.xlsm\"\r\n[vbaloop]\r\nfuture = true\r\n[src]\r\nroot = \"src\"\r\n";
        var updated = ManagedStamp.Upsert(toml, "1.0.0");
        Assert.Contains("[vbaloop]\r\nmanaged_by = \"1.0.0\"\r\nfuture = true\r\n[src]", updated);
        Assert.Equal("1.0.0", ManagedStamp.Read(updated));
    }

    [Fact]
    public void Upsert_KeyInOtherSection_Untouched()
    {
        var toml = "[project]\r\nmanaged_by = \"x\"\r\n";
        var updated = ManagedStamp.Upsert(toml, "1.0.0");
        Assert.StartsWith("[project]\r\nmanaged_by = \"x\"\r\n", updated);
        Assert.Equal("1.0.0", ManagedStamp.Read(updated));
    }

    // ---- ProjectConfig 연동 ----

    [Fact]
    public void ProjectConfig_AcceptsStampSection_AndToleratesUnknownKeysInIt()
    {
        using var t = new TempDir();
        var toml = ManagedStamp.Upsert(Onboarding.BuildTomlScaffold("P", "Book.xlsm"), "0.1.0") + "future_key = \"x\"\r\n";
        var path = t.File("vbaloop.toml", toml);
        var cfg = ProjectConfig.Load(path, t.Path);
        Assert.Equal("0.1.0", cfg.ManagedBy);
    }

    [Fact]
    public void ProjectConfig_UnknownSectionWithNewerStamp_IsExeOlderEnvironmentError()
    {
        using var t = new TempDir();
        var toml = ManagedStamp.Upsert(Onboarding.BuildTomlScaffold("P", "Book.xlsm"), "999.0.0")
                   + "[newfeature]\r\nx = true\r\n";
        var path = t.File("vbaloop.toml", toml);
        var e = Assert.Throws<CliError>(() => ProjectConfig.Load(path, t.Path));
        Assert.True(e.Environment);
        Assert.Equal(ManagedStamp.UpdateHint, e.Hint);
        Assert.Contains("999.0.0", e.Message);
    }

    [Fact]
    public void ProjectConfig_UnknownSectionWithOlderStamp_IsOrdinaryError()
    {
        using var t = new TempDir();
        var toml = ManagedStamp.Upsert(Onboarding.BuildTomlScaffold("P", "Book.xlsm"), "0.0.1")
                   + "[typo]\r\nx = true\r\n";
        var path = t.File("vbaloop.toml", toml);
        var e = Assert.Throws<CliError>(() => ProjectConfig.Load(path, t.Path));
        Assert.False(e.Environment);
        Assert.Contains("unknown section", e.Message);
    }

    [Fact]
    public void ProjectConfig_LoadedHook_Fires()
    {
        using var t = new TempDir();
        var path = t.File("vbaloop.toml", Onboarding.BuildTomlScaffold("P", "Book.xlsm"));
        var seen = new System.Collections.Concurrent.ConcurrentBag<ProjectConfig>(); // 병렬 테스트의 Load도 들어올 수 있다
        ProjectConfig.Loaded = seen.Add;
        try { ProjectConfig.Load(path, t.Path); }
        finally { ProjectConfig.Loaded = null; }
        var mine = Assert.Single(seen, c => c.TomlPath == path);
        Assert.Null(mine.ManagedBy);
    }

    // ---- init 명령 수준 ----

    [Fact]
    public void Init_ExeOlderThanStamp_RefusesWithExit3_AndWritesNothing()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var toml = ManagedStamp.Upsert(Onboarding.BuildTomlScaffold("P", "Book.xlsm"), "999.0.0");
        var path = t.File("vbaloop.toml", toml);

        foreach (var check in new[] { false, true })
        {
            var args = check ? new[] { "--config", path, "--check" } : new[] { "--config", path };
            var o = InitCommand.Run(CliArgs.Parse(args));
            Assert.False(o.Ok);
            Assert.Equal(3, o.ExitCode);
            Assert.Equal(ManagedStamp.UpdateHint, o.Hint);
            var payload = JsonSerializer.SerializeToElement(o.Payload, Output.JsonOptions);
            Assert.Equal("exe-older", payload.GetProperty("state").GetString());
            Assert.Equal("999.0.0", payload.GetProperty("managedBy").GetString());
            Assert.Equal(check ? "check" : "apply", payload.GetProperty("mode").GetString());
        }
        Assert.Equal(toml, File.ReadAllText(path));                         // 스탬프 그대로
        Assert.False(File.Exists(Path.Combine(t.Path, "CLAUDE.md")));        // 관리 파일 미생성
        Assert.False(File.Exists(Path.Combine(t.Path, "src", "LogShim.bas")));
    }

    [Fact]
    public void Init_InvalidStamp_RefusesWithExit2()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var path = t.File("vbaloop.toml", ManagedStamp.Upsert(Onboarding.BuildTomlScaffold("P", "Book.xlsm"), "latest"));
        var o = InitCommand.Run(CliArgs.Parse(new[] { "--config", path }));
        Assert.Equal(2, o.ExitCode);
        Assert.Equal("invalid", JsonSerializer.SerializeToElement(o.Payload, Output.JsonOptions).GetProperty("state").GetString());
    }

    [Fact]
    public void Init_Unstamped_ConvergesAndStamps_ThenCheckIsClean()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var path = t.File("vbaloop.toml", Onboarding.BuildTomlScaffold("P", "Book.xlsm"));

        var check = InitCommand.Run(CliArgs.Parse(new[] { "--config", path, "--check" }));
        Assert.Equal(2, check.ExitCode);                                     // 스탬프 없음 = 드리프트
        Assert.Contains(check.HumanLines, l => l.Contains("vbaloop.toml") && l.Contains("recorded"));

        var apply = InitCommand.Run(CliArgs.Parse(new[] { "--config", path }));
        Assert.True(apply.Ok);
        Assert.Equal(ExeVersion.Current!.ToString(), ManagedStamp.ReadFile(path));

        var again = InitCommand.Run(CliArgs.Parse(new[] { "--config", path, "--check" }));
        Assert.True(again.Ok);
        Assert.Equal("ok", JsonSerializer.SerializeToElement(again.Payload, Output.JsonOptions).GetProperty("state").GetString());
    }

    [Fact]
    public void Init_OlderStampAndNoDrift_LeavesTomlAlone()
    {
        // exe가 더 새롭지만 관리 파일이 이미 같으면 toml을 건드리지 않는다(갱신마다 생기는 churn 방지)
        using var t = new TempDir();
        t.File("Book.xlsm");
        var path = t.File("vbaloop.toml", Onboarding.BuildTomlScaffold("P", "Book.xlsm"));
        Assert.True(InitCommand.Run(CliArgs.Parse(new[] { "--config", path })).Ok);
        var lowered = ManagedStamp.Upsert(File.ReadAllText(path), "0.0.1");
        File.WriteAllText(path, lowered);

        var o = InitCommand.Run(CliArgs.Parse(new[] { "--config", path }));
        Assert.True(o.Ok);
        Assert.Equal(lowered, File.ReadAllText(path));

        // 관리 파일이 바뀌면 그때 스탬프를 올린다
        File.WriteAllText(Path.Combine(t.Path, "src", "LogShim.bas"), "old\r\n");
        Assert.True(InitCommand.Run(CliArgs.Parse(new[] { "--config", path })).Ok);
        Assert.Equal(ExeVersion.Current!.ToString(), ManagedStamp.ReadFile(path));
    }

    // ---- Onboarding 순수 로직 ----

    private static readonly string[] Shim = { "Attribute VB_Name = \"LogShim\"", "Option Explicit" };
    private static readonly string[] Body = { "## 개발 루프 (vbaloop)" };
    private static readonly string[] Doc = { "# doc" };

    [Fact]
    public void Analyze_Scaffold_IncludesStamp()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var items = Onboarding.Analyze(t.Path, null, Shim, Body, Doc, Doc, Doc, stampVersion: "0.5.0");
        Assert.Equal("0.5.0", ManagedStamp.Read(items[0].NewContent!));
        Assert.Contains("workbook = \"Book.xlsm\"", items[0].NewContent);
    }

    [Fact]
    public void Analyze_NeverLowersStamp()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var path = t.File("vbaloop.toml", ManagedStamp.Upsert(Onboarding.BuildTomlScaffold("P", "Book.xlsm"), "0.9.0"));
        // 다른 관리 파일이 전부 Missing(쓰기 발생)이어도 더 낮은 버전으로는 기록하지 않는다
        var items = Onboarding.Analyze(t.Path, path, Shim, Body, Doc, Doc, Doc, stampVersion: "0.5.0");
        Assert.Equal(ArtifactStatus.Ok, items[0].Status);
        Assert.Null(items[0].NewContent);
    }

    [Fact]
    public void Analyze_NewerExeWithDrift_RaisesStampWithDetail()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var path = t.File("vbaloop.toml", ManagedStamp.Upsert(Onboarding.BuildTomlScaffold("P", "Book.xlsm"), "0.4.0"));
        var items = Onboarding.Analyze(t.Path, path, Shim, Body, Doc, Doc, Doc, stampVersion: "0.5.0");
        Assert.Equal(ArtifactStatus.Drift, items[0].Status);
        Assert.Equal("managed_by 0.4.0 -> 0.5.0", items[0].Detail);
        Onboarding.Apply(items);
        Assert.Equal("0.5.0", ManagedStamp.ReadFile(path));
        Assert.Equal("P", ProjectConfig.Load(path, t.Path).ProjectName); // 사용자 내용 보존
    }

    [Fact]
    public void Analyze_NoStampVersion_LeavesTomlOk()
    {
        using var t = new TempDir();
        t.File("Book.xlsm");
        var path = t.File("vbaloop.toml", Onboarding.BuildTomlScaffold("P", "Book.xlsm"));
        var items = Onboarding.Analyze(t.Path, path, Shim, Body, Doc, Doc, Doc);
        Assert.Equal(ArtifactStatus.Ok, items[0].Status);
    }

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
}
