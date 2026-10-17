using VbaLoop.Cli;
using Xunit;

namespace VbaLoop.Tests;

public class CliArgsTests
{
    [Theory]
    // 파싱이 실패해도(--오타플래그) raw argv의 --json은 감지돼야 한다
    [InlineData(new[] { "push", "--json", "--badflag" }, true)]
    // --json이 없으면 false (기존 평문 경로 유지)
    [InlineData(new[] { "push", "--prune" }, false)]
    // 값-옵션이 --json을 값으로 소비하면 JSON 요청이 아니다
    [InlineData(new[] { "test", "--config", "--json" }, false)]
    // --json이 verb 위치에 온 드문 경로도 감지
    [InlineData(new[] { "--json" }, true)]
    // positional이 플래그 앞에 와도 전체를 훑는다
    [InlineData(new[] { "run", "modA.Proc", "--json" }, true)]
    public void RawArgsRequestJson_Detects(string[] argv, bool expected)
    {
        Assert.Equal(expected, CliArgs.RawArgsRequestJson(argv));
    }

    [Fact]
    public void RawArgsRequestJson_EmptyArgv_False()
    {
        Assert.False(CliArgs.RawArgsRequestJson(System.Array.Empty<string>()));
    }

    [Fact]
    public void Parse_KillOrphansAndForce_Flags()
    {
        var r = CliArgs.Parse(new[] { "--kill-orphans", "--force" });
        Assert.True(r.KillOrphans);
        Assert.True(r.Force);
        Assert.False(r.Json);
    }

    [Fact]
    public void Parse_Discard_Flag()
    {
        var r = CliArgs.Parse(new[] { "stop", "--discard" });
        Assert.True(r.Discard);
        Assert.Equal(new[] { "stop" }, r.Positional);
        Assert.False(CliArgs.Parse(new[] { "stop" }).Discard);
    }

    [Fact]
    public void Parse_AllAndList_Flags()
    {
        var r = CliArgs.Parse(new[] { "--all", "--check" });
        Assert.True(r.All);
        Assert.True(r.Check);
        Assert.False(r.List);

        var r2 = CliArgs.Parse(new[] { "--list", "--json" });
        Assert.True(r2.List);
        Assert.True(r2.Json);
        Assert.False(r2.All);
    }

    [Fact]
    public void Parse_Module_TakesValue()
    {
        var r = CliArgs.Parse(new[] { "--module", "modActions", "--json" });
        Assert.Equal("modActions", r.Module);
        Assert.True(r.Json);
    }

    [Fact]
    public void Parse_Module_MissingValue_Throws()
    {
        Assert.Throws<VbaLoop.Core.CliError>(() => CliArgs.Parse(new[] { "--module" }));
    }

    [Fact]
    public void RawArgsRequestJson_ModuleValue_NotMistakenForJson()
    {
        // --module의 값이 우연히 "--json"이어도 JSON 요청으로 오인하지 않는다
        Assert.False(CliArgs.RawArgsRequestJson(new[] { "diff", "--module", "--json" }));
    }

    [Fact]
    public void Parse_SetupFlags()
    {
        var r = CliArgs.Parse(new[] { "--folder", @"C:\work\MyVba", "--yes", "--skip-verify" });
        Assert.Equal(@"C:\work\MyVba", r.Folder);
        Assert.True(r.Yes);
        Assert.True(r.SkipVerify);
        Assert.False(r.InstallOnly);
    }

    [Fact]
    public void Parse_YesShortForm()
        => Assert.True(CliArgs.Parse(new[] { "-y" }).Yes);

    [Fact]
    public void Parse_InstallOnlyAndFix()
    {
        var r = CliArgs.Parse(new[] { "--install-only", "--fix" });
        Assert.True(r.InstallOnly);
        Assert.True(r.Fix);
    }

    [Fact]
    public void Parse_Folder_MissingValue_Throws()
        => Assert.Throws<VbaLoop.Core.CliError>(() => CliArgs.Parse(new[] { "--folder" }));

    [Fact]
    public void RawArgsRequestJson_FolderValue_NotMistakenForJson()
    {
        // 경로 자리에 "--json"이 와도 JSON 요청으로 오인하지 않는다(값-옵션 건너뛰기)
        Assert.False(CliArgs.RawArgsRequestJson(new[] { "setup", "--folder", "--json" }));
    }

    [Fact]
    public void Parse_RollbackForms()
    {
        var latest = CliArgs.Parse(new[] { "--latest", "--json" });
        Assert.True(latest.Latest);
        Assert.Null(latest.Backup);
        var byId = CliArgs.Parse(new[] { "--backup", "20261017-031500-123-push" });
        Assert.Equal("20261017-031500-123-push", byId.Backup);
        Assert.False(byId.Latest);
    }

    [Fact]
    public void Parse_BackupWithoutValue_Fails()
        => Assert.Throws<VbaLoop.Core.CliError>(() => CliArgs.Parse(new[] { "--backup" }));

    [Fact]
    public void RawArgsRequestJson_BackupConsumesValue()
        => Assert.False(CliArgs.RawArgsRequestJson(new[] { "rollback", "--backup", "--json" }));

    [Fact]
    public void Parse_SkipCompile()
        => Assert.True(CliArgs.Parse(new[] { "--skip-compile" }).SkipCompile);

    [Fact]
    public void Parse_SetupWorkbookAndNew()
    {
        var wb = CliArgs.Parse(new[] { "--workbook", @"C:\work\Book.xlsm", "--folder", @"D:\dev\Book" });
        Assert.Equal(@"C:\work\Book.xlsm", wb.Workbook);
        Assert.Equal(@"D:\dev\Book", wb.Folder);
        Assert.Equal("MyTool", CliArgs.Parse(new[] { "--new", "MyTool" }).NewName);
    }

    [Fact]
    public void Positionals_SkipOptionValues()
    {
        // 축약형 "vbaloop Book.xlsm --folder D:\dest --config C:\r\vbaloop.toml --json"의 워크북은 하나다
        var argv = new[] { @"C:\w\Book.xlsm", "--folder", @"D:\dest", "--config", @"C:\r\vbaloop.toml", "--json" };
        Assert.Equal(new[] { @"C:\w\Book.xlsm" }, CliArgs.Positionals(argv));
        Assert.Equal(new[] { "a.xlsm", "b.xlsm" }, CliArgs.Positionals(new[] { "a.xlsm", "b.xlsm" }));
    }

    [Fact]
    public void Parse_EnableVbomAndPlan_AreSeparateFromYes()
    {
        var yesOnly = CliArgs.Parse(new[] { "--yes" });
        Assert.False(yesOnly.EnableVbom); // --yes는 보안 설정 동의를 포함하지 않는다
        var r = CliArgs.Parse(new[] { "--plan", "--enable-vbom" });
        Assert.True(r.Plan);
        Assert.True(r.EnableVbom);
        Assert.False(r.Yes);
    }

    [Theory]
    [InlineData("--workbook")]
    [InlineData("--new")]
    public void Parse_SetupValueOption_MissingValue_Throws(string opt)
        => Assert.Throws<VbaLoop.Core.CliError>(() => CliArgs.Parse(new[] { opt }));

    [Theory]
    [InlineData("--workbook")]
    [InlineData("--new")]
    public void RawArgsRequestJson_SetupValueOption_ConsumesValue(string opt)
        => Assert.False(CliArgs.RawArgsRequestJson(new[] { "setup", opt, "--json" }));

    [Fact]
    public void SessionMeta_Parsed()
    {
        var a = CliArgs.Parse(new[] { "__runner", "--session-meta", "_dbg/ephemeral/x.json" });
        Assert.Equal("_dbg/ephemeral/x.json", a.SessionMeta);
        Assert.Throws<VbaLoop.Core.CliError>(() => CliArgs.Parse(new[] { "__runner", "--session-meta" }));
    }
}
