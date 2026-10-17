using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class SetupPlanTests
{
    private static readonly SetupPlan.Probes AllGood =
        new(ExcelOk: true, ExcelVersion: "Excel.Application.16", VbomOk: true, VbomDetail: "AccessVBOM=1",
            Installed: true, RunningInstalled: true, OnUserPath: true);

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
    }

    private static SetupPlan.Request Req(string? workbook = null, string? newFolder = null, string? folder = null,
        bool installOnly = false, bool noInstall = false, bool skipVerify = false)
        => new(workbook, newFolder, folder, installOnly, noInstall, skipVerify);

    [Fact]
    public void Workbook_AllGood_ReadyWithCopyProjectAndNoConsent()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");

        var r = SetupPlan.Build(Req(workbook: wb), AllGood);

        Assert.True(r.Ready);
        Assert.Empty(r.Consents);
        Assert.Equal("workbook", r.Project!.Start);
        Assert.Equal("copy", r.Project.Action);
        Assert.Equal("pull", r.Project.Verify);
        Assert.Equal(System.IO.Path.Combine(t.Path, "Book"), r.Project.Folder);
        Assert.Equal($"vbaloop setup --workbook \"{wb}\" --yes --json", r.Command);
    }

    [Fact]
    public void Plan_DoesNotTouchFileSystem()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");
        SetupPlan.Build(Req(workbook: wb), AllGood);
        Assert.False(Directory.Exists(System.IO.Path.Combine(t.Path, "Book")));
    }

    [Fact]
    public void VbomOff_ConsentWithApproveFlag_NotInCommand()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");

        var r = SetupPlan.Build(Req(workbook: wb), AllGood with { VbomOk = false });

        Assert.True(r.Ready); // 동의는 막는 요인이 아니다 — 에이전트가 묻고 플래그를 붙인다
        var c = Assert.Single(r.Consents);
        Assert.Equal("accessVbom", c.Id);
        Assert.Equal("--enable-vbom", c.ApproveFlag);
        Assert.Null(c.DeclineFlag);
        Assert.DoesNotContain("--enable-vbom", r.Command);
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void NotInstalledOrNotOnPath_InstallConsent_DeclinableWithNoInstall(bool runningInstalled, bool onPath)
    {
        var r = SetupPlan.Build(Req(installOnly: true),
            AllGood with { RunningInstalled = runningInstalled, OnUserPath = onPath });

        var c = Assert.Single(r.Consents);
        Assert.Equal("install", c.Id);
        Assert.Null(c.ApproveFlag);
        Assert.Equal("--no-install", c.DeclineFlag);
    }

    [Fact]
    public void NoInstall_NoInstallConsent_FlagCarriedIntoCommand()
    {
        var r = SetupPlan.Build(Req(installOnly: true, noInstall: true),
            AllGood with { RunningInstalled = false, OnUserPath = false });

        Assert.Empty(r.Consents);
        Assert.Equal("vbaloop setup --install-only --no-install --yes --json", r.Command);
    }

    [Fact]
    public void NoStart_NotReady_WithHint()
    {
        var r = SetupPlan.Build(Req(), AllGood);
        Assert.False(r.Ready);
        Assert.Null(r.Project);
        Assert.Contains("--workbook", r.Hint);
        Assert.False(r.EnvironmentBlocker);
    }

    [Fact]
    public void NoExcel_EnvironmentBlocker_ButProjectStillDescribed()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");

        var r = SetupPlan.Build(Req(workbook: wb), AllGood with { ExcelOk = false });

        Assert.False(r.Ready);
        Assert.True(r.EnvironmentBlocker);
        Assert.NotNull(r.Project);
    }

    [Fact]
    public void ProjectConflict_BecomesBlocker()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm", "original");
        t.File(@"dest\Other.xlsm");

        var r = SetupPlan.Build(Req(workbook: wb, folder: System.IO.Path.Combine(t.Path, "dest")), AllGood);

        Assert.False(r.Ready);
        Assert.Null(r.Project);
        Assert.Single(r.Blockers);
    }

    [Fact]
    public void NewProject_CommandUsesAbsoluteFolder()
    {
        using var t = new TempDir();
        string folder = System.IO.Path.Combine(t.Path, "MyTool");

        var r = SetupPlan.Build(Req(newFolder: folder), AllGood);

        Assert.Equal("new", r.Project!.Start);
        Assert.Equal("create", r.Project.Action);
        Assert.Equal("MyTool.xlsm", r.Project.Workbook);
        Assert.Equal("push+test", r.Project.Verify);
        Assert.Equal($"vbaloop setup --folder \"{folder}\" --yes --json", r.Command);
    }

    [Theory]
    [InlineData(0, "create", "push+test")]
    [InlineData(1, "adopt", "pull")]
    public void Folder_ZeroOrOneWorkbook(int workbooks, string action, string verify)
    {
        using var t = new TempDir();
        if (workbooks == 1) t.File(@"proj\Main.xlsm");
        string folder = System.IO.Path.Combine(t.Path, "proj");

        var p = SetupPlan.DescribeProject(Req(folder: folder))!;

        Assert.Equal(action, p.Action);
        Assert.Equal(verify, p.Verify);
    }

    [Fact]
    public void Folder_MultipleWorkbooks_Throws()
    {
        using var t = new TempDir();
        t.File(@"proj\A.xlsm");
        t.File(@"proj\B.xlsm");
        Assert.Throws<CliError>(() => SetupPlan.DescribeProject(Req(folder: System.IO.Path.Combine(t.Path, "proj"))));
    }

    [Fact]
    public void SkipVerify_ReflectedInProjectAndCommand()
    {
        using var t = new TempDir();
        var wb = t.File("Book.xlsm");

        var r = SetupPlan.Build(Req(workbook: wb, skipVerify: true), AllGood);

        Assert.Equal("skipped", r.Project!.Verify);
        Assert.EndsWith("--skip-verify --yes --json", r.Command);
    }
}
