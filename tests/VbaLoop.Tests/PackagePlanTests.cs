using VbaLoop.Core;
using Xunit;
using static VbaLoop.Core.PackagePlan;

namespace VbaLoop.Tests;

public class PackagePlanTests
{
    private static PackageFormat.Module Mod(string name, string ext, PackageModuleKind kind, params string[] body)
    {
        var header = ext switch
        {
            ".bas" => new[] { $"Attribute VB_Name = \"{name}\"" },
            ".frm" => new[] { "VERSION 5.00", $"Begin {{C62A69F0-16DC-11CE-9E98-00AA00574A4F}} {name}", "End", $"Attribute VB_Name = \"{name}\"" },
            _ => VbaHeader.SynthesizeHeader(kind == PackageModuleKind.Class ? VbaComponentKind.ClassModule : VbaComponentKind.Document, name),
        };
        return new PackageFormat.Module(name, ext, kind, header.Concat(body).ToArray(), kind == PackageModuleKind.Form ? new byte[] { 1 } : null);
    }

    private static WorkbookComponent Wb(string name, VbaComponentKind? kind, string body = "", bool isWorkbook = false)
        => new(name, kind, isWorkbook, body);

    private static readonly WorkbookComponent[] Korean =
    {
        Wb("현재_통합_문서", VbaComponentKind.Document, "Option Explicit", isWorkbook: true),
        Wb("Sheet1", VbaComponentKind.Document, "Option Explicit"),
        Wb("Sheet9", VbaComponentKind.Document, "' 현장에서 추가"),
        Wb("modMain", VbaComponentKind.StdModule, "Option Explicit\r\nSub A()\r\nEnd Sub"),
        Wb("modOld", VbaComponentKind.StdModule, "Option Explicit"),
        Wb("frmInput", VbaComponentKind.MsForm, "Option Explicit"),
    };

    private static readonly PackageFormat.Module[] Package =
    {
        Mod("ThisWorkbook", ".cls", PackageModuleKind.Workbook, "Option Explicit"),
        Mod("Sheet1", ".cls", PackageModuleKind.Document, "Option Explicit"),
        Mod("modMain", ".bas", PackageModuleKind.Std, "Option Explicit", "Sub A()", "End Sub"),
        Mod("clsNew", ".cls", PackageModuleKind.Class, "Option Explicit"),
        Mod("frmInput", ".frm", PackageModuleKind.Form, "Option Explicit"),
    };

    private static StepAction ActionOf(Result r, string target) => r.Steps.Single(s => s.Target == target).Action;

    [Fact]
    public void Build_MapsEveryComponent()
    {
        var r = Build(Korean, Package);
        Assert.Empty(r.Errors);
        Assert.Equal(StepAction.Inject, ActionOf(r, "현재_통합_문서"));   // 통합 문서는 종류로 맞춘다
        Assert.Equal(StepAction.Inject, ActionOf(r, "Sheet1"));
        Assert.Equal(StepAction.IgnoreDocument, ActionOf(r, "Sheet9"));   // 현장 시트는 건드리지 않는다
        Assert.Equal(StepAction.Replace, ActionOf(r, "modMain"));
        Assert.Equal(StepAction.Remove, ActionOf(r, "modOld"));
        Assert.Equal(StepAction.Add, ActionOf(r, "clsNew"));
        Assert.Equal(StepAction.ReimportForm, ActionOf(r, "frmInput"));
        Assert.Equal(new[] { "Sub A()", "End Sub" }, r.Steps.Single(s => s.Target == "modMain").Body.Skip(1));
    }

    [Fact]
    public void Build_MissingSheet_IsError()
    {
        var r = Build(Korean, Package.Append(Mod("Sheet2", ".cls", PackageModuleKind.Document)).ToArray());
        Assert.Contains(r.Errors, e => e.Contains("Sheet2") && e.Contains("cannot be created"));
    }

    [Fact]
    public void Build_KindMismatch_IsError()
    {
        var r = Build(Korean, new[] { Mod("modMain", ".cls", PackageModuleKind.Class) });
        Assert.Single(r.Errors);
    }

    [Fact]
    public void Build_UnsupportedComponent_LeftAlone()
    {
        var r = Build(Korean.Append(Wb("Designer1", null)).ToArray(), Package);
        Assert.DoesNotContain(r.Steps, s => s.Target == "Designer1");
    }

    [Fact]
    public void CodeDifferences_SameCode_Empty()
    {
        var wb = Korean.Where(c => c.Name != "modOld").Append(Wb("clsNew", VbaComponentKind.ClassModule, "Option Explicit")).ToArray();
        var plan = Build(wb, Package);
        Assert.Empty(CodeDifferences(plan, wb));
    }

    [Fact]
    public void CodeDifferences_ReportsSetAndBodyChanges()
    {
        var plan = Build(Korean, Package);
        var diffs = CodeDifferences(plan, Korean);
        Assert.Contains("modOld", diffs);   // 제거 대상
        Assert.Contains("clsNew", diffs);   // 추가 대상
        Assert.DoesNotContain("Sheet9", diffs);   // 워크북에만 있는 시트는 비교에서 뺀다
        Assert.DoesNotContain("modMain", diffs);
    }

    [Fact]
    public void CodeDifferences_IdentifierCasingOnly_IsSame()
    {
        var wb = new[] { Wb("modMain", VbaComponentKind.StdModule, "option explicit\r\nsub a()\r\nend sub\r\n\r\n") };
        var plan = Build(wb, new[] { Mod("modMain", ".bas", PackageModuleKind.Std, "Option Explicit", "Sub A()", "End Sub") });
        Assert.Empty(CodeDifferences(plan, wb));
    }

    [Fact]
    public void CodeDifferences_StringLiteralChange_IsDifferent()
    {
        var wb = new[] { Wb("modMain", VbaComponentKind.StdModule, "Const S = \"abc\"") };
        var plan = Build(wb, new[] { Mod("modMain", ".bas", PackageModuleKind.Std, "Const S = \"ABC\"") });
        Assert.Equal(new[] { "modMain" }, CodeDifferences(plan, wb));
    }

    [Fact]
    public void CodeDifferences_FormLeadingBlankLine_IsSame()
    {
        var wb = new[] { Wb("frmInput", VbaComponentKind.MsForm, "\r\nOption Explicit") };
        var plan = Build(wb, new[] { Mod("frmInput", ".frm", PackageModuleKind.Form, "Option Explicit") });
        Assert.Empty(CodeDifferences(plan, wb));
    }

    [Fact]
    public void CodeDifferences_BuildInfoStampOnly_IsSame()
    {
        // 같은 src를 다시 빌드하면 합성 modBuildInfo의 BuiltAt만 달라진다 — 코드 차이가 아니다
        var wb = new[] { Wb("modBuildInfo", VbaComponentKind.StdModule, "Public Const BUILT_AT As String = \"2026-01-01\"") };
        var plan = Build(wb, new[] { Mod("modBuildInfo", ".bas", PackageModuleKind.Std, "Public Const BUILT_AT As String = \"2026-02-02\"") });
        Assert.Empty(CodeDifferences(plan, wb));
    }
}
