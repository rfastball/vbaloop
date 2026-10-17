using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ModuleDiffTests
{
    private static byte[] Bytes(params string[] lines)
        => CanonicalText.Utf8NoBom.GetBytes(CanonicalText.ToCanonical(lines));

    private static string Body(params string[] lines) => string.Join("\r\n", lines);

    private static ModuleDiff.Entry Single(
        ModuleDiff.WorkbookModule[] wb, ModuleDiff.SrcFile[] src)
        => Assert.Single(ModuleDiff.Compare(wb, src));

    [Fact]
    public void ByteIdenticalSrc_ReportsSame()
    {
        // pull이 다시 써도 바이트가 같을 src — same
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, Body("Option Explicit", "", "Sub A()", "End Sub")) },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", Bytes("Attribute VB_Name = \"modA\"", "Option Explicit", "", "Sub A()", "End Sub")) });
        Assert.Equal("same", e.Verdict);
        Assert.Equal("modA.bas", e.File);
        Assert.Equal("stdModule", e.Kind);
    }

    [Fact]
    public void BodyEdit_ReportsChanged_WithSrcFileRelativeFirstDiffLine_AndExcerpts()
    {
        // 헤더 1행 + 본문 2행째가 다름 — 파일 기준 3행
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, Body("Option Explicit", "n = n + 1")) },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", Bytes("Attribute VB_Name = \"modA\"", "Option Explicit", "n = n + 2")) });
        Assert.Equal("changed", e.Verdict);
        Assert.Equal(3, e.Detail!.FirstDiffLine);
        Assert.Equal("n = n + 1", e.Detail.WorkbookLine);
        Assert.Equal("n = n + 2", e.Detail.SrcLine);
        Assert.Equal(3, e.Detail.WorkbookLines);
        Assert.Equal(3, e.Detail.SrcLines);
        Assert.False(e.Detail.FormatOnly);
    }

    [Fact]
    public void MissingSrcFile_ReportsOnlyInWorkbook_WithLineCount()
    {
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modNew", VbaComponentKind.StdModule, Body("Option Explicit", "Sub B()", "End Sub")) },
            System.Array.Empty<ModuleDiff.SrcFile>());
        Assert.Equal("onlyInWorkbook", e.Verdict);
        Assert.Null(e.File);
        Assert.Equal(3, e.Detail!.WorkbookLines);
    }

    [Fact]
    public void SrcWithoutComponent_ReportsOnlyInSrc()
    {
        var e = Single(
            System.Array.Empty<ModuleDiff.WorkbookModule>(),
            new[] { new ModuleDiff.SrcFile("OldHelper", ".bas", Bytes("Attribute VB_Name = \"OldHelper\"")) });
        Assert.Equal("onlyInSrc", e.Verdict);
        Assert.Equal("OldHelper.bas", e.File);
        Assert.Equal("stdModule", e.Kind);
    }

    [Fact]
    public void BomViolation_SameLines_ReportsChangedFormatOnly()
    {
        // 라인은 동일하지만 선행 BOM 3바이트 — pull이 재정규화할 파일. formatOnly로 구분
        var canonical = Bytes("Attribute VB_Name = \"modA\"", "Option Explicit");
        var withBom = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(canonical).ToArray();
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, "Option Explicit") },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", withBom) });
        Assert.Equal("changed", e.Verdict);
        Assert.True(e.Detail!.FormatOnly);
        Assert.Null(e.Detail.FirstDiffLine);
    }

    [Fact]
    public void LfViolation_SameLines_ReportsChangedFormatOnly()
    {
        var lfBytes = CanonicalText.Utf8NoBom.GetBytes("Attribute VB_Name = \"modA\"\nOption Explicit\n");
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, "Option Explicit") },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", lfBytes) });
        Assert.Equal("changed", e.Verdict);
        Assert.True(e.Detail!.FormatOnly);
    }

    [Fact]
    public void Form_ComparesBodyOnly_TrailingEmptyConverges_LayoutNotCompared()
    {
        // 워크북 코드비하인드 끝의 빈 줄(VBE Import 동작)은 비교 전 제거돼 same으로 수렴
        var frm = Bytes(
            "VERSION 5.00",
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 ",
            "   Caption         =   \"f\"",
            "End",
            "Attribute VB_Name = \"UserForm1\"",
            "Option Explicit",
            "Sub UserForm_Click()",
            "End Sub");
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("UserForm1", VbaComponentKind.MsForm, Body("Option Explicit", "Sub UserForm_Click()", "End Sub", "", "")) },
            new[] { new ModuleDiff.SrcFile("UserForm1", ".frm", frm) });
        Assert.Equal("same", e.Verdict);
        Assert.Equal("form", e.Kind);
        Assert.False(e.Detail!.LayoutCompared);
    }

    [Fact]
    public void Form_LeadingBlankInWorkbookCodeModule_Converges()
    {
        // 폼 CodeModule은 선언부 자리의 선행 빈 줄을 포함하지만 Export 텍스트(src .frm)에는
        // 없다(실측) — pull이 다시 쓰지 않을 파일이므로 same이어야 한다
        var frm = Bytes(
            "VERSION 5.00",
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 ",
            "End",
            "Attribute VB_Name = \"UserForm1\"",
            "Option Explicit");
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("UserForm1", VbaComponentKind.MsForm, Body("", "Option Explicit")) },
            new[] { new ModuleDiff.SrcFile("UserForm1", ".frm", frm) });
        Assert.Equal("same", e.Verdict);
    }

    [Fact]
    public void Form_BodyEdit_FirstDiffLineIsFileRelative()
    {
        // 헤더 5행(VERSION+Begin블록3+Attribute) 뒤 본문 2행째 차이 — 파일 기준 7행
        var frm = Bytes(
            "VERSION 5.00",
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 ",
            "   Caption         =   \"f\"",
            "End",
            "Attribute VB_Name = \"UserForm1\"",
            "Option Explicit",
            "n = 2");
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("UserForm1", VbaComponentKind.MsForm, Body("Option Explicit", "n = 1")) },
            new[] { new ModuleDiff.SrcFile("UserForm1", ".frm", frm) });
        Assert.Equal("changed", e.Verdict);
        Assert.Equal(7, e.Detail!.FirstDiffLine);
        Assert.Equal("n = 1", e.Detail.WorkbookLine);
        Assert.Equal("n = 2", e.Detail.SrcLine);
        Assert.False(e.Detail.LayoutCompared);
    }

    [Fact]
    public void EmptyModule_HeaderOnlyClsFile_ReportsSame()
    {
        var header = VbaHeader.SynthesizeHeader(VbaComponentKind.Document, "Sheet1");
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("Sheet1", VbaComponentKind.Document, "") },
            new[] { new ModuleDiff.SrcFile("Sheet1", ".cls", Bytes(header)) });
        Assert.Equal("same", e.Verdict);
        Assert.Equal("document", e.Kind);
    }

    [Fact]
    public void SrcHeaderMissingVbName_FallsBackToSynthesizedHeader()
    {
        // 헤더 없는 src: 합성 헤더(1행)와 코드 1행이 어긋남 — 파일 기준 1행부터 changed
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, "Option Explicit") },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", Bytes("Option Explicit")) });
        Assert.Equal("changed", e.Verdict);
        Assert.Equal(1, e.Detail!.FirstDiffLine);
        Assert.Equal("Attribute VB_Name = \"modA\"", e.Detail.WorkbookLine);
    }

    [Fact]
    public void NameMatch_IsCaseInsensitive()
    {
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("ModA", VbaComponentKind.StdModule, "Option Explicit") },
            new[] { new ModuleDiff.SrcFile("moda", ".bas", Bytes("Attribute VB_Name = \"ModA\"", "Option Explicit")) });
        Assert.Equal("same", e.Verdict);
    }

    [Fact]
    public void UnsupportedType_ReportsSkipped()
    {
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("weird", null, "", RawType: 11) },
            System.Array.Empty<ModuleDiff.SrcFile>());
        Assert.Equal("skipped", e.Verdict);
        Assert.Contains("11", e.Note);
    }

    [Fact]
    public void VendoredModule_ReportsManaged_EvenWhenBodyDiffers()
    {
        // LogShim은 정본이 init — COM 왕복이 본문을 미세 변형해도 changed가 아니라 managed로
        // 보고돼 clean 판정에서 빠진다(이슈 D). src 무결성은 init --check 담당.
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("LogShim", VbaComponentKind.StdModule, Body("Option Explicit", "' 워크북 왕복으로 미세 변형된 본문")) },
            new[] { new ModuleDiff.SrcFile("LogShim", ".bas", Bytes("Attribute VB_Name = \"LogShim\"", "Option Explicit", "' 원래 본문")) });
        Assert.Equal("managed", e.Verdict);
        Assert.Null(e.Detail);
        Assert.Contains("init", e.Note);
    }

    [Fact]
    public void CasingOnlyDifference_ReportsCasingOnly_NotChanged()
    {
        // 워크북 본문이 src와 케이싱만 다름(.path vs .Path) — pull이
        // src 철자를 보존하므로 changed가 아니라 casingOnly로 보고돼 clean 판정에서 빠진다.
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, Body("Option Explicit", "x = ThisWorkbook.path")) },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", Bytes("Attribute VB_Name = \"modA\"", "Option Explicit", "x = ThisWorkbook.Path")) });
        Assert.Equal("casingOnly", e.Verdict);
        Assert.False(e.Detail!.FormatOnly);
        Assert.Equal(3, e.Detail.FirstDiffLine);
        Assert.Contains("casing", e.Note);
    }

    [Fact]
    public void CasingPlusRealChange_StaysChanged()
    {
        // 케이싱 차이 + 진짜 내용 변경이 섞이면 casingOnly가 아니라 changed
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, Body("Option Explicit", "x = a.path", "n = 1")) },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", Bytes("Attribute VB_Name = \"modA\"", "Option Explicit", "x = a.Path", "n = 2")) });
        Assert.Equal("changed", e.Verdict);
    }

    [Fact]
    public void StringLiteralOnlyChange_ReportsChanged_NotCasingOnly()
    {
        // 문자열 리터럴 내용 변경은 케이싱이 아니라 실변경 — 전체-라인 OrdinalIgnoreCase 미봉책이
        // casingOnly로 숨기던 것을 방어(토큰 등가관계는 문자열을 gap으로 Ordinal 비교).
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, Body("Option Explicit", "s = \"hello\"")) },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", Bytes("Attribute VB_Name = \"modA\"", "Option Explicit", "s = \"Hello\"")) });
        Assert.Equal("changed", e.Verdict);
    }

    [Fact]
    public void CommentOnlyChange_ReportsChanged_NotCasingOnly()
    {
        // 주석 텍스트 변경도 의미적 — casingOnly로 숨기지 않는다
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, Body("Option Explicit", "x = 1 ' note")) },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", Bytes("Attribute VB_Name = \"modA\"", "Option Explicit", "x = 1 ' Note")) });
        Assert.Equal("changed", e.Verdict);
    }

    [Fact]
    public void InsertPlusRecasing_ReportsChanged_NotCasingOnly()
    {
        // 라인 삽입(라인 수 불일치)은 구조 변경 — casingOnly 아님
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, Body("Option Explicit", "Dim q", "x = a.path")) },
            new[] { new ModuleDiff.SrcFile("modA", ".bas", Bytes("Attribute VB_Name = \"modA\"", "Option Explicit", "x = a.Path")) });
        Assert.Equal("changed", e.Verdict);
    }

    [Fact]
    public void Form_CasingOnlyCodeBehind_ReportsCasingOnly()
    {
        // 폼도 코드비하인드 재케이싱이 casingOnly로 보고돼 clean에서 빠진다(결함 #3 diff 측).
        var frm = Bytes(
            "VERSION 5.00",
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 ",
            "End",
            "Attribute VB_Name = \"UserForm1\"",
            "Option Explicit",
            "x = ThisWorkbook.Path");
        var e = Single(
            new[] { new ModuleDiff.WorkbookModule("UserForm1", VbaComponentKind.MsForm, Body("Option Explicit", "x = ThisWorkbook.path")) },
            new[] { new ModuleDiff.SrcFile("UserForm1", ".frm", frm) });
        Assert.Equal("casingOnly", e.Verdict);
    }

    [Fact]
    public void ExtensionMismatch_ClassFileForStdModule_NotMatched()
    {
        // 컴포넌트는 .bas를 기대 — .cls 파일은 매칭되지 않아 onlyInWorkbook.
        // 이름은 컴포넌트와 겹치므로 pull orphan 규칙대로 onlyInSrc도 아님
        var entries = ModuleDiff.Compare(
            new[] { new ModuleDiff.WorkbookModule("modA", VbaComponentKind.StdModule, "Option Explicit") },
            new[] { new ModuleDiff.SrcFile("modA", ".cls", Bytes("Attribute VB_Name = \"modA\"", "Option Explicit")) });
        var e = Assert.Single(entries);
        Assert.Equal("onlyInWorkbook", e.Verdict);
    }
}
