using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class VbaHeaderTests
{
    [Fact]
    public void Split_Bas_HeaderAndBody()
    {
        var lines = new[]
        {
            "Attribute VB_Name = \"modFoo\"",
            "' @desc 테스트 모듈",
            "Option Explicit",
            "",
            "Public Sub Foo()",
            "End Sub",
        };
        var r = VbaHeader.Split(lines, "modFoo.bas");
        Assert.Equal(new[] { "Attribute VB_Name = \"modFoo\"" }, r.Header);
        Assert.Equal(5, r.Body.Count);
        Assert.Equal("' @desc 테스트 모듈", r.Body[0]);
    }

    [Fact]
    public void Split_Cls_StandardNineLineHeader()
    {
        var header = VbaHeader.SynthesizeHeader(VbaComponentKind.ClassModule, "clsBar");
        var lines = header.Concat(new[] { "Option Explicit", "Private m_x As Long" }).ToArray();
        var r = VbaHeader.Split(lines, "clsBar.cls");
        Assert.Equal(9, r.Header.Count);
        Assert.Equal(2, r.Body.Count);
        Assert.Equal("Option Explicit", r.Body[0]);
    }

    [Fact]
    public void Split_EmptyModule_HeaderOnly()
    {
        // 현재_통합_문서.cls 케이스: 헤더 9줄, 본문 0줄 (설계 §4 빈 모듈 가드)
        var lines = VbaHeader.SynthesizeHeader(VbaComponentKind.Document, "현재_통합_문서");
        var r = VbaHeader.Split(lines, "현재_통합_문서.cls");
        Assert.Equal(9, r.Header.Count);
        Assert.Empty(r.Body);
    }

    [Fact]
    public void Split_Form_NestedBeginEndBlocks()
    {
        var lines = new[]
        {
            "VERSION 5.00",
            "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilter01 ",
            "   Caption         =   \"필터\"",
            "   Begin MSForms.CommandButton cmdOK ",
            "      Caption         =   \"End\"",
            "   End",
            "   OleObjectBlob   =   \"frmFilter01.frx\":0000",
            "End",
            "Attribute VB_Name = \"frmFilter01\"",
            "Attribute VB_GlobalNameSpace = False",
            "Option Explicit",
            "Private Sub cmdOK_Click()",
            "End Sub",
        };
        var r = VbaHeader.Split(lines, "frmFilter01.frm");
        Assert.Equal(10, r.Header.Count);
        Assert.Equal(3, r.Body.Count);
        Assert.Equal("Option Explicit", r.Body[0]);
    }

    [Fact]
    public void Split_AttributeAfterCode_StrictThrows()
    {
        var lines = new[]
        {
            "Attribute VB_Name = \"modFoo\"",
            "Public Function Value() As Long",
            "Attribute Value.VB_UserMemId = 0",
            "End Function",
        };
        var ex = Assert.Throws<CliError>(() => VbaHeader.Split(lines, "modFoo.bas"));
        Assert.Contains("Attribute", ex.Message);
    }

    [Fact]
    public void Split_AttributeAfterCode_NonStrictKeepsBody()
    {
        var lines = new[]
        {
            "Attribute VB_Name = \"modFoo\"",
            "Public Function Value() As Long",
            "Attribute Value.VB_UserMemId = 0",
            "End Function",
        };
        var r = VbaHeader.Split(lines, "modFoo.bas", strict: false);
        Assert.Single(r.Header);
        Assert.Equal(3, r.Body.Count);
    }

    [Fact]
    public void Split_CodeNamedBeginIsNotHeader()
    {
        var lines = new[]
        {
            "Attribute VB_Name = \"modFoo\"",
            "Begin = 5", // 코드: Begin이라는 변수 대입 (Begin 블록 아님)
        };
        var r = VbaHeader.Split(lines, "modFoo.bas");
        Assert.Single(r.Header);
        Assert.Single(r.Body);
    }

    [Fact]
    public void Synthesize_DocumentVariant()
    {
        var doc = VbaHeader.SynthesizeHeader(VbaComponentKind.Document, "Sheet1");
        Assert.Contains("Attribute VB_PredeclaredId = True", doc);
        Assert.Contains("Attribute VB_Exposed = True", doc);
        var cls = VbaHeader.SynthesizeHeader(VbaComponentKind.ClassModule, "clsX");
        Assert.Contains("Attribute VB_PredeclaredId = False", cls);
        Assert.Contains("Attribute VB_Exposed = False", cls);
    }

    [Fact]
    public void Synthesize_Bas_SingleLine()
    {
        var h = VbaHeader.SynthesizeHeader(VbaComponentKind.StdModule, "modErpParser");
        Assert.Equal(new[] { "Attribute VB_Name = \"modErpParser\"" }, h);
    }
}
