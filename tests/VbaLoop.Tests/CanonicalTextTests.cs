using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class CanonicalTextTests
{
    [Fact]
    public void SplitLines_MixedEndings_Tolerated()
    {
        var lines = CanonicalText.SplitLines("a\r\nb\nc\rd");
        Assert.Equal(new[] { "a", "b", "c", "d" }, lines);
    }

    [Fact]
    public void SplitLines_TrailingNewline_RemovedOnce()
    {
        Assert.Equal(new[] { "a", "b" }, CanonicalText.SplitLines("a\r\nb\r\n"));
        Assert.Equal(new[] { "a", "b", "" }, CanonicalText.SplitLines("a\r\nb\r\n\r\n"));
    }

    [Fact]
    public void ToCanonical_CrlfWithSingleTrailingNewline()
    {
        Assert.Equal("a\r\nb\r\n", CanonicalText.ToCanonical(new[] { "a", "b" }));
    }

    [Fact]
    public void ReadFileStripBom_BomRemoved()
    {
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        try
        {
            File.WriteAllText(path, "Attribute VB_Name = \"x\"\r\n", new UTF8Encoding(true)); // BOM 포함
            var text = CanonicalText.ReadFileStripBom(path);
            Assert.False(text.StartsWith('﻿'));
            Assert.StartsWith("Attribute", text);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TrimTrailingEmpty_RemovesOnlyTrailing()
    {
        // 폼 왕복 수렴의 전제: Import/Export가 끝쪽에 누적시키는 빈 줄 제거 (실측 596->592 수렴)
        Assert.Equal(new[] { "a", "", "b" },
            CanonicalText.TrimTrailingEmpty(new[] { "a", "", "b", "", "  ", "" }));
        Assert.Empty(CanonicalText.TrimTrailingEmpty(new[] { "", "" }));
        Assert.Equal(new[] { "a" }, CanonicalText.TrimTrailingEmpty(new[] { "a" }));
    }

    [Fact]
    public void Roundtrip_Idempotent()
    {
        // 정규화 후 재정규화는 항등 (§7-2 왕복 무손실의 파일측 전제)
        var original = "Option Explicit\r\n\r\nSub A()\r\nEnd Sub\r\n";
        var lines = CanonicalText.SplitLines(original);
        Assert.Equal(original, CanonicalText.ToCanonical(lines));
    }

    [Fact]
    public void PreserveCasing_CasingOnlyDifference_KeepsSrcSpelling()
    {
        // VBE가 .Path를 .path로 재케이싱해도 src 철자를 유지
        var wb = new[] { "    If LenB(ThisWorkbook.path) = 0 Then Exit Function" };
        var src = new[] { "    If LenB(ThisWorkbook.Path) = 0 Then Exit Function" };
        Assert.Equal(src, CanonicalText.PreserveCasing(wb, src));
    }

    [Fact]
    public void PreserveCasing_RealContentChange_AdoptsWorkbook()
    {
        // 대소문자 무시로도 다른 라인(진짜 내용 변경)은 워크북을 채택
        var wb = new[] { "n = n + 1" };
        var src = new[] { "n = n + 2" };
        Assert.Equal(wb, CanonicalText.PreserveCasing(wb, src));
    }

    [Fact]
    public void PreserveCasing_MixedLines_RescuesOnlyCasingLines()
    {
        var wb = new[] { "x.name = 1", "n = n + 1", "y.path" };
        var src = new[] { "x.Name = 1", "n = n + 2", "y.Path" };
        Assert.Equal(new[] { "x.Name = 1", "n = n + 1", "y.Path" }, CanonicalText.PreserveCasing(wb, src));
    }

    [Fact]
    public void PreserveCasing_WorkbookLongerThanSrc_ExtraLinesUnchanged()
    {
        // src에 대응 라인이 없는 워크북 뒷줄(추가된 코드)은 그대로 워크북 채택
        var wb = new[] { "a.name", "newLine" };
        var src = new[] { "a.Name" };
        Assert.Equal(new[] { "a.Name", "newLine" }, CanonicalText.PreserveCasing(wb, src));
    }

    [Fact]
    public void PreserveCasing_EmptyEitherSide_ReturnsWorkbook()
    {
        Assert.Empty(CanonicalText.PreserveCasing(System.Array.Empty<string>(), new[] { "a" }));
        Assert.Equal(new[] { "a" }, CanonicalText.PreserveCasing(new[] { "a" }, System.Array.Empty<string>()));
    }

    [Fact]
    public void StripMemberAttributes_RemovesDottedAttributeLines()
    {
        // VBE가 WithEvents 멤버에 삽입하는 .VB_VarHelpID 라인 제거
        var input = new[]
        {
            "Private WithEvents m_btn As MSForms.CommandButton",
            "Attribute m_btn.VB_VarHelpID = -1",
            "Private WithEvents m_cbo As MSForms.ComboBox",
            "Attribute m_cbo.VB_VarHelpID = -1",
        };
        Assert.Equal(new[]
        {
            "Private WithEvents m_btn As MSForms.CommandButton",
            "Private WithEvents m_cbo As MSForms.ComboBox",
        }, CanonicalText.StripMemberAttributes(input));
    }

    [Fact]
    public void StripMemberAttributes_PreservesModuleHeaderAndCode()
    {
        // 모듈 헤더 Attribute(점 없음)와 일반 코드 라인은 보존
        var input = new[]
        {
            "Attribute VB_Name = \"frmX\"",
            "Attribute VB_Exposed = False",
            "    total = 1 + m_btn.Width",   // 코드 중 '.'은 무관 (Attribute 접두 아님)
            "    Attribute = 3",             // 'Attribute ' 접두지만 점 없음 → 보존
        };
        Assert.Equal(input, CanonicalText.StripMemberAttributes(input));
    }

    [Fact]
    public void StripMemberAttributes_IndentedMemberAttribute_AlsoStripped()
    {
        Assert.Equal(new[] { "Dim x" },
            CanonicalText.StripMemberAttributes(new[] { "Dim x", "   Attribute foo.VB_ProcData = \"x\"" }));
    }
}
