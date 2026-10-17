using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class VbaLineTests
{
    [Fact]
    public void SplitCodeComment_InlineComment_Split()
    {
        var (code, comment) = VbaLine.SplitCodeComment("x = 1 ' 설명");
        Assert.Equal("x = 1 ", code);
        Assert.Equal("' 설명", comment);
    }

    [Fact]
    public void SplitCodeComment_ApostropheInsideString_NotComment()
    {
        // §9.5 핵심: 문자열 리터럴 안의 '는 주석이 아니다
        var (code, comment) = VbaLine.SplitCodeComment("x = \"it's a test\"");
        Assert.Equal("x = \"it's a test\"", code);
        Assert.Null(comment);
    }

    [Fact]
    public void SplitCodeComment_EscapedQuoteThenComment()
    {
        var (code, comment) = VbaLine.SplitCodeComment("s = \"a\"\"b\" ' 꼬리");
        Assert.Equal("s = \"a\"\"b\" ", code);
        Assert.Equal("' 꼬리", comment);
    }

    [Fact]
    public void SplitCodeComment_FullCommentLine()
    {
        var (code, comment) = VbaLine.SplitCodeComment("   ' 온전한 주석");
        Assert.Equal("   ", code);
        Assert.Equal("' 온전한 주석", comment);
    }

    [Fact]
    public void Identifiers_SkipStringContents()
    {
        var toks = VbaLine.Identifiers("Call Foo(\"Bar Baz\", Qux)");
        Assert.Equal(new[] { "Call", "Foo", "Qux" }, toks.Select(t => t.Text));
    }

    [Fact]
    public void Identifiers_TrackDotAdjacency()
    {
        var code = "LogShim.LogD x";
        var toks = VbaLine.Identifiers(code);
        var logShim = toks[0];
        var logD = toks[1];
        Assert.Equal("LogShim", logShim.Text);
        Assert.Equal('.', code[logShim.End]); // followedByDot
        Assert.Equal('.', code[logD.Start - 1]); // precededByDot
    }
}
