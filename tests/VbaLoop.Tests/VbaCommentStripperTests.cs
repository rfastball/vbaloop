using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class VbaCommentStripperTests
{
    [Fact]
    public void Strip_RemovesFullCommentLines()
    {
        var body = new[] { "' 온전 주석", "Dim x As Long", "   ' 들여쓴 주석", "x = 1" };
        var r = VbaCommentStripper.Strip(body);
        Assert.Equal(new[] { "Dim x As Long", "x = 1" }, r);
    }

    [Fact]
    public void Strip_RemovesInlineComments_TrimsTrailing()
    {
        var r = VbaCommentStripper.Strip(new[] { "x = 1    ' 설명" });
        Assert.Equal(new[] { "x = 1" }, r);
    }

    [Fact]
    public void Strip_PreservesApostropheInsideString()
    {
        // §9.5 핵심: 문자열 안의 '는 보존
        var r = VbaCommentStripper.Strip(new[] { "s = \"it's a test\"" });
        Assert.Equal(new[] { "s = \"it's a test\"" }, r);
    }

    [Fact]
    public void Strip_RemovesRemComments()
    {
        var r = VbaCommentStripper.Strip(new[] { "REM 옛 주석", "Rem 또", "x = 1" });
        Assert.Equal(new[] { "x = 1" }, r);
    }

    [Fact]
    public void Strip_KeepsStringWithHashAndQuotes()
    {
        var r = VbaCommentStripper.Strip(new[] { "s = \"a\"\"b\"    ' 꼬리" });
        Assert.Equal(new[] { "s = \"a\"\"b\"" }, r);
    }
}
