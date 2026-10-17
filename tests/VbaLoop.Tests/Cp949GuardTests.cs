using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class Cp949GuardTests
{
    static Cp949GuardTests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Fact]
    public void Korean_Passes()
    {
        var v = Cp949Guard.Check(new[] { "' 한글 주석은 무손실 왕복 (실측: U+AC00 → B0 A1)" });
        Assert.Empty(v);
    }

    [Fact]
    public void EmDash_Caught_WithPosition()
    {
        // 실측: U+2014 → 0x3F 확정 손실. 가드가 save 전에 탐지해야 한다.
        var v = Cp949Guard.Check(new[] { "' 정상 라인", "' em dash — 여기" });
        var hit = Assert.Single(v);
        Assert.Equal(2, hit.Line);
        Assert.Equal(11, hit.Column);
        Assert.Equal("U+2014", hit.CodePoint);
    }

    [Fact]
    public void Emoji_SurrogatePair_Caught()
    {
        var v = Cp949Guard.Check(new[] { "' 이모지 \U0001F600 테스트" });
        var hit = Assert.Single(v);
        Assert.Equal("U+1F600", hit.CodePoint);
    }

    [Fact]
    public void MultipleViolations_AllReported()
    {
        var v = Cp949Guard.Check(new[] { "—a—" });
        Assert.Equal(2, v.Count);
    }
}
