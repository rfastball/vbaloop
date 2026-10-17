using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class LogShimReleaseTests
{
    [Fact]
    public void FlipToRelease_TurnsFalseToTrue()
    {
        var lines = new[]
        {
            "Attribute VB_Name = \"LogShim\"",
            "Private Const IS_RELEASE As Boolean = False",
            "Public Sub LogD(level As String, msg As String)",
            "    If IS_RELEASE Then Exit Sub",
            "End Sub",
        };
        var r = LogShimRelease.FlipToRelease(lines);
        Assert.Equal("Private Const IS_RELEASE As Boolean = True", r[1]);
        // 호출부는 불변 — 컴파일 안전
        Assert.Equal("    If IS_RELEASE Then Exit Sub", r[3]);
    }

    [Fact]
    public void FlipToRelease_PreservesTrailingComment()
    {
        var lines = new[] { "Private Const IS_RELEASE As Boolean = False ' 배포 스위치" };
        var r = LogShimRelease.FlipToRelease(lines);
        Assert.Equal("Private Const IS_RELEASE As Boolean = True ' 배포 스위치", r[0]);
    }

    [Fact]
    public void FlipToRelease_MissingConst_Throws()
    {
        var lines = new[] { "Attribute VB_Name = \"LogShim\"", "Option Explicit" };
        var ex = Assert.Throws<CliError>(() => LogShimRelease.FlipToRelease(lines));
        Assert.Contains("IS_RELEASE", ex.Message);
    }

    [Fact]
    public void FlipToRelease_OnlyFirstOccurrence()
    {
        var lines = new[]
        {
            "Private Const IS_RELEASE As Boolean = False",
            "x = False", // 다른 False는 건드리지 않음
        };
        var r = LogShimRelease.FlipToRelease(lines);
        Assert.Contains("= True", r[0]);
        Assert.Equal("x = False", r[1]);
    }

    // MutateProject는 같은 루프에서 strip을 먼저, flip을 이미 스트립된 body에 적용한다
    // (DistCommand.MutateProject). 각각의 단위 테스트는 이 조합을 검증하지 않으므로,
    // 주석 제거가 IS_RELEASE 상수의 트레일링 주석을 지워도 flip이 여전히 상수를 찾아
    // 뒤집는지 고정한다(리팩터 회귀 방어, 감사 §5).
    [Fact]
    public void StripThenFlip_RemovesTrailingComment_AndStillFlips()
    {
        var body = new[]
        {
            "' @desc LogShim - 로깅/어서션 파실리티", // 온전한 주석 라인 → strip이 제거
            "Private Const IS_RELEASE As Boolean = False ' 배포 무해화 스위치(9.3)",
            "Public Sub LogD(level As String, msg As String)",
            "    If IS_RELEASE Then Exit Sub  ' 무해화 게이트",
            "End Sub",
        };
        var stripped = VbaCommentStripper.Strip(body);
        var flipped = LogShimRelease.FlipToRelease(stripped);

        Assert.Contains("Private Const IS_RELEASE As Boolean = True", flipped);
        Assert.DoesNotContain(flipped, l => l.Contains("무해화 스위치")); // 트레일링 주석 제거됨
        Assert.DoesNotContain(flipped, l => l.TrimStart().StartsWith("' @desc")); // 온전한 주석 라인 제거됨
        Assert.Contains("    If IS_RELEASE Then Exit Sub", flipped); // 호출부/게이트는 보존(컴파일 안전)
    }
}
