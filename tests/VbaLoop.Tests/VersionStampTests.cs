using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class VersionStampTests
{
    [Fact]
    public void BuildInfoModule_HasHeaderAndConstants()
    {
        var text = VersionStamp.BuildInfoModule("1.2.0", "8acb552-dirty", "2026-07-09 14:00:00");
        Assert.Contains("Attribute VB_Name = \"modBuildInfo\"", text);
        Assert.Contains("Public Const BUILD_VERSION As String = \"1.2.0\"", text);
        Assert.Contains("Public Const BUILD_REV As String = \"8acb552-dirty\"", text);
        Assert.Contains("Public Const BUILT_AT As String = \"2026-07-09 14:00:00\"", text);
        Assert.EndsWith("\r\n", text); // 정규 EOF 개행
    }

    [Fact]
    public void BuildInfoModule_IsSplittableBackToConstants()
    {
        // 왕복: 합성 → VbaHeader.Split → 본문에 상수 3개
        var text = VersionStamp.BuildInfoModule("9.9.9", "abc", "t");
        var lines = CanonicalText.SplitLines(text);
        var split = VbaHeader.Split(lines, "modBuildInfo.bas");
        Assert.Single(split.Header);
        Assert.Contains(split.Body, l => l.Contains("BUILD_VERSION"));
        Assert.Equal(3, split.Body.Count(l => l.TrimStart().StartsWith("Public Const")));
    }

    [Fact]
    public void Escape_QuotesDoubled()
    {
        var text = VersionStamp.BuildInfoModule("1\"0", "r", "t");
        Assert.Contains("\"1\"\"0\"", text);
    }

    [Fact]
    public void Body_MatchesModuleMinusHeader()
    {
        var body = VersionStamp.BuildInfoBody("1.0", "r", "t");
        Assert.Equal("Option Explicit", body[1]);
        Assert.Equal(6, body.Count);
    }
}
