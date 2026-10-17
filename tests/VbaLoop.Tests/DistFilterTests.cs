using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class DistFilterTests
{
    private static DistModule Mod(string name, string ext, params string[] body)
        => new(name, ext, new[] { $"Attribute VB_Name = \"{name}\"" }, body);

    private static DistModule ModWithHeader(string name, string[] header, params string[] body)
        => new(name, ".bas", header, body);

    [Fact]
    public void Glob_ExcludesTestPrefix()
    {
        var mods = new[] { Mod("modEngine", ".bas"), Mod("Test_Foo", ".bas"), Mod("SmokeBar", ".bas") };
        var r = DistFilter.Classify(mods, new[] { "Test_*", "Smoke*" });
        Assert.Equal(new[] { "modEngine" }, r.Included.Select(m => m.Name));
        Assert.Equal(new[] { "Test_Foo", "SmokeBar" }, r.ExcludedNames);
    }

    [Fact]
    public void HeaderMarker_Excludes()
    {
        var header = new[]
        {
            "Attribute VB_Name = \"modProbe\"",
            "' @dist exclude",
            "Option Explicit",
        };
        var mods = new[] { ModWithHeader("modProbe", header, "Sub P()", "End Sub"), Mod("modKeep", ".bas") };
        var r = DistFilter.Classify(mods, Array.Empty<string>());
        Assert.Equal(new[] { "modProbe" }, r.ExcludedNames);
        Assert.Equal(new[] { "modKeep" }, r.Included.Select(m => m.Name));
    }

    [Fact]
    public void HeaderMarker_CaseInsensitive_AndMultipleApostrophes()
    {
        var header = new[] { "Attribute VB_Name = \"m\"", "''  @DIST EXCLUDE" };
        var r = DistFilter.Classify(new[] { ModWithHeader("m", header) }, Array.Empty<string>());
        Assert.Equal(new[] { "m" }, r.ExcludedNames);
    }

    [Fact]
    public void MarkerInsideStringLiteral_NotExcluded()
    {
        // 인라인/문자열 안의 텍스트는 마커가 아니다 — 온전한 주석 라인만 인정
        var r = DistFilter.Classify(new[] { Mod("m", ".bas", "s = \"' @dist exclude\"") }, Array.Empty<string>());
        Assert.Empty(r.ExcludedNames);
        Assert.Single(r.Included);
    }

    [Fact]
    public void InlineComment_NotAMarker()
    {
        var r = DistFilter.Classify(new[] { Mod("m", ".bas", "x = 1 ' @dist exclude") }, Array.Empty<string>());
        Assert.Empty(r.ExcludedNames);
    }

    [Fact]
    public void BodyMarker_NotExcluded_HeaderOnly()
    {
        // 마커는 헤더 한정(§9.1). 본문의 온전한 주석 마커는 제외하지 않는다 —
        // 문서 인용·주석처리된 마커 오탐으로 인한 조용한 반쪽 배포 방지.
        var header = new[] { "Attribute VB_Name = \"m\"", "Option Explicit" };
        var mods = new[] { ModWithHeader("m", header, "Sub P()", "' @dist exclude", "End Sub") };
        var r = DistFilter.Classify(mods, Array.Empty<string>());
        Assert.Empty(r.ExcludedNames);
        Assert.Single(r.Included);
    }
}
