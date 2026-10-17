using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class AccessVbomTests
{
    [Theory]
    [InlineData("Excel.Application.16", "16.0")]
    [InlineData("Excel.Application.15", "15.0")]
    [InlineData("Excel.Application.14", "14.0")]
    public void VersionFromCurVer_Standard(string curVer, string expected)
        => Assert.Equal(expected, AccessVbom.VersionFromCurVer(curVer));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Excel.Application")]     // 마이너 없음
    [InlineData("Excel.Application.")]    // 빈 마이너
    [InlineData("Excel.Application.x")]   // 비숫자
    public void VersionFromCurVer_Unparseable_ReturnsNull(string? curVer)
        => Assert.Null(AccessVbom.VersionFromCurVer(curVer));
}
