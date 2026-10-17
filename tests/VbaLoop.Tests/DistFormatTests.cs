using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class DistFormatTests
{
    [Theory]
    [InlineData("out.xlsm")]
    [InlineData("OUT.XLSM")]
    [InlineData(@"C:\a\b\dist\Book.xlsm")]
    public void FromOutPath_Xlsm_MacroWorkbook(string path)
        => Assert.Equal(DistOutputKind.MacroWorkbook, DistFormat.FromOutPath(path));

    [Theory]
    [InlineData("out.xlam")]
    [InlineData("OUT.XLAM")]
    [InlineData(@"C:\a\dist\Addin.xlam")]
    public void FromOutPath_Xlam_AddIn(string path)
        => Assert.Equal(DistOutputKind.AddIn, DistFormat.FromOutPath(path));

    [Theory]
    [InlineData("out.xlsx")]  // 매크로 없음
    [InlineData("out.xls")]
    [InlineData("out.xla")]   // 구 바이너리 애드인 미지원
    [InlineData("out.txt")]
    [InlineData("out")]        // 무확장
    public void FromOutPath_Unsupported_Throws(string path)
        => Assert.Throws<CliError>(() => DistFormat.FromOutPath(path));

    [Fact]
    public void DefaultOut_UnderDistWithXlsmExtension()
    {
        var outPath = DistFormat.DefaultOut(@"C:\proj", @"C:\proj\My Book(v2).xlsm");
        Assert.Equal(System.IO.Path.Combine(@"C:\proj", "dist", "My Book(v2).xlsm"), outPath);
        // 기본 출력은 매크로 워크북으로 판정되어야 한다(왕복 자기일관성)
        Assert.Equal(DistOutputKind.MacroWorkbook, DistFormat.FromOutPath(outPath));
    }
}
