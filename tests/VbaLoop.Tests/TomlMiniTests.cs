using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class TomlMiniTests
{
    [Fact]
    public void HappyPath_SectionsAndValues()
    {
        var doc = TomlMini.Parse("""
            # 주석 라인
            [project]
            name = "SharedDb"

            [excel]
            workbook = "SharedDb.xlsm"   # 인라인 주석
            visible = true

            [src]
            root = "src"
            """, "test.toml");
        Assert.Equal("SharedDb", doc["project"]["name"]);
        Assert.Equal("SharedDb.xlsm", doc["excel"]["workbook"]);
        Assert.Equal(true, doc["excel"]["visible"]);
        Assert.Equal("src", doc["src"]["root"]);
    }

    [Fact]
    public void WindowsPath_BackslashLiteral()
    {
        var doc = TomlMini.Parse("[excel]\nworkbook = \"sub\\dir\\Book.xlsm\"", "t");
        Assert.Equal(@"sub\dir\Book.xlsm", doc["excel"]["workbook"]);
    }

    [Theory]
    [InlineData("[excel]\ncount = 3.5")]         // 실수 미지원
    [InlineData("[backup]\nkeep = 1_000")]       // 밑줄 구분자 미지원
    [InlineData("[backup]\nkeep = 0x10")]        // 16진 미지원
    [InlineData("[backup]\nkeep = 99999999999")] // int 범위 초과
    [InlineData("[backup]\nkeep = -")]           // 부호만
    [InlineData("key = \"before section\"")]     // 섹션 이전 키
    [InlineData("[excel]\nworkbook = \"x")]      // 닫는 따옴표 없음
    [InlineData("[excel\nworkbook = \"x\"")]     // 섹션 헤더 미완
    [InlineData("[excel]\nworkbook \"x\"")]      // = 없음
    [InlineData("[dist]\nexclude = [\"a\" \"b\"]")] // 구분자 없음
    [InlineData("[dist]\nexclude = [\"a\"")]     // 닫는 대괄호 없음
    public void UnsupportedSyntax_FailsLoudly(string toml)
    {
        // 조용한 오파싱 금지 (설계 §5)
        Assert.Throws<CliError>(() => TomlMini.Parse(toml, "t"));
    }

    [Theory]
    [InlineData("keep = 10", 10)]
    [InlineData("keep = 0   # 비활성", 0)]
    [InlineData("keep = -3", -3)]
    [InlineData("keep = +7", 7)]
    public void Integer_Parsed(string line, int expected)
    {
        var doc = TomlMini.Parse("[backup]\n" + line, "t");
        Assert.Equal(expected, Assert.IsType<int>(doc["backup"]["keep"]));
    }

    [Fact]
    public void StringArray_Parsed()
    {
        var doc = TomlMini.Parse("[dist]\nexclude = [\"Test_*\", \"Smoke*\"]   # 보강 글롭", "t");
        var list = Assert.IsType<List<string>>(doc["dist"]["exclude"]);
        Assert.Equal(new[] { "Test_*", "Smoke*" }, list);
    }

    [Fact]
    public void EmptyArray_Parsed()
    {
        var doc = TomlMini.Parse("[dist]\nexclude = []", "t");
        Assert.Empty(Assert.IsType<List<string>>(doc["dist"]["exclude"]));
    }

    [Fact]
    public void DuplicateKey_Fails()
    {
        Assert.Throws<CliError>(() => TomlMini.Parse("[a]\nx = \"1\"\nx = \"2\"", "t"));
    }
}
