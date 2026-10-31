using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class CompileErrorLocationTests
{
    // 실측(2026-10-31): 미정의 Sub 호출은 토큰만, 구문 오류는 행 전체가 선택된다
    private static readonly CompileErrorLocation Semantic = new("PerfSynth", 15, 5, 15, 15, "    NoSuchProc y, 3");
    private static readonly CompileErrorLocation Syntax =
        new("PerfSynth", 7, 1, 7, 43, "Public Sub Generate(ByVal scale As Double)");

    [Fact]
    public void Token_PartialSelection_ReturnsSelectedText()
        => Assert.Equal("NoSuchProc", Semantic.Token);

    [Fact]
    public void Token_WholeLineOrIndentedWholeLine_IsNull()
    {
        Assert.Null(Syntax.Token);
        Assert.Null(new CompileErrorLocation("M", 3, 1, 3, 12, "    x = 1  ").Token); // 들여쓰기 포함 전체 선택
    }

    [Fact]
    public void Token_MultiLineOrEmptySelection_IsNull()
    {
        Assert.Null(new CompileErrorLocation("M", 3, 1, 4, 5, "Sub A()").Token);
        Assert.Null(new CompileErrorLocation("M", 3, 5, 3, 5, "Sub A()").Token);
        Assert.Null(new CompileErrorLocation("M", 3, 40, 3, 45, "Sub A()").Token);
    }

    private static readonly string[] Body =
    {
        "Option Explicit", "", "Public Sub Gen_Inspect()", "End Sub", "",
        "Public Sub Generate(ByVal scale As Double)", "End Sub",
    };

    [Fact]
    public void MapToFileLine_SamePosition_AddsHeaderCount()
    {
        var loc = new CompileErrorLocation("PerfSynth", 6, 1, 6, 43, "Public Sub Generate(ByVal scale As Double)");
        Assert.Equal(7, loc.MapToFileLine(1, Body));
    }

    [Fact]
    public void MapToFileLine_ShiftedByAutoOptionExplicit_FindsUniqueMatch()
    {
        // 새 모듈에 VBE가 Option Explicit을 자동 삽입하면 코드 행이 한 줄 밀린다
        var loc = new CompileErrorLocation("PerfSynth", 7, 1, 7, 43, "Public Sub Generate(ByVal scale As Double)");
        Assert.Equal(7, loc.MapToFileLine(1, Body));
    }

    [Fact]
    public void MapToFileLine_AmbiguousOrMissing_IsNull()
    {
        var dup = new CompileErrorLocation("PerfSynth", 9, 1, 9, 8, "End Sub");
        Assert.Null(dup.MapToFileLine(1, Body));
        var missing = new CompileErrorLocation("PerfSynth", 2, 1, 2, 9, "x = 1");
        Assert.Null(missing.MapToFileLine(1, Body));
    }

    [Fact]
    public void Describe_UsesFileLineWhenMapped_ElseCodeLine()
    {
        Assert.Equal(@"src\PerfSynth.bas:16 'NoSuchProc' in: NoSuchProc y, 3", Semantic.Describe(@"src\PerfSynth.bas", 16));
        Assert.Equal("src\\PerfSynth.bas (code line 15) 'NoSuchProc' in: NoSuchProc y, 3", Semantic.Describe(@"src\PerfSynth.bas"));
        Assert.Equal("PerfSynth (code line 7): Public Sub Generate(ByVal scale As Double)", Syntax.Describe());
    }
}
