using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class MacroNameTests
{
    [Theory]
    [InlineData("modEngine.RunAll", "modEngine", "RunAll")]
    [InlineData("  modA.Proc1  ", "modA", "Proc1")]
    [InlineData("현재_통합_문서.초기화", "현재_통합_문서", "초기화")] // 한글 식별자 허용
    public void Valid_QualifiedNames(string input, string module, string proc)
    {
        Assert.True(MacroName.TryParse(input, out var m, out var p, out var err));
        Assert.Equal(module, m);
        Assert.Equal(proc, p);
        Assert.Null(err);
    }

    [Theory]
    [InlineData("RunAll")]                    // 비한정 — 모듈 한정 위반
    [InlineData("a.b.c")]                     // 점 2개
    [InlineData("mod.")]                      // proc 없음
    [InlineData(".Proc")]                     // module 없음
    [InlineData("'Book.xlsm'!mod.Proc")]      // 워크북 한정은 자동 — 직접 지정 금지
    [InlineData("mod!Proc")]
    [InlineData("mod Proc.x")]                // 공백
    [InlineData("")]
    [InlineData(null)]
    public void Invalid_Rejected(string? input)
    {
        Assert.False(MacroName.TryParse(input, out _, out _, out var err));
        Assert.NotNull(err);
    }
}
