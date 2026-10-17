using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class SampleProjectTests
{
    [Fact]
    public void FileName_MatchesModuleName()
    {
        // push는 파일명과 VB_Name 정합을 요구한다 — 샘플이 그 규약을 스스로 어기면 안 된다
        Assert.Equal(SampleProject.ModuleName + ".bas", SampleProject.FileName);
    }

    [Fact]
    public void ModuleLines_StartWithVbNameHeaderMatchingFileName()
        => Assert.Equal($"Attribute VB_Name = \"{SampleProject.ModuleName}\"",
                        SampleProject.ModuleLines()[0]);

    [Fact]
    public void ModuleLines_SplitStrict_HeaderOnlyHoldsAttribute()
    {
        // strict=true는 push 경로와 같은 검증 — 코드 시작 후 Attribute가 있으면 CliError.
        // 샘플이 프로시저 수준 Attribute를 달고 있으면 첫 push부터 깨진다.
        var split = VbaHeader.Split(SampleProject.ModuleLines().ToArray(), SampleProject.FileName);
        Assert.Contains(split.Header, l => l.Contains("VB_Name"));
        Assert.DoesNotContain(split.Body, l => l.TrimStart().StartsWith("Attribute", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ModuleLines_ExposeExactlyOneCollectableTest()
    {
        // 마법사의 성공 조건: setup 직후 vbaloop test가 green. 수집 대상이 0이면 초록불이 공허하다
        var procs = TestEnumerator.FindTestSubs(SampleProject.ModuleLines());
        Assert.Equal(new[] { "Test_Greet" }, procs);
    }

    [Fact]
    public void ModuleLines_UseAssertEq_NotDebugAssert()
    {
        var code = SampleProject.ModuleLines()
            .Where(l => !l.TrimStart().StartsWith('\''))   // 주석은 Debug.Assert 금지를 "설명"한다
            .ToList();
        Assert.Contains(code, l => l.Contains("AssertEq"));
        // 헤드리스에서 아무 신호도 내지 않으므로 호출은 금지 — 설명하는 주석과 구분해서 본다
        Assert.DoesNotContain(code, l => l.Contains("Debug.Assert"));
    }

    [Fact]
    public void ModuleLines_DeclareOptionExplicit()
        => Assert.Contains("Option Explicit", SampleProject.ModuleLines());

    [Fact]
    public void ModuleLines_AreAsciiOnly()
    {
        // push는 시스템 ACP 밖 문자를 save 전에 거부한다 - 샘플이 거기 걸리면 마법사가 자멸한다.
        // CP949/CP932/CP936/CP1252 어디서든 안전하도록 주석까지 포함해 전부 ASCII로 유지한다
        // (주입 자산 전체 검증은 VbaAssetAsciiTests).
        foreach (var line in SampleProject.ModuleLines())
            Assert.All(line, c => Assert.True(c < 128, $"비ASCII 문자 '{c}' in: {line}"));
    }
}
