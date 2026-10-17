using System.Text.RegularExpressions;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>
/// install.ps1의 실행 경로 불변식. 실행 자체는 네트워크/레지스트리가 필요해 xUnit 밖이지만,
/// 깨지면 설치 첫 줄에서 죽는 텍스트 성질은 여기서 고정한다.
/// </summary>
public class InstallScriptTests
{
    private static byte[] ReadScript()
    {
        var root = RepoLocator.FindRepoRoot(AppContext.BaseDirectory)
            ?? throw new InvalidOperationException("레포 루트(VbaLoop.sln)를 찾지 못함");
        return File.ReadAllBytes(Path.Combine(root, "install.ps1"));
    }

    [Fact]
    public void Script_IsAsciiWithoutBom()
    {
        // BOM이 있으면 `irm | iex`가 U+FEFF를 첫 토큰에 붙여 "'<#' is not recognized"로 죽는다
        // (PS 5.1 실측). BOM이 없으면 PS 5.1은 로컬 .ps1을 ANSI 코드페이지로 읽으므로, 어느
        // 로캘에서도 같은 텍스트가 되려면 ASCII만 남는다.
        var bytes = ReadScript();
        Assert.False(bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF,
            "install.ps1에 UTF-8 BOM이 있음");
        int bad = Array.FindIndex(bytes, b => b > 127);
        Assert.True(bad < 0, $"install.ps1 오프셋 {bad}에 비ASCII 바이트");
    }

    [Fact]
    public void Script_NeverCallsExit()
    {
        // `irm | iex`는 호출자 세션에서 돈다 — exit는 설치 실패/완료 시 사용자의 셸 창을 닫는다
        var text = System.Text.Encoding.ASCII.GetString(ReadScript());
        var hits = Regex.Matches(text, @"^\s*exit\b", RegexOptions.Multiline | RegexOptions.IgnoreCase);
        Assert.Empty(hits);
    }
}
