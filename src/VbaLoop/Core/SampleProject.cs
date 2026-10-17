namespace VbaLoop.Core;

/// <summary>
/// setup이 빈 폴더에 넣는 최소 샘플(설계: 마법사는 빈 스캐폴드가 아니라 "이미 도는 루프"를 준다).
/// 목적은 두 가지 — 사람은 <c>vbaloop test</c> green을 즉시 보고, 에이전트는 Test_ 규약과
/// LogShim 단언 사용법을 문서가 아니라 동작하는 코드로 본다. 순수 텍스트 생성이라 xUnit 검증.
/// </summary>
public static class SampleProject
{
    public const string ModuleName = "modMain";

    /// <summary>파일명은 VB_Name과 일치해야 한다(push의 모듈명 정합 검사).</summary>
    public static string FileName => ModuleName + ".bas";

    /// <summary>
    /// 헤더 포함 정규 본문. 신규 사용자가 손으로 쓰기 어려운 <c>Attribute VB_Name</c> 헤더와
    /// src 규약(BOM 없는 UTF-8 + CRLF + 끝 개행 1개)을 처음부터 맞춰 준다 —
    /// 기록은 <see cref="CanonicalText.WriteCanonical"/>이 보장한다.
    /// </summary>
    public static IReadOnlyList<string> ModuleLines() => new[]
    {
        $"Attribute VB_Name = \"{ModuleName}\"",
        "' @desc Sample created by vbaloop setup - edit or delete freely",
        "Option Explicit",
        "",
        "' @desc Takes a target name and returns a greeting string",
        "Public Function Greet(ByVal target As String) As String",
        "    Greet = \"Hello, \" & target & \"!\"",
        "End Function",
        "",
        "' vbaloop test collects only Public Sub Test_*() with no arguments.",
        "' Assert with LogShim.AssertEq - Debug.Assert gives no signal when headless, so do not use it.",
        "Public Sub Test_Greet()",
        "    AssertEq \"greet\", \"Hello, vbaloop!\", Greet(\"vbaloop\")",
        "End Sub",
    };
}
