namespace VbaLoop.Core;

/// <summary>
/// 버전 스탬핑(설계 §9.4): dist가 아티팩트에만 합성하는 <c>modBuildInfo</c> 모듈.
/// src엔 없고(벤더링 경로 아님) 아티팩트 빌드 산출물로만 존재한다. About 박스·로그에서
/// 코드가 상수를 읽고, 외부 감사는 CustomDocumentProperties(명령 계층)로 본다.
/// </summary>
public static class VersionStamp
{
    public const string ModuleName = "modBuildInfo";

    /// <summary>합성 모듈의 정규 텍스트(헤더 + Option Explicit + 3개 Public 상수).</summary>
    public static string BuildInfoModule(string version, string rev, string builtAt)
    {
        var lines = new List<string>
        {
            $"Attribute VB_Name = \"{ModuleName}\"",
            "' @desc Provides dist build stamp constants (synthesized; not in src)",
            "Option Explicit",
            "",
            $"Public Const BUILD_VERSION As String = \"{Escape(version)}\"",
            $"Public Const BUILD_REV As String = \"{Escape(rev)}\"",
            $"Public Const BUILT_AT As String = \"{Escape(builtAt)}\"",
        };
        return CanonicalText.ToCanonical(lines);
    }

    /// <summary>본문 라인(헤더 제외) — VBProject 주입 경로(AddFromString)용.</summary>
    public static IReadOnlyList<string> BuildInfoBody(string version, string rev, string builtAt) => new[]
    {
        "' @desc Provides dist build stamp constants (synthesized; not in src)",
        "Option Explicit",
        "",
        $"Public Const BUILD_VERSION As String = \"{Escape(version)}\"",
        $"Public Const BUILD_REV As String = \"{Escape(rev)}\"",
        $"Public Const BUILT_AT As String = \"{Escape(builtAt)}\"",
    };

    /// <summary>VBA 문자열 리터럴 안의 따옴표는 두 개로 이스케이프.</summary>
    private static string Escape(string s) => s.Replace("\"", "\"\"");
}
