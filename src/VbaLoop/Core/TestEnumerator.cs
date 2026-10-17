using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>
/// 테스트 열거 규칙(설계 §4 test): 표준 모듈의 인자 0개 Sub 중 이름이 ^Test_ 매칭
/// (접근수식어 없는 암묵 Public 포함, Private/Friend 제외). 모듈 텍스트를 직접 파싱해
/// Excel 없이 검증 가능하게 한다.
/// </summary>
public static class TestEnumerator
{
    private static readonly Regex SubDecl = new(
        @"^\s*(?:(?<access>Public|Private|Friend)\s+)?(?:Static\s+)?Sub\s+(?<name>\w+)\s*\(\s*\)\s*(?:'.*)?$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static List<string> FindTestSubs(IReadOnlyList<string> lines)
    {
        var result = new List<string>();
        for (int i = 0; i < lines.Count; i++)
        {
            string line = lines[i];
            // 라인 계속 문자(_) 결합 — 선언부가 여러 줄인 드문 케이스
            while (line.TrimEnd().EndsWith(" _", StringComparison.Ordinal) && i + 1 < lines.Count)
                line = line.TrimEnd()[..^1] + lines[++i].TrimStart();

            var m = SubDecl.Match(line);
            if (!m.Success) continue;
            string access = m.Groups["access"].Value;
            if (access.Equals("Private", StringComparison.OrdinalIgnoreCase) ||
                access.Equals("Friend", StringComparison.OrdinalIgnoreCase))
                continue;
            string name = m.Groups["name"].Value;
            if (name.StartsWith("Test_", StringComparison.Ordinal))
                result.Add(name);
        }
        return result;
    }
}
