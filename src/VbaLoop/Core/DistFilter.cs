namespace VbaLoop.Core;

/// <summary>
/// dist 배포 축의 최소 표현: 헤더/본문을 분리해 이고 있는 모듈(설계 §9.1).
/// vendored(동기화 축)와 직교 — 여기엔 벤더링 개념이 없다.
/// </summary>
public sealed record DistModule(
    string Name, string Ext, IReadOnlyList<string> Header, IReadOnlyList<string> Body);

/// <summary>
/// 모듈 제외 판정(설계 §9.1): 배포 축은 (1) 헤더의 `' @dist exclude` 마커(1차) +
/// (2) toml `[dist] exclude` 글롭(보강)만 따른다. 입도는 모듈 단위 — 프로시저 제거는 금지.
/// vendored 여부는 참조하지 않는다(LogShim은 벤더링이지만 런타임 의존이라 배포 포함).
/// </summary>
public static class DistFilter
{
    public sealed record Result(
        IReadOnlyList<DistModule> Included,
        IReadOnlyList<DistModule> Excluded,
        IReadOnlyList<string> ExcludedNames);

    /// <summary>
    /// 각 모듈을 (포함/제외)로 분할. 제외 사유는 헤더 마커 또는 글롭 매칭.
    /// </summary>
    public static Result Classify(IEnumerable<DistModule> modules, IEnumerable<string> excludeGlobs)
    {
        var globs = excludeGlobs?.ToList() ?? new List<string>();
        var included = new List<DistModule>();
        var excluded = new List<DistModule>();

        foreach (var m in modules)
        {
            // 마커는 헤더 한정(설계 §9.1: `Attribute VB_Name` 직후 헤더 영역). 본문 스캔은
            // 문서 인용·주석처리된 마커를 오탐해 모듈을 조용히 반쪽 배포시킬 수 있어 배제한다.
            if (HasExcludeMarker(m.Header) || MatchesAnyGlob(m.Name, globs))
                excluded.Add(m);
            else
                included.Add(m);
        }

        return new Result(included, excluded,
            excluded.Select(m => m.Name).ToList());
    }

    /// <summary>
    /// 독립 주석 라인 `' @dist exclude`(선행 `'` 다중/공백 허용, 대소문자 무시)를 찾음.
    /// 인라인이 아니라 온전한 주석 라인만 마커로 인정한다(오탐 방지).
    /// </summary>
    private static bool HasExcludeMarker(IReadOnlyList<string> lines)
    {
        foreach (var line in lines)
        {
            var (code, comment) = VbaLine.SplitCodeComment(line);
            if (comment == null || code.Trim().Length != 0) continue; // 온전한 주석 라인만
            string body = comment.TrimStart('\'').Trim();
            if (body.Equals("@dist exclude", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static bool MatchesAnyGlob(string name, IReadOnlyList<string> globs)
    {
        foreach (var g in globs)
            if (Glob.IsMatch(name, g)) return true;
        return false;
    }
}

/// <summary>`*`(0+ 임의)와 `?`(1 임의)만 지원하는 최소 글롭. 대소문자 무시.</summary>
public static class Glob
{
    public static bool IsMatch(string input, string pattern)
    {
        // 백트래킹 없는 선형 매칭(설계상 글롭은 짧다). 별표는 최대 하나로도 충분하나 일반화.
        int s = 0, p = 0, star = -1, mark = 0;
        while (s < input.Length)
        {
            if (p < pattern.Length && (pattern[p] == '?' ||
                char.ToUpperInvariant(pattern[p]) == char.ToUpperInvariant(input[s])))
            {
                s++; p++;
            }
            else if (p < pattern.Length && pattern[p] == '*')
            {
                star = p++; mark = s;
            }
            else if (star != -1)
            {
                p = star + 1; s = ++mark;
            }
            else
            {
                return false;
            }
        }
        while (p < pattern.Length && pattern[p] == '*') p++;
        return p == pattern.Length;
    }
}
