using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>
/// push 무결성 체크(설계 §4 push 6단계): 컴포넌트 이름 집합 diff.
/// 기대하지 않은 새 이름(특히 &lt;기존명&gt;+숫자 패턴, 중복 모듈 시그니처) 출현이 실패 조건.
/// 의도한 Add/Import 이름은 기대 집합에 포함해 오탐을 방지.
/// </summary>
public static class NameSetDiff
{
    public sealed record Result(
        IReadOnlyList<string> Unexpected,
        IReadOnlyList<string> Missing,
        IReadOnlyList<string> DuplicateSignatures);

    public static Result Check(
        IEnumerable<string> before,
        IEnumerable<string> expectedAdded,
        IEnumerable<string> expectedRemoved,
        IEnumerable<string> after)
    {
        var cmp = StringComparer.OrdinalIgnoreCase;
        // (before − removed) ∪ added — 폼 왕복(Remove 후 같은 이름 Import)이 기대 집합에 남도록
        var expected = new HashSet<string>(before, cmp);
        expected.ExceptWith(expectedRemoved);
        expected.UnionWith(expectedAdded);

        var afterSet = new HashSet<string>(after, cmp);

        var unexpected = afterSet.Where(n => !expected.Contains(n)).OrderBy(n => n, cmp).ToList();
        var missing = expected.Where(n => !afterSet.Contains(n)).OrderBy(n => n, cmp).ToList();

        // <기존명>+숫자 패턴: 중복 모듈 시그니처 (예: modCache1)
        var dupSignatures = new List<string>();
        foreach (var name in unexpected)
        {
            var m = Regex.Match(name, @"^(.+?)(\d+)$");
            if (m.Success && expected.Contains(m.Groups[1].Value))
                dupSignatures.Add(name);
        }

        return new Result(unexpected, missing, dupSignatures);
    }
}
