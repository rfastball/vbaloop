namespace VbaLoop.Core;

/// <summary>
/// 컴파일 안전 검증(설계 §9.2): 제외 후 "잔존 모듈이 사라진 심볼을 참조 → 배포본이 안 열림"을
/// 막는다. 헤드리스 전체 강제 컴파일은 원천 불가이므로 <b>본문 토큰 참조 스캔이 권위</b>다
/// (`@depends`는 힌트일 뿐 신뢰 원천 아님 — modEngine 실측). 제외 모듈들의 {모듈명 ∪ Public
/// 심볼명} 집합을 만들고, 잔존 모듈 본문에서 그 이름 참조를 찾으면 1건이라도 dist 거부.
/// </summary>
public static class DistCompileScan
{
    public sealed record DanglingRef(string InModule, int Line, string Symbol, string Kind);

    private static readonly HashSet<string> ProcKeywords = new(StringComparer.OrdinalIgnoreCase)
        { "Sub", "Function", "Property" };
    private static readonly HashSet<string> PublicDataKeywords = new(StringComparer.OrdinalIgnoreCase)
        { "Const", "Enum", "Type", "Declare" };
    private static readonly HashSet<string> Modifiers = new(StringComparer.OrdinalIgnoreCase)
        { "Public", "Friend", "Global", "Static" };

    /// <summary>
    /// 잔존(included) 본문에서 제외(excluded) 모듈의 공개 표면 참조를 스캔.
    /// 반환이 비어 있어야 dist 안전. 한계: 도달 불가 모듈은 못 잡으므로 "검증 통과 =
    /// 완전 컴파일 보장 아님"(§9.2) — 카나리아는 명령 계층에서 보강.
    /// </summary>
    public static List<DanglingRef> Scan(
        IReadOnlyList<DistModule> included, IReadOnlyList<DistModule> excluded)
    {
        var moduleNames = new HashSet<string>(
            excluded.Select(m => m.Name), StringComparer.OrdinalIgnoreCase);
        var publicSymbols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in excluded)
            foreach (var sym in PublicSymbols(m.Body))
                publicSymbols.Add(sym);

        var findings = new List<DanglingRef>();
        foreach (var m in included)
        {
            for (int li = 0; li < m.Body.Count; li++)
            {
                string code = VbaLine.SplitCodeComment(m.Body[li]).Code;
                foreach (var tok in VbaLine.Identifiers(code))
                {
                    bool precededByDot = tok.Start > 0 && code[tok.Start - 1] == '.';
                    bool followedByDot = tok.End < code.Length && code[tok.End] == '.';

                    if (followedByDot && !precededByDot && moduleNames.Contains(tok.Text))
                        findings.Add(new DanglingRef(m.Name, li + 1, tok.Text, "module"));
                    else if (!precededByDot && !followedByDot && publicSymbols.Contains(tok.Text))
                        findings.Add(new DanglingRef(m.Name, li + 1, tok.Text, "symbol"));
                }
            }
        }
        return findings;
    }

    /// <summary>
    /// 한 모듈 본문의 공개 표면 심볼: (a) Private 아닌 Sub/Function/Property — 표준/클래스
    /// 모듈에서 수식어 없는 프로시저는 기본 Public, (b) 명시 Public/Global Const/Enum/Type/
    /// Declare/변수. 데이터 심볼은 명시 Public일 때만(모듈 수준 무수식어는 Private).
    /// </summary>
    public static IEnumerable<string> PublicSymbols(IReadOnlyList<string> body)
    {
        foreach (var line in body)
        {
            string code = VbaLine.SplitCodeComment(line).Code.Trim();
            if (code.Length == 0 || VbaLine.IsRemComment(code)) continue;

            var words = code.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) continue;

            int idx = 0;
            bool explicitPublic = false, isPrivate = false;
            while (idx < words.Length && Modifiers.Contains(words[idx]))
            {
                if (words[idx].Equals("Public", StringComparison.OrdinalIgnoreCase) ||
                    words[idx].Equals("Global", StringComparison.OrdinalIgnoreCase) ||
                    words[idx].Equals("Friend", StringComparison.OrdinalIgnoreCase))
                    explicitPublic = true;
                idx++;
            }
            if (words[0].Equals("Private", StringComparison.OrdinalIgnoreCase)) isPrivate = true;
            if (idx >= words.Length) continue;

            string kw = words[idx];
            if (ProcKeywords.Contains(kw) && !isPrivate)
            {
                string? name = ProcName(words, idx);
                if (name != null) yield return name;
            }
            else if (explicitPublic && PublicDataKeywords.Contains(kw))
            {
                string? name = DataName(words, idx);
                if (name != null) yield return name;
            }
            else if (explicitPublic)
            {
                // 명시 Public 변수 선언: Public gFoo As Long / Public WithEvents obj As X.
                // 수식어를 이미 소비했으므로 kw가 곧 변수명(또는 WithEvents 뒤 이름).
                if (kw.Equals("WithEvents", StringComparison.OrdinalIgnoreCase))
                    kw = idx + 1 < words.Length ? words[idx + 1] : "";
                string name = StripName(kw);
                if (name.Length > 0) yield return name;
            }
        }
    }

    /// <summary>Property Get/Let/Set의 접근자 키워드를 건너뛰고 프로시저명 추출.</summary>
    private static string? ProcName(string[] words, int kwIdx)
    {
        int nameIdx = kwIdx + 1;
        if (words[kwIdx].Equals("Property", StringComparison.OrdinalIgnoreCase))
            nameIdx = kwIdx + 2; // Get/Let/Set
        return nameIdx < words.Length ? StripName(words[nameIdx]) : null;
    }

    /// <summary>Const/Enum/Type/Declare의 대상명 추출(Declare는 [PtrSafe] Sub/Function 뒤).</summary>
    private static string? DataName(string[] words, int kwIdx)
    {
        if (words[kwIdx].Equals("Declare", StringComparison.OrdinalIgnoreCase))
        {
            int j = kwIdx + 1;
            if (j < words.Length && words[j].Equals("PtrSafe", StringComparison.OrdinalIgnoreCase)) j++;
            j++; // Sub|Function
            return j < words.Length ? StripName(words[j]) : null;
        }
        return kwIdx + 1 < words.Length ? StripName(words[kwIdx + 1]) : null;
    }

    /// <summary>선언 토큰에서 이름만(뒤따르는 `(`·`,`·타입 접미 제거).</summary>
    private static string StripName(string token)
    {
        int end = 0;
        while (end < token.Length &&
               (char.IsAsciiLetterOrDigit(token[end]) || token[end] == '_'))
            end++;
        return token[..end];
    }
}
