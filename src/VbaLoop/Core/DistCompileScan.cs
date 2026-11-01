namespace VbaLoop.Core;

/// <summary>
/// 컴파일 안전 검증(설계 §9.2): 제외 후 "잔존 모듈이 사라진 심볼을 참조 → 배포본이 안 열림"을
/// 막는다. 헤드리스 전체 강제 컴파일은 원천 불가이므로 <b>본문 토큰 참조 스캔이 권위</b>다
/// (`@depends`는 힌트일 뿐 신뢰 원천 아님 — modEngine 실측). 제외 모듈들의 {모듈명 ∪ Public
/// 심볼명} 집합을 만들고, 잔존 모듈 본문에서 그 이름 참조를 찾으면 1건이라도 dist 거부.
/// 심볼 후보는 제외 <b>표준</b> 모듈의 Public만(클래스 멤버는 비한정 참조 불가), 참조 모듈의
/// 지역/매개변수/모듈 수준 선언이나 잔존 표준 모듈 Public이 같은 이름을 가리면 제외한다.
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
        // 비한정으로 참조 가능한 공개 표면은 표준 모듈 것뿐 — 클래스/폼/문서 모듈 멤버는
        // 항상 개체 한정(`obj.Member`)이라 다른 모듈의 비한정 이름과 묶일 수 없다.
        var publicSymbols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in excluded.Where(IsStandardModule))
            foreach (var sym in PublicSymbols(m.Body))
                publicSymbols.Add(sym);
        // 제외 후에도 잔존 표준 모듈이 같은 Public 이름을 제공하면 참조는 그쪽으로 묶인다.
        var stillProvided = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var m in included.Where(IsStandardModule))
            foreach (var sym in PublicSymbols(m.Body))
                stillProvided.Add(sym);

        var findings = new List<DanglingRef>();
        foreach (var m in included)
        {
            var scope = ModuleScope.Build(m.Body);
            for (int li = 0; li < m.Body.Count; li++)
            {
                string code = VbaLine.SplitCodeComment(m.Body[li]).Code;
                foreach (var tok in VbaLine.Identifiers(code))
                {
                    if (tok.Start > 0 && code[tok.Start - 1] == '.') continue; // 멤버 접근
                    bool isModule = moduleNames.Contains(tok.Text);
                    if (!isModule && !publicSymbols.Contains(tok.Text)) continue;
                    // 지역/모듈 수준 선언(자기 Public 포함)이 같은 이름을 가리면 그쪽으로 바인딩
                    if (scope.Declares(li, tok.Text) || stillProvided.Contains(tok.Text)) continue;

                    bool followedByDot = tok.End < code.Length && code[tok.End] == '.';
                    if (isModule)
                        // `Mod.X` 한정 참조 또는 `As Cls`/`New Cls`/`Unload Frm` 같은 이름 자체 참조
                        findings.Add(new DanglingRef(m.Name, li + 1, tok.Text, "module"));
                    else if (!followedByDot)
                        findings.Add(new DanglingRef(m.Name, li + 1, tok.Text, "symbol"));
                }
            }
        }
        return findings;
    }

    private static bool IsStandardModule(DistModule m)
        => m.Ext.Equals(".bas", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 한 모듈의 이름 가림(shadowing) 표: 모듈 수준 선언(Private 포함 프로시저/변수/상수/Enum
    /// 멤버 등)과 프로시저별 지역(매개변수·Dim/Static/Const/ReDim). VBA 지역 선언은 프로시저
    /// 전체 범위라 위치와 무관하게 프로시저 단위로 모은다. 완전 파서가 아니라 선언 이름만
    /// 건지는 근사 — 놓쳐도 결과는 기존과 같은 보수적 판정(과검출)이다.
    /// </summary>
    private sealed class ModuleScope
    {
        private static readonly HashSet<string> DeclModifiers = new(StringComparer.OrdinalIgnoreCase)
            { "Public", "Private", "Friend", "Global", "Static", "Dim" };
        private static readonly HashSet<string> ItemPrefixes = new(StringComparer.OrdinalIgnoreCase)
            { "Optional", "ByVal", "ByRef", "ParamArray", "WithEvents", "Preserve" };

        private readonly HashSet<string> _module = new(StringComparer.OrdinalIgnoreCase);
        private HashSet<string>?[] _procByLine = Array.Empty<HashSet<string>?>();

        public bool Declares(int line, string name)
            => _module.Contains(name) || _procByLine[line]?.Contains(name) == true;

        public static ModuleScope Build(IReadOnlyList<string> body)
        {
            var s = new ModuleScope { _procByLine = new HashSet<string>?[body.Count] };
            HashSet<string>? proc = null;
            bool inEnum = false, inType = false;
            int li = 0;
            while (li < body.Count)
            {
                // 논리 라인: 줄 끝 ` _` 연속을 이어 붙인다(주석은 연속 대상 아님)
                int start = li;
                string logical = "";
                while (true)
                {
                    string code = VbaLine.SplitCodeComment(body[li]).Code.TrimEnd();
                    bool continues = code.EndsWith(" _") || code == "_";
                    logical += (continues ? code[..^1] : code) + " ";
                    li++;
                    if (!continues || li >= body.Count) break;
                }
                if (VbaLine.IsRemComment(logical)) continue;

                bool closes = false;
                foreach (string stmt in Statements(logical))
                {
                    var words = stmt.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                    if (words.Length == 0) continue;
                    if (words[0].Equals("End", StringComparison.OrdinalIgnoreCase) && words.Length > 1)
                    {
                        if (ProcKeywords.Contains(words[1])) closes = true;
                        else if (words[1].Equals("Enum", StringComparison.OrdinalIgnoreCase)) inEnum = false;
                        else if (words[1].Equals("Type", StringComparison.OrdinalIgnoreCase)) inType = false;
                        continue;
                    }
                    if (inEnum) { AddIdent(s._module, stmt); continue; } // Enum 멤버는 모듈(또는 전역) 이름
                    if (inType) continue;                                  // Type 멤버는 한정 접근만

                    int idx = 0;
                    while (idx < words.Length && DeclModifiers.Contains(words[idx])) idx++;
                    bool hadModifier = idx > 0;
                    if (idx >= words.Length) continue;
                    string kw = words[idx];
                    string rest = SkipWords(stmt, idx + 1);

                    if (ProcKeywords.Contains(kw))
                    {
                        string? name = ProcName(words, idx);
                        if (name == null) continue;
                        s._module.Add(name);
                        proc = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        int open = stmt.IndexOf('(');
                        if (open >= 0) AddDeclList(proc, ParenContent(stmt, open));
                    }
                    else if (kw.Equals("Const", StringComparison.OrdinalIgnoreCase))
                        AddDeclList(proc ?? s._module, rest);
                    else if (kw.Equals("ReDim", StringComparison.OrdinalIgnoreCase))
                        AddDeclList(proc ?? s._module, rest);
                    else if (proc == null && kw.Equals("Enum", StringComparison.OrdinalIgnoreCase))
                    { AddIdent(s._module, rest); inEnum = true; }
                    else if (proc == null && kw.Equals("Type", StringComparison.OrdinalIgnoreCase))
                    { AddIdent(s._module, rest); inType = true; }
                    else if (proc == null && (kw.Equals("Declare", StringComparison.OrdinalIgnoreCase) ||
                                              kw.Equals("Event", StringComparison.OrdinalIgnoreCase)))
                    {
                        string? name = kw.Equals("Event", StringComparison.OrdinalIgnoreCase)
                            ? StripName(rest.TrimStart()) : DataName(words, idx);
                        if (!string.IsNullOrEmpty(name)) s._module.Add(name);
                    }
                    else if (hadModifier)
                        AddDeclList(proc ?? s._module, SkipWords(stmt, idx));
                }
                for (int i = start; i < li; i++) s._procByLine[i] = proc;
                if (closes) proc = null;
            }
            return s;
        }

        private static void AddIdent(HashSet<string> set, string text)
        {
            string name = StripName(text.TrimStart());
            if (name.Length > 0) set.Add(name);
        }

        /// <summary>`a As Long, b(1 To 2, 3) As String, Optional ByVal c = "x,y"`의 각 항목 이름.</summary>
        private static void AddDeclList(HashSet<string> set, string list)
        {
            foreach (string item in SplitTopLevel(list, ','))
            {
                var words = item.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                int j = 0;
                while (j < words.Length && ItemPrefixes.Contains(words[j])) j++;
                if (j >= words.Length) continue;
                string name = StripName(words[j]);
                if (name.Length > 0) set.Add(name);
            }
        }

        /// <summary>문장 구분 `:`로 분할 — 문자열 안과 명명 인수 `:=`는 제외.</summary>
        private static IEnumerable<string> Statements(string logical)
        {
            foreach (string part in SplitTopLevel(logical, ':', ignoreParens: true))
                yield return part.Trim();
        }

        /// <summary>문자열 밖(옵션: 괄호 깊이 0)에서만 구분자로 분할. `:=`의 `:`는 구분자가 아니다.</summary>
        private static List<string> SplitTopLevel(string text, char sep, bool ignoreParens = false)
        {
            var parts = new List<string>();
            int depth = 0, from = 0;
            bool inStr = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (inStr) { if (c == '"') inStr = false; continue; } // `""`는 닫힘+열림으로 상쇄
                if (c == '"') inStr = true;
                else if (c == '(') depth++;
                else if (c == ')') depth--;
                else if (c == sep && (ignoreParens || depth == 0) &&
                         !(sep == ':' && i + 1 < text.Length && text[i + 1] == '='))
                {
                    parts.Add(text[from..i]);
                    from = i + 1;
                }
            }
            parts.Add(text[from..]);
            return parts;
        }

        /// <summary>`open` 위치의 `(`와 짝인 `)` 사이 내용(문자열 안 괄호 무시).</summary>
        private static string ParenContent(string text, int open)
        {
            int depth = 0;
            bool inStr = false;
            for (int i = open; i < text.Length; i++)
            {
                char c = text[i];
                if (inStr) { if (c == '"') inStr = false; continue; }
                if (c == '"') inStr = true;
                else if (c == '(') depth++;
                else if (c == ')' && --depth == 0) return text[(open + 1)..i];
            }
            return text[(open + 1)..];
        }

        /// <summary>공백 구분 단어 n개를 건너뛴 나머지 원문.</summary>
        private static string SkipWords(string text, int n)
        {
            int i = 0;
            for (int w = 0; w < n; w++)
            {
                while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
                while (i < text.Length && !char.IsWhiteSpace(text[i])) i++;
            }
            return text[i..];
        }
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
