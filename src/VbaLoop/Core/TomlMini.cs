namespace VbaLoop.Core;

/// <summary>
/// vbaloop.toml 지원 서브셋 한정 미니 파서(설계 §5).
/// 지원: [section] 헤더, key = "문자열" | true | false | 정수 | ["문자열", ...], # 라인/인라인 주석.
/// 정수는 10진 부호 있는 int 범위만([backup] keep 등) — 실수/16진/밑줄 구분자는 미지원.
/// 문자열 배열은 한 줄(멀티라인 미지원) — [dist] exclude 글롭 보강용(§9.1).
/// 미지원 문법은 hint 포함 에러로 즉시 실패한다(조용한 오파싱 금지).
/// 문자열 값 내부의 백슬래시는 이스케이프가 아니라 문자 그대로다(Windows 경로 친화).
/// </summary>
public static class TomlMini
{
    public static Dictionary<string, Dictionary<string, object>> Parse(string text, string sourceName)
    {
        var result = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, object>? current = null;
        var lines = text.Replace("\r\n", "\n").Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            int lineNo = i + 1;
            string raw = lines[i];
            string line = raw.Trim();
            if (line.Length == 0 || line.StartsWith('#')) continue;

            if (line.StartsWith('['))
            {
                if (!line.EndsWith(']'))
                    throw Bad(sourceName, lineNo, raw, "a section header must be a single [name] line");
                string section = line[1..^1].Trim();
                if (section.Length == 0 || section.Contains('[') || section.Contains(']'))
                    throw Bad(sourceName, lineNo, raw, "section name is empty, or nested sections are not supported");
                if (!result.TryGetValue(section, out current))
                {
                    current = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    result[section] = current;
                }
                continue;
            }

            int eq = line.IndexOf('=');
            if (eq < 1)
                throw Bad(sourceName, lineNo, raw, "not in key = value form");
            if (current == null)
                throw Bad(sourceName, lineNo, raw, "key appears before any section header (e.g. [project])");

            string key = line[..eq].Trim();
            string valuePart = line[(eq + 1)..].Trim();
            if (key.Length == 0 || key.Contains(' ') || key.Contains('"'))
                throw Bad(sourceName, lineNo, raw, "invalid key name format");

            object value;
            if (valuePart.StartsWith('"'))
            {
                int close = valuePart.IndexOf('"', 1);
                if (close < 0)
                    throw Bad(sourceName, lineNo, raw, "missing closing quote");
                string rest = valuePart[(close + 1)..].Trim();
                if (rest.Length > 0 && !rest.StartsWith('#'))
                    throw Bad(sourceName, lineNo, raw, "content other than a comment follows the string value");
                value = valuePart[1..close];
            }
            else if (valuePart.StartsWith('['))
            {
                value = ParseStringArray(valuePart, sourceName, lineNo, raw);
            }
            else
            {
                string bare = StripInlineComment(valuePart).Trim();
                value = bare switch
                {
                    "true" => true,
                    "false" => false,
                    _ when IsDecimalInteger(bare) => int.TryParse(bare, System.Globalization.NumberStyles.AllowLeadingSign,
                        System.Globalization.CultureInfo.InvariantCulture, out var n)
                            ? n
                            : throw Bad(sourceName, lineNo, raw, "integer value is outside int range"),
                    _ => throw Bad(sourceName, lineNo, raw,
                        "supported values are only \"string\", true, false, integer, [\"string\", ...] (floats/tables not supported)"),
                };
            }

            if (current.ContainsKey(key))
                throw Bad(sourceName, lineNo, raw, $"duplicate key: {key}");
            current[key] = value;
        }

        return result;
    }

    /// <summary>한 줄 문자열 배열 ["a", "b"] 파싱. 뒤 인라인 주석 허용. 이스케이프 미지원(문자열 값과 동일).</summary>
    private static List<string> ParseStringArray(string s, string source, int lineNo, string raw)
    {
        var list = new List<string>();
        int i = 1; // 여는 '[' 다음
        while (true)
        {
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            if (i < s.Length && s[i] == ']') { i++; break; }
            if (i >= s.Length || s[i] != '"')
                throw Bad(source, lineNo, raw, "array elements must be \"strings\"");
            int close = s.IndexOf('"', i + 1);
            if (close < 0) throw Bad(source, lineNo, raw, "missing closing quote on array element");
            list.Add(s[(i + 1)..close]);
            i = close + 1;
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            if (i < s.Length && s[i] == ',') { i++; continue; }
            if (i < s.Length && s[i] == ']') { i++; break; }
            throw Bad(source, lineNo, raw, "expected , or ] after array element");
        }
        string rest = s[i..].Trim();
        if (rest.Length > 0 && !rest.StartsWith('#'))
            throw Bad(source, lineNo, raw, "content other than a comment follows the array value");
        return list;
    }

    /// <summary>[+-]?숫자+ 만 정수로 인정 — 1_000, 0x10, 1.5, 1e3 등은 미지원 문법으로 실패시킨다.</summary>
    private static bool IsDecimalInteger(string s)
    {
        int start = s.Length > 0 && (s[0] == '+' || s[0] == '-') ? 1 : 0;
        if (s.Length == start) return false;
        for (int i = start; i < s.Length; i++)
            if (s[i] < '0' || s[i] > '9') return false;
        return true;
    }

    private static string StripInlineComment(string s)
    {
        int hash = s.IndexOf('#');
        return hash >= 0 ? s[..hash] : s;
    }

    private static CliError Bad(string source, int lineNo, string raw, string why)
        => new($"{source}:{lineNo}: TOML parse failed ({why}): {raw.Trim()}",
               "vbaloop.toml supports only [section] / key = \"string\"|true|false|integer / # comments");
}
