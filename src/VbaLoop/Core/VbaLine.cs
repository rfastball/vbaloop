namespace VbaLoop.Core;

/// <summary>
/// VBA 물리 라인의 최소 렉싱(설계 §9.5). `'` 주석은 문자열 리터럴 밖에서만 주석이므로
/// (`x = "it's"`의 `'`는 주석이 아님) 정규식이 아니라 실제 문자 스캔이 필요하다.
/// 주석제거(§9.5)와 컴파일 안전 스캔(§9.2)이 공유하는 토크나이저 바닥.
/// </summary>
public static class VbaLine
{
    /// <summary>
    /// 라인을 (코드부, 주석부)로 분리. 주석부는 여는 `'`부터 줄 끝까지(없으면 null).
    /// 문자열 리터럴 내부의 `'`와 `""`(이스케이프된 따옴표)를 정확히 건너뛴다.
    /// </summary>
    public static (string Code, string? Comment) SplitCodeComment(string line)
    {
        bool inStr = false;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (inStr)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"') i++; // 이스케이프된 따옴표
                    else inStr = false;
                }
            }
            else if (c == '"')
            {
                inStr = true;
            }
            else if (c == '\'')
            {
                return (line[..i], line[i..]);
            }
        }
        return (line, null);
    }

    /// <summary>선두 토큰이 REM 주석 키워드인지(코드부 기준, 대소문자 무시).</summary>
    public static bool IsRemComment(string code)
    {
        string t = code.TrimStart();
        return t.Equals("REM", StringComparison.OrdinalIgnoreCase)
               || t.StartsWith("REM ", StringComparison.OrdinalIgnoreCase)
               || t.StartsWith("REM\t", StringComparison.OrdinalIgnoreCase);
    }

    public readonly record struct Token(string Text, int Start, int End);

    /// <summary>
    /// 코드부에서 VBA 식별자 토큰(`[A-Za-z][A-Za-z0-9_]*`)을 위치와 함께 추출.
    /// 문자열/주석은 호출부가 <see cref="SplitCodeComment"/>로 먼저 제거한다.
    /// End는 토큰 바로 다음 인덱스(반열림) — 앞/뒤 `.` 인접 판정용.
    /// </summary>
    public static List<Token> Identifiers(string code)
    {
        var tokens = new List<Token>();
        int i = 0, n = code.Length;
        while (i < n)
        {
            char c = code[i];
            if (IsIdentStart(c))
            {
                int start = i;
                i++;
                while (i < n && IsIdentPart(code[i])) i++;
                tokens.Add(new Token(code[start..i], start, i));
            }
            else if (c == '"')
            {
                // 코드부에도 문자열이 남아있을 수 있다(주석만 잘렸으므로) — 내용 스킵
                i++;
                while (i < n)
                {
                    if (code[i] == '"')
                    {
                        if (i + 1 < n && code[i + 1] == '"') { i += 2; continue; }
                        i++;
                        break;
                    }
                    i++;
                }
            }
            else
            {
                i++;
            }
        }
        return tokens;
    }

    private static bool IsIdentStart(char c) => char.IsAsciiLetter(c) || c == '_';
    private static bool IsIdentPart(char c) => char.IsAsciiLetterOrDigit(c) || c == '_';
}
