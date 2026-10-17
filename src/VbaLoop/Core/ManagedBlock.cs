namespace VbaLoop.Core;

/// <summary>
/// CLAUDE.md/AGENTS.md의 vbaloop 관리 블록 upsert (v1.1 init).
/// 우선순위: 마커 사이 교체 > 마커 없는 블록 헤딩 교체 > EOF append.
/// 원본 파일의 EOL(LF/CRLF)을 보존해 전체 파일 churn을 피한다.
/// </summary>
public static class ManagedBlock
{
    public const string StartMarker = "<!-- vbaloop:managed:start -->";
    public const string EndMarker = "<!-- vbaloop:managed:end -->";
    /// <summary>블록 자신의 헤딩 — 마커가 지워진 블록을 알아보고, end 마커 유실 복구 시 블록 경계로 오인하지 않는다.</summary>
    private const string OwnHeading = "## Development loop (vbaloop)";
    /// <summary>한국어 시절의 헤딩 — 이미 배포된 블록(마커 유실 포함)을 계속 알아본다.</summary>
    private const string LegacyOwnHeading = "## 개발 루프 (vbaloop)";

    /// <summary>
    /// existingText가 null이면 파일 부재로 간주해 블록만으로 새 문서(CRLF)를 만든다.
    /// 반환 텍스트는 파일 끝 개행 정확히 1개로 정규화된다.
    /// </summary>
    public static (string NewText, bool Changed) Upsert(string? existingText, IReadOnlyList<string> bodyLines)
    {
        var block = RenderBlock(bodyLines);

        if (existingText == null)
            return (CanonicalText.ToCanonical(block), true);

        // EOL 보존: CRLF가 하나라도 있거나 개행이 아예 없으면 CRLF, 순수 LF 파일은 LF 유지
        string eol = existingText.Contains("\r\n") || !existingText.Contains('\n') ? "\r\n" : "\n";
        var lines = CanonicalText.SplitLines(existingText).ToList();

        List<string> merged;
        int start = FindTrimmed(lines, 0, StartMarker);
        if (start >= 0)
        {
            int end = FindTrimmed(lines, start + 1, EndMarker);
            if (end < 0)
            {
                // end 마커 유실: 블록 자신의 헤딩은 건너뛰고
                // 다음 "## " 섹션 직전(없으면 EOF)까지를 블록 잔해로 간주해 복구
                end = TrimBack(lines, start, FindNextH2ExceptOwn(lines, start + 1) - 1);
            }
            merged = Splice(lines, start, end, block);
        }
        else
        {
            int heading = FindTrimmed(lines, 0, OwnHeading);
            if (heading < 0) heading = FindTrimmed(lines, 0, LegacyOwnHeading);
            if (heading >= 0)
            {
                int next = FindNextH2ExceptOwn(lines, heading + 1); // exclusive 경계
                merged = Splice(lines, heading, TrimBack(lines, heading, next - 1), block);
            }
            else
            {
                merged = new List<string>(lines);
                if (merged.Count > 0 && merged[^1].Trim().Length > 0)
                    merged.Add(string.Empty);
                merged.AddRange(block);
            }
        }

        string newText = string.Join(eol, merged) + eol;
        return (newText, newText != existingText);
    }

    private static List<string> RenderBlock(IReadOnlyList<string> bodyLines)
    {
        var block = new List<string> { StartMarker };
        block.AddRange(CanonicalText.TrimTrailingEmpty(bodyLines));
        block.Add(EndMarker);
        return block;
    }

    /// <summary>[from..to] 구간(양끝 포함)을 block으로 교체. 교체 구간 양옆의 잉여 빈 줄은 그대로 둔다.</summary>
    private static List<string> Splice(List<string> lines, int from, int to, List<string> block)
    {
        var merged = new List<string>(lines[..from]);
        merged.AddRange(block);
        if (to + 1 < lines.Count) merged.AddRange(lines[(to + 1)..]);
        return merged;
    }

    /// <summary>교체 구간 끝의 빈 줄들은 구간에서 제외해 다음 섹션과의 구분 빈 줄을 보존.</summary>
    private static int TrimBack(List<string> lines, int from, int to)
    {
        while (to > from && lines[to].Trim().Length == 0) to--;
        return to;
    }

    private static int FindTrimmed(List<string> lines, int startIndex, string target)
    {
        for (int i = startIndex; i < lines.Count; i++)
            if (lines[i].Trim() == target) return i;
        return -1;
    }

    private static bool IsOwnHeading(string line)
    {
        var t = line.Trim();
        return t == OwnHeading || t == LegacyOwnHeading;
    }

    /// <summary>startIndex 이후 첫 "## " 헤딩 인덱스(블록 자신의 헤딩은 제외). 없으면 lines.Count.</summary>
    private static int FindNextH2ExceptOwn(List<string> lines, int startIndex)
    {
        for (int i = startIndex; i < lines.Count; i++)
            if (lines[i].StartsWith("## ") && !IsOwnHeading(lines[i])) return i;
        return lines.Count;
    }
}
