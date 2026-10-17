namespace VbaLoop.Core;

/// <summary>
/// CLAUDE.md/AGENTS.md의 vbaloop 관리 블록 upsert (v1.1 init).
/// 우선순위: 마커 사이 교체 > 레거시 헤딩 마이그레이션 > EOF append.
/// 원본 파일의 EOL(LF/CRLF)을 보존해 전체 파일 churn을 피한다.
/// </summary>
public static class ManagedBlock
{
    public const string StartMarker = "<!-- vbaloop:managed:start -->";
    public const string EndMarker = "<!-- vbaloop:managed:end -->";
    /// <summary>마커 도입 전 수동 작성된 섹션 헤딩(레거시). 발견 시 마커 블록으로 마이그레이션.</summary>
    public const string LegacyHeading = LegacyNames.DocsHeading;
    /// <summary>블록 자신의 헤딩들(현재 + 개명 전) — end 마커 유실 복구 시 블록 경계로 오인하지 않는다.</summary>
    private static readonly string[] OwnHeadings = { LegacyNames.DocsHeading, "## 개발 루프 (vbaloop)" };

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
        MigrateLegacyMarkers(lines);

        List<string> merged;
        int start = FindTrimmed(lines, 0, StartMarker);
        if (start >= 0)
        {
            int end = FindTrimmed(lines, start + 1, EndMarker);
            if (end < 0)
            {
                // end 마커 유실: 블록 자신의 헤딩(LegacyHeading)은 건너뛰고
                // 다음 "## " 섹션 직전(없으면 EOF)까지를 블록 잔해로 간주해 복구
                end = TrimBack(lines, start, FindNextH2ExceptLegacy(lines, start + 1) - 1);
            }
            merged = Splice(lines, start, end, block);
        }
        else
        {
            int heading = OwnHeadings.Select(h => FindTrimmed(lines, 0, h)).Where(i => i >= 0).DefaultIfEmpty(-1).Min();
            if (heading >= 0)
            {
                int next = FindNextH2ExceptLegacy(lines, heading + 1); // exclusive 경계
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

    /// <summary>startIndex 이후 첫 "## " 헤딩 인덱스(LegacyHeading 자신은 제외). 없으면 lines.Count.</summary>
    private static int FindNextH2ExceptLegacy(List<string> lines, int startIndex)
    {
        for (int i = startIndex; i < lines.Count; i++)
            if (lines[i].StartsWith("## ") && !OwnHeadings.Contains(lines[i].Trim())) return i;
        return lines.Count;
    }

    /// <summary>개명 전 마커(vbaflow:managed)를 새 마커로 치환 — 새 마커가 이미 있으면 손대지 않는다.</summary>
    private static void MigrateLegacyMarkers(List<string> lines)
    {
        if (FindTrimmed(lines, 0, StartMarker) >= 0) return;
        for (int i = 0; i < lines.Count; i++)
        {
            string tr = lines[i].Trim();
            if (tr == LegacyNames.ManagedStart) lines[i] = StartMarker;
            else if (tr == LegacyNames.ManagedEnd) lines[i] = EndMarker;
        }
    }
}
