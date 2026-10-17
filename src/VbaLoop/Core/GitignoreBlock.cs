namespace VbaLoop.Core;

/// <summary>
/// 프로젝트 .gitignore의 vbaloop 관리 블록 upsert(설계 §9.7). <see cref="ManagedBlock"/>은 HTML
/// 주석 마커 + '## ' 마크다운 레거시 마이그레이션이라 재사용 불가 — .gitignore는 '#' 주석 마커로
/// 별도 관리한다. 사용자 소유 파일이므로 전체 덮어쓰기 없이 마커 사이만 교체(없으면 EOF append),
/// 원본 EOL 보존.
/// </summary>
public static class GitignoreBlock
{
    public const string StartMarker = "# vbaloop:managed:start";
    public const string EndMarker = "# vbaloop:managed:end";

    /// <summary>dist가 낳는 배포 아티팩트 폴더 + 도구 스크래치 — 커밋 대상 아님(§9.7).</summary>
    public static readonly string[] ManagedLines = { "_dbg/", "dist/" };

    /// <summary>
    /// existingText가 null이면 파일 부재 → 블록만으로 새 파일(CRLF)을 만든다. 마커가 있으면 그
    /// 사이를 교체(end 마커 유실 시 start부터 EOF까지 잔해로 간주해 복구), 없으면 EOF에 append
    /// (파일이 비어있지 않으면 구분 빈 줄 1개). 원본 EOL(LF/CRLF)을 보존한다.
    /// </summary>
    public static (string NewText, bool Changed) Upsert(string? existingText, IReadOnlyList<string> bodyLines)
    {
        var block = RenderBlock(bodyLines);

        if (existingText == null)
            return (CanonicalText.ToCanonical(block), true);

        string eol = existingText.Contains("\r\n") || !existingText.Contains('\n') ? "\r\n" : "\n";
        var lines = CanonicalText.SplitLines(existingText).ToList();
        if (FindTrimmed(lines, StartMarker, 0) < 0) // 개명 전 마커(vbaflow:managed) → 새 마커
            for (int i = 0; i < lines.Count; i++)
            {
                string tr = lines[i].Trim();
                if (tr == LegacyNames.GitignoreStart) lines[i] = StartMarker;
                else if (tr == LegacyNames.GitignoreEnd) lines[i] = EndMarker;
            }

        List<string> merged;
        int start = FindTrimmed(lines, StartMarker, 0);
        if (start >= 0)
        {
            int end = FindTrimmed(lines, EndMarker, start + 1);
            if (end < 0) end = lines.Count - 1; // end 마커 유실(사용자 훼손) — 잔해 복구
            merged = new List<string>(lines[..start]);
            merged.AddRange(block);
            if (end + 1 < lines.Count) merged.AddRange(lines[(end + 1)..]);
        }
        else
        {
            merged = new List<string>(lines);
            if (merged.Count > 0 && merged[^1].Trim().Length > 0)
                merged.Add(string.Empty);
            merged.AddRange(block);
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

    private static int FindTrimmed(List<string> lines, string target, int startIndex)
    {
        for (int i = startIndex; i < lines.Count; i++)
            if (lines[i].Trim() == target) return i;
        return -1;
    }
}
