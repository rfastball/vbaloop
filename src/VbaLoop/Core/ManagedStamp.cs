using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>exe 버전 vs 프로젝트 관리 스탬프의 관계(설계 §12.2).</summary>
public enum StampState
{
    /// <summary>스탬프 없음(스탬프 도입 이전 온보딩). 수렴 허용, 다음 init이 기록.</summary>
    Unstamped,
    /// <summary>exe == 스탬프.</summary>
    Current,
    /// <summary>exe &gt; 스탬프. 수렴 허용, 관리 파일을 쓰면 스탬프를 올린다.</summary>
    ExeNewer,
    /// <summary>exe &lt; 스탬프(또는 exe 버전 판독 불가). 관리 파일 쓰기 거부 — 해결은 exe 갱신.</summary>
    ExeOlder,
    /// <summary>스탬프가 semver가 아님. 쓰기 거부 — 사람이 toml을 확인해야 한다.</summary>
    Invalid,
}

/// <summary>
/// 관리 스탬프: vbaloop.toml의 <c>[vbaloop] managed_by = "&lt;버전&gt;"</c> = 관리 파일(CLAUDE.md/AGENTS.md 관리 블록,
/// 스킬, .gitignore 블록, LogShim.bas)을 마지막으로 쓴 exe 버전(설계 §12). 스탬프는 내려가지 않는다 —
/// 더 오래된 exe는 관리 파일을 쓰지 않으므로(init 거부) 프로젝트를 다운그레이드할 수 없다.
/// toml은 사용자 소유 파일이라 통째로 다시 쓰지 않고 이 한 줄만 라인 단위로 upsert한다.
/// 판독도 TomlMini 대신 라인 스캔이다 — 미래 버전 toml에 이 exe가 모르는 문법이 있어도 스탬프는 읽혀야
/// "exe가 오래됨"을 정확히 말할 수 있다.
/// </summary>
public static class ManagedStamp
{
    public const string Section = "vbaloop";
    public const string Key = "managed_by";
    public const string UpdateHint = "ask the user to run vbaloop update (do not run init or edit managed_by)";

    private static readonly Regex SectionHeader = new(@"^\s*\[\s*([^\[\]]+?)\s*\]\s*(#.*)?$");
    private static readonly Regex KeyLine = new(@"^\s*managed_by\s*=\s*""([^""]*)""\s*(#.*)?$");

    /// <summary>toml 텍스트에서 [vbaloop] managed_by 값. 없으면 null.</summary>
    public static string? Read(string tomlText)
    {
        string? section = null;
        foreach (var line in Lines(tomlText))
        {
            var h = SectionHeader.Match(line);
            if (h.Success) { section = h.Groups[1].Value; continue; }
            if (!string.Equals(section, Section, StringComparison.OrdinalIgnoreCase)) continue;
            var k = KeyLine.Match(line);
            if (k.Success) return k.Groups[1].Value;
        }
        return null;
    }

    /// <summary>파일에서 판독. 파일이 없거나 읽기 실패면 null(스탬프 없음과 같다).</summary>
    public static string? ReadFile(string? tomlPath)
    {
        if (tomlPath == null) return null;
        try { return File.Exists(tomlPath) ? Read(File.ReadAllText(tomlPath)) : null; }
        catch (IOException) { return null; }
        catch (UnauthorizedAccessException) { return null; }
    }

    /// <summary>판정 매트릭스. exe가 null(버전 판독 불가)이면 스탬프가 있는 한 쓰지 않는다(안전 방향).</summary>
    public static StampState Decide(SemVer? exe, string? stamp)
    {
        if (stamp == null) return StampState.Unstamped;
        var s = SemVer.ParseOrNull(stamp);
        if (s == null) return StampState.Invalid;
        if (exe == null) return StampState.ExeOlder;
        int c = exe.CompareTo(s);
        return c == 0 ? StampState.Current : c > 0 ? StampState.ExeNewer : StampState.ExeOlder;
    }

    /// <summary>이 상태에서 관리 파일을 써도 되는가.</summary>
    public static bool AllowsWrite(StampState s)
        => s is StampState.Unstamped or StampState.Current or StampState.ExeNewer;

    /// <summary>
    /// 관리 파일을 쓸 때 기록할 스탬프. 스탬프는 내려가지 않는다: 쓰기 허용 상태에서만 부르며,
    /// exe가 판독 불가면 기록하지 않는다(null).
    /// </summary>
    public static string? StampToWrite(SemVer? exe, StampState state)
        => AllowsWrite(state) && exe != null ? exe.ToString() : null;

    /// <summary>JSON/사람용 상태 문자열(kebab-case).</summary>
    public static string Label(StampState s) => s switch
    {
        StampState.Unstamped => "unstamped",
        StampState.Current => "current",
        StampState.ExeNewer => "exe-newer",
        StampState.ExeOlder => "exe-older",
        _ => "invalid",
    };

    /// <summary>exe-older 상태의 한 줄 설명(에러·경고 공용).</summary>
    public static string DescribeExeOlder(string stamp, string exeVersion)
        => $"This project is managed by vbaloop {stamp} but the running exe is older ({exeVersion})";

    /// <summary>
    /// [vbaloop] managed_by를 version으로 upsert. 다른 라인·주석·개행 스타일은 보존한다.
    /// 섹션이 있으면 그 안의 키를 교체(없으면 헤더 바로 아래 삽입), 섹션이 없으면 파일 끝에 덧붙인다.
    /// </summary>
    public static string Upsert(string tomlText, string version)
    {
        string nl = tomlText.Contains("\r\n") || tomlText.Length == 0 ? "\r\n" : "\n";
        var lines = Lines(tomlText).ToList();
        // Lines는 끝 개행 뒤의 빈 조각을 버리므로 재조립 시 끝 개행 1개로 정규화된다
        string keyLine = $"{Key} = \"{version}\"";

        int header = -1;
        for (int i = 0; i < lines.Count; i++)
        {
            var h = SectionHeader.Match(lines[i]);
            if (!h.Success) continue;
            if (header >= 0) break; // 우리 섹션이 끝남
            if (string.Equals(h.Groups[1].Value, Section, StringComparison.OrdinalIgnoreCase))
                header = i;
        }

        if (header >= 0)
        {
            int end = lines.Count;
            for (int i = header + 1; i < lines.Count; i++)
                if (SectionHeader.IsMatch(lines[i])) { end = i; break; }
            for (int i = header + 1; i < end; i++)
            {
                if (!KeyLine.IsMatch(lines[i])) continue;
                lines[i] = keyLine;
                return string.Join(nl, lines) + nl;
            }
            lines.Insert(header + 1, keyLine);
            return string.Join(nl, lines) + nl;
        }

        while (lines.Count > 0 && lines[^1].Trim().Length == 0) lines.RemoveAt(lines.Count - 1);
        if (lines.Count > 0) lines.Add("");
        lines.Add($"[{Section}]");
        lines.Add("# vbaloop version that last wrote the managed files. Recorded by vbaloop init - do not edit directly");
        lines.Add(keyLine);
        return string.Join(nl, lines) + nl;
    }

    private static IEnumerable<string> Lines(string text)
    {
        var parts = text.Replace("\r\n", "\n").Split('\n');
        int n = parts.Length;
        if (n > 0 && parts[n - 1].Length == 0) n--; // 끝 개행
        return parts.Take(n);
    }
}
