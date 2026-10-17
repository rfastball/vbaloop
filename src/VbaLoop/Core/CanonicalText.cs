using System.Text;

namespace VbaLoop.Core;

/// <summary>
/// 정규 텍스트 규약(설계 §4): src 파일은 UTF-8 BOM 없음 + CRLF + 파일 끝 개행 정확히 1개.
/// 읽기는 BOM/개행 혼재를 관용하고(기존 저장소 실측), 쓰기는 항상 규약대로.
/// </summary>
public static class CanonicalText
{
    public static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    /// <summary>파일을 UTF-8로 읽고 선행 BOM(U+FEFF)을 항상 스트립.</summary>
    public static string ReadFileStripBom(string path)
    {
        var text = File.ReadAllText(path, Encoding.UTF8); // 바이트 BOM은 디코더가 소비
        if (text.Length > 0 && text[0] == '\uFEFF') text = text[1..];
        return text;
    }

    /// <summary>CRLF/LF/CR 혼재를 관용하며 라인 배열로 분해. 마지막 빈 라인(trailing 개행)은 제거.</summary>
    public static string[] SplitLines(string text)
    {
        var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        int end = lines.Length;
        // trailing 개행이 만든 마지막 빈 원소 하나만 제거 (본문 중간의 빈 줄은 보존)
        if (end > 0 && lines[end - 1].Length == 0) end--;
        return lines[..end];
    }

    /// <summary>라인들을 CRLF로 결합, 파일 끝 개행 정확히 1개.</summary>
    public static string ToCanonical(IReadOnlyList<string> lines)
    {
        if (lines.Count == 0) return "\r\n";
        return string.Join("\r\n", lines) + "\r\n";
    }

    /// <summary>끝쪽 빈 라인 제거. 폼 Import 시 VBE가 코드비하인드 끝에 빈 줄을 추가하는
    /// 실측 동작 때문에 폼 경로 양쪽을 이 규칙으로 정규화해야 왕복이 수렴한다.</summary>
    public static string[] TrimTrailingEmpty(IReadOnlyList<string> lines)
    {
        int end = lines.Count;
        while (end > 0 && lines[end - 1].Trim().Length == 0) end--;
        return lines.Take(end).ToArray();
    }

    public static void WriteCanonical(string path, IReadOnlyList<string> lines)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, ToCanonical(lines), Utf8NoBom);
    }

    /// <summary>
    /// VBE 프로젝트 전역 케이스 테이블의 재케이싱을 src 철자로 흡수한다.
    /// <see cref="VbaCasing.ReconcileBlock"/>로 위임 — 토큰 단위 등가관계 + LCS 정렬이라 라인
    /// 삽입/삭제에 견고하고, 문자열 리터럴·주석 변경은 되돌리지 않는다(구 인덱스 정렬 대체).
    /// </summary>
    public static IReadOnlyList<string> PreserveCasing(
        IReadOnlyList<string> workbook, IReadOnlyList<string> src)
        => VbaCasing.ReconcileBlock(workbook, src);

    /// <summary>
    /// VBE 폼 Export가 WithEvents 등 멤버 선언에 자동 삽입하는 멤버/프로시저 수준 Attribute 라인
    /// (예: <c>Attribute m_x.VB_VarHelpID = -1</c>)을 제거한다. push는 이런
    /// 라인을 "코드 시작 이후 Attribute — 왕복 불가"로 거부하므로, pull이 같은 라인을 src에 쓰면
    /// pull→push 항등성이 깨진다. 속성명에 <c>.</c>이 있으면 멤버/프로시저 수준으로 판정 —
    /// 모듈 헤더(<c>VB_Name</c>/<c>VB_Exposed</c> 등)는 <c>.</c>이 없어 항상 보존된다.
    /// </summary>
    public static IReadOnlyList<string> StripMemberAttributes(IReadOnlyList<string> lines)
    {
        var result = new List<string>(lines.Count);
        foreach (var line in lines)
            if (!IsMemberAttribute(line)) result.Add(line);
        return result;
    }

    private static bool IsMemberAttribute(string line)
    {
        string t = line.TrimStart();
        const string prefix = "Attribute ";
        if (!t.StartsWith(prefix, StringComparison.Ordinal)) return false;
        int eq = t.IndexOf('=', prefix.Length);
        string namePart = eq >= 0 ? t[prefix.Length..eq] : t[prefix.Length..];
        return namePart.Contains('.');
    }
}
