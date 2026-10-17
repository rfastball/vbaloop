using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>
/// 세션 불필요 전수 검사(설계 §8 lint): push의 MVP 가드가 파일 단위로 수행하는 검사의
/// 일괄 실행 버전 + 정규 텍스트 규약 검사. 규약 검사(BOM/개행/EOF)는 CanonicalText가
/// 관용 정규화로 위반을 지워버리므로 raw 바이트 경로로 신규 작성(검토 지적).
/// </summary>
public static class LintChecks
{
    public sealed record Finding(string File, int? Line, string Code, string Message);

    private static readonly Regex VbNameAttr = new(
        "^\\s*Attribute VB_Name = \"(?<name>.*)\"\\s*$", RegexOptions.Compiled);

    public static List<Finding> CheckProject(string srcRoot, int codePage)
    {
        var findings = new List<Finding>();
        if (!Directory.Exists(srcRoot))
        {
            findings.Add(new Finding(srcRoot, null, "SRC_MISSING", "src 루트가 없음"));
            return findings;
        }

        var sourceFiles = Directory.EnumerateFiles(srcRoot)
            .Where(f => Path.GetExtension(f).ToLowerInvariant() is ".bas" or ".cls" or ".frm")
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var frxFiles = Directory.EnumerateFiles(srcRoot)
            .Where(f => Path.GetExtension(f).ToLowerInvariant() == ".frx")
            .ToList();

        // basename 충돌: 같은 이름의 .bas와 .cls 등 — 컴포넌트 이름 공간은 하나다
        foreach (var group in sourceFiles.GroupBy(
                     f => Path.GetFileNameWithoutExtension(f), StringComparer.OrdinalIgnoreCase))
        {
            if (group.Count() > 1)
                findings.Add(new Finding(
                    string.Join(", ", group.Select(Path.GetFileName)), null, "BASENAME_CONFLICT",
                    $"컴포넌트 이름 충돌: {group.Key} — VBA 이름 공간은 확장자를 구분하지 않음"));
        }

        // frm↔frx 짝
        var frmBases = new HashSet<string>(
            sourceFiles.Where(f => Path.GetExtension(f).ToLowerInvariant() == ".frm")
                .Select(f => Path.GetFileNameWithoutExtension(f)),
            StringComparer.OrdinalIgnoreCase);
        foreach (var frm in frmBases)
        {
            if (!frxFiles.Any(x => Path.GetFileNameWithoutExtension(x)
                    .Equals(frm, StringComparison.OrdinalIgnoreCase)))
                findings.Add(new Finding(frm + ".frm", null, "FRX_MISSING",
                    "짝이 되는 .frx가 없음 — 폼 Import 불가"));
        }
        foreach (var frx in frxFiles)
        {
            if (!frmBases.Contains(Path.GetFileNameWithoutExtension(frx)))
                findings.Add(new Finding(Path.GetFileName(frx), null, "FRX_ORPHAN",
                    "짝이 되는 .frm이 없는 고아 .frx"));
        }

        foreach (var file in sourceFiles)
            CheckFile(file, codePage, findings);

        return findings;
    }

    public static void CheckFile(string file, int codePage, List<Finding> findings)
    {
        string name = Path.GetFileName(file);
        byte[] raw = File.ReadAllBytes(file);

        // ---- raw 바이트 규약 검사 (UTF-8 BOM 없음 + CRLF + EOF 개행 정확히 1개) ----
        if (raw.Length >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF)
            findings.Add(new Finding(name, 1, "BOM", "UTF-8 BOM 존재 — 규약은 BOM 없음"));

        int loneLfAt = -1, bareCrAt = -1;
        for (int i = 0; i < raw.Length; i++)
        {
            if (raw[i] == 0x0A && (i == 0 || raw[i - 1] != 0x0D) && loneLfAt < 0) loneLfAt = i;
            if (raw[i] == 0x0D && (i + 1 >= raw.Length || raw[i + 1] != 0x0A) && bareCrAt < 0) bareCrAt = i;
        }
        if (loneLfAt >= 0)
            findings.Add(new Finding(name, LineOf(raw, loneLfAt), "EOL_LF",
                "CR 없는 LF 개행 — 규약은 CRLF"));
        if (bareCrAt >= 0)
            findings.Add(new Finding(name, LineOf(raw, bareCrAt), "EOL_CR",
                "LF 없는 단독 CR — 규약은 CRLF"));

        // 끝 개행 "부재"만 위반 — 본문 끝 빈 줄(모듈의 정당한 내용)은 이중 CRLF로
        // 직렬화되므로 개수 상한 검사는 위양성이 된다(Test_Project 실측)
        bool endsWithCrlf = raw.Length >= 2 && raw[^2] == 0x0D && raw[^1] == 0x0A;
        if (!endsWithCrlf)
            findings.Add(new Finding(name, null, "EOF_NEWLINE",
                raw.Length == 0 ? "빈 파일" : "파일 끝 CRLF 개행 부재"));

        // ---- 텍스트 계층 검사 (push 가드 재사용) ----
        var lines = CanonicalText.SplitLines(CanonicalText.ReadFileStripBom(file));

        foreach (var v in Cp949Guard.Check(lines, codePage))
            findings.Add(new Finding(name, v.Line, "ENCODING",
                $"{v.Column}열 '{v.Char}' ({v.CodePoint}) — CP{codePage} 표현 불가, push 시 확정 손실"));

        IReadOnlyList<string> header;
        try
        {
            header = VbaHeader.Split(lines, name, strict: true).Header;
        }
        catch (CliError e)
        {
            findings.Add(new Finding(name, null, "HEADER_ATTR_AFTER_CODE", e.Message));
            header = VbaHeader.Split(lines, name, strict: false).Header;
        }

        string? vbName = header.Select(l => VbNameAttr.Match(l))
            .Where(m => m.Success).Select(m => m.Groups["name"].Value).FirstOrDefault();
        string baseName = Path.GetFileNameWithoutExtension(file);
        if (vbName == null)
            findings.Add(new Finding(name, null, "NAME_MISSING",
                "Attribute VB_Name 헤더가 없음 — Export 호환 포맷 아님"));
        else if (!vbName.Equals(baseName, StringComparison.OrdinalIgnoreCase))
            findings.Add(new Finding(name, null, "NAME_MISMATCH",
                $"VB_Name(\"{vbName}\")과 파일명(\"{baseName}\") 불일치 — push는 파일명으로 매칭함"));
    }

    private static int LineOf(byte[] raw, int offset)
    {
        int line = 1;
        for (int i = 0; i < offset; i++)
            if (raw[i] == 0x0A) line++;
        return line;
    }
}
