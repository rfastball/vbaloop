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
            findings.Add(new Finding(srcRoot, null, "SRC_MISSING", "src root not found"));
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
                    $"component name collision: {group.Key} — the VBA namespace does not distinguish extensions"));
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
                    "no matching .frx — form Import impossible"));
        }
        foreach (var frx in frxFiles)
        {
            if (!frmBases.Contains(Path.GetFileNameWithoutExtension(frx)))
                findings.Add(new Finding(Path.GetFileName(frx), null, "FRX_ORPHAN",
                    "orphan .frx with no matching .frm"));
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
            findings.Add(new Finding(name, 1, "BOM", "UTF-8 BOM present — convention is no BOM"));

        int loneLfAt = -1, bareCrAt = -1;
        for (int i = 0; i < raw.Length; i++)
        {
            if (raw[i] == 0x0A && (i == 0 || raw[i - 1] != 0x0D) && loneLfAt < 0) loneLfAt = i;
            if (raw[i] == 0x0D && (i + 1 >= raw.Length || raw[i + 1] != 0x0A) && bareCrAt < 0) bareCrAt = i;
        }
        if (loneLfAt >= 0)
            findings.Add(new Finding(name, LineOf(raw, loneLfAt), "EOL_LF",
                "LF line ending without CR — convention is CRLF"));
        if (bareCrAt >= 0)
            findings.Add(new Finding(name, LineOf(raw, bareCrAt), "EOL_CR",
                "lone CR without LF — convention is CRLF"));

        // 끝 개행 "부재"만 위반 — 본문 끝 빈 줄(모듈의 정당한 내용)은 이중 CRLF로
        // 직렬화되므로 개수 상한 검사는 위양성이 된다(Test_Project 실측)
        bool endsWithCrlf = raw.Length >= 2 && raw[^2] == 0x0D && raw[^1] == 0x0A;
        if (!endsWithCrlf)
            findings.Add(new Finding(name, null, "EOF_NEWLINE",
                raw.Length == 0 ? "empty file" : "missing CRLF line ending at end of file"));

        // ---- 텍스트 계층 검사 (push 가드 재사용) ----
        var lines = CanonicalText.SplitLines(CanonicalText.ReadFileStripBom(file));

        foreach (var v in Cp949Guard.Check(lines, codePage))
            findings.Add(new Finding(name, v.Line, "ENCODING",
                $"column {v.Column} '{v.Char}' ({v.CodePoint}) — not representable in CP{codePage}, definitively lost on push"));

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
                "Attribute VB_Name header missing — not an Export-compatible format"));
        else if (!vbName.Equals(baseName, StringComparison.OrdinalIgnoreCase))
            findings.Add(new Finding(name, null, "NAME_MISMATCH",
                $"VB_Name(\"{vbName}\") and file name (\"{baseName}\") mismatch — push matches by file name"));
    }

    private static int LineOf(byte[] raw, int offset)
    {
        int line = 1;
        for (int i = 0; i < offset; i++)
            if (raw[i] == 0x0A) line++;
        return line;
    }
}
