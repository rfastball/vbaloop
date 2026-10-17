using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop lint — 규약/인코딩/헤더 전수 검사 (설계 §8). 세션/Excel/lease 불필요(읽기 전용,
/// toml만 요구). push의 파일 단위 가드를 레포 전체에 일괄 실행하는 사전 검사 버전.
/// </summary>
public static class LintCommand
{
    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "lint" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "lint");
        o.LogRecords = log.Records;

        int codePage = NativeMethods.GetACP(); // 상수 하드코딩 금지
        var findings = LintChecks.CheckProject(cfg.SrcRoot, codePage);

        int fileCount = Directory.Exists(cfg.SrcRoot)
            ? Directory.EnumerateFiles(cfg.SrcRoot)
                .Count(f => Path.GetExtension(f).ToLowerInvariant() is ".bas" or ".cls" or ".frm")
            : 0;

        o.Payload = new
        {
            codePage,
            filesChecked = fileCount,
            findingCount = findings.Count,
            files = findings
                .GroupBy(f => f.File, StringComparer.OrdinalIgnoreCase)
                .Select(g => new
                {
                    file = g.Key,
                    findings = g.Select(f => new { line = f.Line, code = f.Code, message = f.Message }),
                }),
        };

        if (findings.Count == 0)
        {
            o.HumanLines.Add($"lint 통과: {fileCount}개 파일, 위반 없음 (CP{codePage})");
            log.Write("info", "lint_clean", new { fileCount });
            return o;
        }

        foreach (var f in findings.Take(50))
            o.HumanLines.Add($"  {f.File}{(f.Line != null ? $":{f.Line}" : "")} [{f.Code}] {f.Message}");
        if (findings.Count > 50)
            o.HumanLines.Add($"  ... 외 {findings.Count - 50}건");

        log.Write("warn", "lint_findings", new { count = findings.Count });
        o.Fail($"lint 위반 {findings.Count}건 / {fileCount}개 파일",
            "위 항목 수정 후 재실행 — ENCODING은 push가 어차피 거부하므로 최우선");
        return o;
    }
}
