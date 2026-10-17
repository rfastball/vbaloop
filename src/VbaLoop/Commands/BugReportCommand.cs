using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Commands;

/// <summary>
/// bugreport: vbaloop 자체 결함을 표준 템플릿으로 남기는 창구. Excel/세션/toml 불필요
/// (toml 자체의 버그도 리포트 가능해야 하므로 유효 toml을 요구하지 않는다).
///
/// 기록 위치는 exe가 어디에 있느냐로 갈린다:
/// - 소스 레포 아래(개발): <c>vbaloop\docs\bugreports</c> — 리포트가 곧 레포 자산이다
/// - 설치본: 프로젝트의 <c>_dbg\bugreports</c> + payload.issueUrl(GitHub 새 이슈)
///   설치 사용자에게 이 명령은 유일한 피드백 창구라, 레포가 없다고 실패하면 창구가 닫힌다.
/// </summary>
public static class BugReportCommand
{
    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "bugreport" };

        var exeDir = Environment.ProcessPath != null
            ? Path.GetDirectoryName(Environment.ProcessPath)! : Environment.CurrentDirectory;
        var repoRoot = RepoLocator.FindRepoRoot(exeDir);

        // 프로젝트 루트: toml 위치, 없으면 cwd (설치본 폴백 기록 위치의 기준)
        var cwd = Environment.CurrentDirectory;
        var tomlPath = ProjectConfig.FindTomlPath(opts.Config, cwd);
        var projectRoot = tomlPath != null ? Path.GetDirectoryName(tomlPath)! : cwd;

        bool upstream = repoRoot != null;
        var reportsDir = upstream
            ? Path.Combine(repoRoot!, BugReport.ReportsRelDir)
            : Path.Combine(projectRoot, BugReport.LocalReportsRelDir);

        if (opts.List)
        {
            var reports = BugReport.ListReports(reportsDir);
            foreach (var r in reports)
                o.HumanLines.Add($"[{(r.Resolved ? "resolved" : "open")}] {r.File} - {r.Title}");
            if (reports.Count == 0)
                o.HumanLines.Add("리포트 없음");
            o.Payload = new
            {
                dir = reportsDir,
                upstream,
                reports = reports.Select(r => new { file = r.File, title = r.Title, resolved = r.Resolved }).ToArray(),
            };
            return o;
        }

        var title = string.Join(" ", opts.Positional).Trim();
        if (title.Length == 0)
            throw new CliError("bugreport 뒤에 제목 필요",
                "예: vbaloop bugreport \"push가 X를 깨뜨림\" / 목록은 vbaloop bugreport --list");

        // 보고자 = 프로젝트 폴더명 (toml이 있으면 그 위치 기준, 없으면 cwd 폴더명 폴백)
        var reporter = new DirectoryInfo(projectRoot).Name;

        var date = DateOnly.FromDateTime(DateTime.Now);
        var path = BugReport.ResolveReportPath(reportsDir, date, BugReport.Slugify(title));
        Directory.CreateDirectory(reportsDir);
        File.WriteAllText(path,
            BugReport.BuildTemplate(title, date, reporter, VersionCommand.GetBuildInfo()),
            CanonicalText.Utf8NoBom);

        o.HumanLines.Add($"생성됨: {path}");
        o.HumanLines.Add("다음: 파일을 열어 TODO/괄호 자리를 채울 것 (증거는 명령과 --json 출력 원문)");
        string? issueUrl = upstream ? null : BugReport.NewIssueUrl(title);
        if (issueUrl != null)
        {
            o.HumanLines.Add("작성 후 상류에 제보: 본문을 그대로 붙여넣으면 된다");
            o.HumanLines.Add($"  {issueUrl}");
        }
        o.Payload = new
        {
            path, reporter, upstream, issueUrl,
            date = date.ToString("yyyy-MM-dd"),
        };
        return o;
    }
}
