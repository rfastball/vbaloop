using System.Text;

namespace VbaLoop.Core;

/// <summary>
/// bugreport(vbaloop 자체 결함 리포트 창구)의 순수 로직. 프로젝트 세션이 표준 템플릿을
/// vbaloop\docs\bugreports\{날짜}_{슬러그}.md로 남기는 진입점 — 형식은 기존 수기 리포트
/// 3건(2026-07-08/09)의 관행을 그대로 미러링한다. Excel/세션 불필요.
/// </summary>
public static class BugReport
{
    public const string ReportsRelDir = @"docs\bugreports";

    /// <summary>vbaloop 상류 저장소 — 설치본 사용자의 제보 창구.</summary>
    public const string RepoUrl = "https://github.com/rfastball/vbaloop";

    /// <summary>설치본(소스 레포 없음) 폴백 기록 위치(프로젝트 루트 기준).
    /// <c>_dbg</c>는 init이 관리하는 .gitignore로 무시되므로 실수로 커밋되지 않는다.</summary>
    public const string LocalReportsRelDir = @"_dbg\bugreports";

    /// <summary>제목이 채워진 GitHub 새 이슈 URL. 순수 함수 — xUnit 검증.</summary>
    public static string NewIssueUrl(string title)
        => $"{RepoUrl}/issues/new?title={Uri.EscapeDataString(title)}";
    private const int SlugMaxLength = 60;

    /// <summary>제목 -> 파일명 슬러그: 소문자화, 공백/구분자 -> '-', 영숫자·한글·'-' 외 제거,
    /// '-' 연쇄 축약, 양끝 trim, 60자 캡, 빈 결과 -> "report". 한글은 보존(Windows 파일명 무해).</summary>
    public static string Slugify(string title)
    {
        var sb = new StringBuilder(title.Length);
        foreach (var ch in title.ToLowerInvariant())
        {
            if (ch is >= 'a' and <= 'z' or >= '0' and <= '9' or >= '가' and <= '힣')
                sb.Append(ch);
            else
                sb.Append('-'); // 공백·문장부호·기타 전부 구분자로 — 아래에서 연쇄 축약
        }
        var collapsed = new StringBuilder(sb.Length);
        foreach (var ch in sb.ToString())
        {
            if (ch == '-' && (collapsed.Length == 0 || collapsed[^1] == '-')) continue;
            collapsed.Append(ch);
        }
        var slug = collapsed.ToString().TrimEnd('-');
        if (slug.Length > SlugMaxLength) slug = slug[..SlugMaxLength].TrimEnd('-');
        return slug.Length == 0 ? "report" : slug;
    }

    /// <summary>bugreportsDir\{yyyy-MM-dd}_{slug}.md, 이미 존재하면 -2, -3 ... 접미로 회피.</summary>
    public static string ResolveReportPath(string bugreportsDir, DateOnly date, string slug)
    {
        var baseName = $"{date:yyyy-MM-dd}_{slug}";
        var path = Path.Combine(bugreportsDir, baseName + ".md");
        for (int n = 2; File.Exists(path); n++)
            path = Path.Combine(bugreportsDir, $"{baseName}-{n}.md");
        return path;
    }

    /// <summary>표준 템플릿 생성. exeBuildInfo(리비전+빌드시각)를 자동 기입해
    /// "exe가 stale이었나" 판별(07-09 리포트의 수동 mtime 고고학)을 원천 대체한다.</summary>
    public static string BuildTemplate(string title, DateOnly date, string reporterProject, string exeBuildInfo)
    {
        var lines = new[]
        {
            $"# [BUG] {title}",
            "",
            $"- 보고일: {date:yyyy-MM-dd}",
            $"- 보고자: {reporterProject} 세션 (Claude Code) — (작업 맥락 한 줄로 교체)",
            "- 심각도: **TODO(Low/Medium/High)** — (판단 근거 한 줄)",
            $"- 실행 exe 빌드: {exeBuildInfo}",
            "",
            "## 요약",
            "",
            "(무엇이 어떻게 잘못되는지 2-4문장. 관련 기존 리포트가 있으면 파일명 인용)",
            "",
            "## 증거",
            "",
            "(실행한 명령과 --json 출력, git diff 등을 코드블록으로)",
            "",
            "## 제안 수정",
            "",
            "(선택 — 코드 포인터가 있으면 파일:라인으로)",
            "",
            "## 재현 절차",
            "",
            "1. (없으면 이 섹션 삭제)",
        };
        return CanonicalText.ToCanonical(lines);
    }

    /// <summary>--list용: 리포트 열거. Resolved = "## 해소"로 시작하는 라인 존재
    /// (기존 관행: 수정 완료 시 리포트에 "## 해소 (날짜)" 섹션 추기).</summary>
    public static List<(string File, string Title, bool Resolved)> ListReports(string bugreportsDir)
    {
        var result = new List<(string, string, bool)>();
        if (!Directory.Exists(bugreportsDir)) return result;
        foreach (var path in Directory.EnumerateFiles(bugreportsDir, "*.md", SearchOption.TopDirectoryOnly)
                     .OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
        {
            var lines = CanonicalText.SplitLines(CanonicalText.ReadFileStripBom(path));
            var title = lines.FirstOrDefault(l => l.StartsWith("# ", StringComparison.Ordinal))?[2..].Trim()
                        ?? Path.GetFileNameWithoutExtension(path);
            var resolved = lines.Any(l => l.StartsWith("## 해소", StringComparison.Ordinal));
            result.Add((Path.GetFileName(path), title, resolved));
        }
        return result;
    }
}
