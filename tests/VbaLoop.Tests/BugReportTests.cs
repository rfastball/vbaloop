using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class BugReportTests
{
    private sealed class TempDir : IDisposable
    {
        public string Path { get; } =
            System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        public TempDir() => Directory.CreateDirectory(Path);
        public void Dispose() { try { Directory.Delete(Path, recursive: true); } catch { } }
        public string File(string name, string content = "")
        {
            var p = System.IO.Path.Combine(Path, name);
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(p)!);
            System.IO.File.WriteAllText(p, content);
            return p;
        }
    }

    [Theory]
    [InlineData("push breaks X", "push-breaks-x")]
    [InlineData("Pull이 식별자를 재케이싱함", "pull이-식별자를-재케이싱함")]
    [InlineData("push: read-only? (split/brain!)", "push-read-only-split-brain")]
    [InlineData("---", "report")]
    [InlineData("", "report")]
    public void Slugify_Cases(string title, string expected)
    {
        Assert.Equal(expected, BugReport.Slugify(title));
    }

    [Fact]
    public void Slugify_CapsAt60_NoTrailingHyphen()
    {
        var slug = BugReport.Slugify(string.Join(" ", Enumerable.Repeat("word", 30)));
        Assert.True(slug.Length <= 60);
        Assert.False(slug.EndsWith('-'));
        Assert.StartsWith("word-word", slug);
    }

    [Fact]
    public void ResolveReportPath_CollisionSuffixes()
    {
        using var t = new TempDir();
        var date = new DateOnly(2026, 7, 9);
        var first = BugReport.ResolveReportPath(t.Path, date, "my-bug");
        Assert.Equal(System.IO.Path.Combine(t.Path, "2026-07-09_my-bug.md"), first);

        t.File("2026-07-09_my-bug.md");
        var second = BugReport.ResolveReportPath(t.Path, date, "my-bug");
        Assert.Equal(System.IO.Path.Combine(t.Path, "2026-07-09_my-bug-2.md"), second);

        t.File("2026-07-09_my-bug-2.md");
        var third = BugReport.ResolveReportPath(t.Path, date, "my-bug");
        Assert.Equal(System.IO.Path.Combine(t.Path, "2026-07-09_my-bug-3.md"), third);
    }

    [Fact]
    public void BuildTemplate_ContainsRequiredLines()
    {
        var text = BugReport.BuildTemplate("pull이 X를 깨뜨림", new DateOnly(2026, 7, 9),
            "SampleRpa", "8acb552, built 2026-07-09 10:00:00");
        Assert.Contains("# [BUG] pull이 X를 깨뜨림", text);
        Assert.Contains("- 보고일: 2026-07-09", text);
        Assert.Contains("- 보고자: SampleRpa 세션 (Claude Code)", text);
        Assert.Contains("- 실행 exe 빌드: 8acb552, built 2026-07-09 10:00:00", text);
        Assert.Contains("## 요약", text);
        Assert.Contains("## 증거", text);
        Assert.Contains("## 제안 수정", text);
        Assert.Contains("## 재현 절차", text);
        Assert.EndsWith("\r\n", text); // 정규 텍스트 규약
        Assert.DoesNotContain("\n\n\n", text.Replace("\r", ""));
    }

    [Fact]
    public void ListReports_ResolvedIffHaesoSection()
    {
        using var t = new TempDir();
        t.File("2026-07-08_fixed-bug.md",
            "# [BUG] 고쳐진 버그\r\n\r\n## 요약\r\n\r\n본문\r\n\r\n## 해소 (2026-07-08)\r\n\r\n수정 완료\r\n");
        t.File("2026-07-09_open-bug.md",
            "# [BUG] 열린 버그\r\n\r\n## 요약\r\n\r\n본문\r\n");

        var reports = BugReport.ListReports(t.Path);
        Assert.Equal(2, reports.Count);
        Assert.Equal(("2026-07-08_fixed-bug.md", "[BUG] 고쳐진 버그", true), reports[0]);
        Assert.Equal(("2026-07-09_open-bug.md", "[BUG] 열린 버그", false), reports[1]);
    }

    [Fact]
    public void ListReports_MissingDir_Empty()
    {
        using var t = new TempDir();
        Assert.Empty(BugReport.ListReports(System.IO.Path.Combine(t.Path, "없는폴더")));
    }

    [Fact]
    public void NewIssueUrl_EncodesTitleIntoUpstreamRepo()
    {
        // 설치본 사용자의 유일한 제보 창구 — 제목에 공백/한글/특수문자가 있어도 링크가 깨지면 안 된다
        var url = BugReport.NewIssueUrl("push가 A/B를 깨뜨림 & 재현됨");
        Assert.StartsWith(BugReport.RepoUrl + "/issues/new?title=", url);
        Assert.DoesNotContain(" ", url);
        Assert.DoesNotContain("&", url[(url.IndexOf("title=") + 6)..]);  // 쿼리 인자 분리 오인 방지
    }
}
