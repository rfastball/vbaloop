using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class GitignoreBlockTests
{
    private static readonly string[] Body = { "_dbg/", "dist/" };

    [Fact]
    public void Upsert_NullExisting_CreatesBlockOnly()
    {
        var (text, changed) = GitignoreBlock.Upsert(null, Body);
        Assert.True(changed);
        Assert.Equal(
            GitignoreBlock.StartMarker + "\r\n_dbg/\r\ndist/\r\n" + GitignoreBlock.EndMarker + "\r\n",
            text);
    }

    [Fact]
    public void Upsert_ExistingUserLines_AppendsBlockPreservingContent()
    {
        var (text, changed) = GitignoreBlock.Upsert("*.bak\r\nsecret.txt\r\n", Body);
        Assert.True(changed);
        Assert.Contains("*.bak", text);
        Assert.Contains("secret.txt", text);
        Assert.Contains(GitignoreBlock.StartMarker, text);
        Assert.Contains("dist/", text);
        // 사용자 라인이 마커보다 앞에 온다(EOF append)
        Assert.True(text.IndexOf("*.bak", System.StringComparison.Ordinal)
            < text.IndexOf(GitignoreBlock.StartMarker, System.StringComparison.Ordinal));
    }

    [Fact]
    public void Upsert_Idempotent_SecondRunUnchanged()
    {
        var (first, _) = GitignoreBlock.Upsert("node_modules/\r\n", Body);
        var (second, changed) = GitignoreBlock.Upsert(first, Body);
        Assert.False(changed);
        Assert.Equal(first, second);
    }

    [Fact]
    public void Upsert_ReplacesBetweenMarkers_WhenBlockDrifted()
    {
        string drifted = "keep/\r\n" + GitignoreBlock.StartMarker + "\r\nold_only/\r\n" + GitignoreBlock.EndMarker + "\r\n";
        var (text, changed) = GitignoreBlock.Upsert(drifted, Body);
        Assert.True(changed);
        Assert.Contains("keep/", text);      // 마커 밖 사용자 라인 보존
        Assert.DoesNotContain("old_only/", text); // 마커 안은 정본으로 교체
        Assert.Contains("dist/", text);
    }

    [Fact]
    public void Upsert_PreservesLfEol()
    {
        var (text, _) = GitignoreBlock.Upsert("foo/\n", Body);
        Assert.DoesNotContain("\r\n", text); // 순수 LF 파일은 LF 유지
        Assert.Contains("\n", text);
    }

    [Fact]
    public void Upsert_MissingEndMarker_RecoversFromCorruption()
    {
        // end 마커 유실(사용자 훼손) — start부터 EOF까지 잔해로 간주해 복구, 중복 블록 없이 수렴
        string corrupt = "a/\r\n" + GitignoreBlock.StartMarker + "\r\nstale/\r\n";
        var (text, _) = GitignoreBlock.Upsert(corrupt, Body);
        Assert.Contains("a/", text);
        Assert.DoesNotContain("stale/", text);
        Assert.Single(FindAll(text, GitignoreBlock.StartMarker)); // 마커 정확히 1개
        Assert.Contains(GitignoreBlock.EndMarker, text);
    }

    private static System.Collections.Generic.List<int> FindAll(string haystack, string needle)
    {
        var idxs = new System.Collections.Generic.List<int>();
        for (int i = haystack.IndexOf(needle, System.StringComparison.Ordinal); i >= 0;
             i = haystack.IndexOf(needle, i + needle.Length, System.StringComparison.Ordinal))
            idxs.Add(i);
        return idxs;
    }
}
