using System.Text.Json;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class UpdateCheckTests
{
    private static readonly DateTimeOffset Now = new(2026, 10, 17, 12, 0, 0, TimeSpan.Zero);
    private static SemVer? V(string? s) => SemVer.ParseOrNull(s);

    // ---- 캐시 신선도 ----

    [Theory]
    [InlineData("2026-10-17T11:00:00.0000000+00:00", true)]   // 1시간 전
    [InlineData("2026-10-16T12:00:01.0000000+00:00", true)]   // 24시간 직전
    [InlineData("2026-10-16T12:00:00.0000000+00:00", false)]  // 정확히 24시간
    [InlineData("2026-10-10T00:00:00.0000000+00:00", false)]
    [InlineData("2026-10-17T12:03:00.0000000+00:00", true)]   // 약간의 시계 차이 허용
    [InlineData("2026-10-18T12:00:00.0000000+00:00", false)]  // 미래 = 시계 되감김 -> stale
    [InlineData("not a date", false)]
    [InlineData(null, false)]
    public void IsFresh(string? checkedAt, bool fresh)
        => Assert.Equal(fresh, UpdateCheck.IsFresh(new UpdateCheck.Cache { CheckedAt = checkedAt }, Now));

    [Fact]
    public void IsFresh_NoCache_IsStale() => Assert.False(UpdateCheck.IsFresh(null, Now));

    [Fact]
    public void Cache_RoundTrips_AndCorruptReadsAsNull()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        try
        {
            var path = Path.Combine(dir, "update.json");
            Assert.Null(UpdateCheck.Load(path));
            UpdateCheck.Save(path, new UpdateCheck.Cache { CheckedAt = Now.ToString("o"), Latest = "0.4.0" });
            var loaded = UpdateCheck.Load(path)!;
            Assert.Equal("0.4.0", loaded.Latest);
            Assert.True(UpdateCheck.IsFresh(loaded, Now));
            using (var doc = JsonDocument.Parse(File.ReadAllText(path)))
            {
                Assert.True(doc.RootElement.TryGetProperty("checkedAt", out _));  // 계약된 키 이름
                Assert.True(doc.RootElement.TryGetProperty("latest", out _));
            }

            File.WriteAllText(path, "{ broken");
            Assert.Null(UpdateCheck.Load(path));
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }

    // ---- 건너뛰기 ----

    [Theory]
    [InlineData(false, null, null, null)]
    [InlineData(true, null, null, "repo-build")]
    [InlineData(false, "1", null, UpdateCheck.OptOutEnv)]
    [InlineData(false, "true", null, UpdateCheck.OptOutEnv)]
    [InlineData(false, "0", null, null)]
    [InlineData(false, null, "true", "ci")]
    [InlineData(true, "1", "true", UpdateCheck.OptOutEnv)]
    public void SkipReason(bool repoBuild, string? optOut, string? ci, string? expected)
    {
        var env = new Dictionary<string, string?> { [UpdateCheck.OptOutEnv] = optOut, ["CI"] = ci };
        Assert.Equal(expected, UpdateCheck.SkipReason(repoBuild, k => env.GetValueOrDefault(k)));
    }

    [Fact]
    public void SkipReason_GitHubActions_IsCi()
        => Assert.Equal("ci", UpdateCheck.SkipReason(false, k => k == "GITHUB_ACTIONS" ? "true" : null));

    // ---- 알림 판정 ----

    [Theory]
    [InlineData("0.3.0", "0.4.0", true)]
    [InlineData("0.4.0", "0.4.0", false)]
    [InlineData("0.5.0", "0.4.0", false)]         // repo 빌드가 더 새로운 경우 등
    [InlineData("0.4.0", "0.5.0-rc.1", false)]    // 정식 사용자에게 prerelease 알리지 않음
    [InlineData("0.4.0-rc.1", "0.4.0", true)]     // prerelease 사용자에게 정식은 알림
    [InlineData("0.4.0-rc.1", "0.4.0-rc.2", true)]
    [InlineData("0.4.0-rc.2", "0.4.0-rc.1", false)]
    [InlineData(null, "0.4.0", false)]            // 현재 버전 판독 불가
    [InlineData("0.4.0", null, false)]            // 캐시에 latest 없음(오프라인)
    [InlineData("0.4.0", "garbage", false)]
    public void ShouldNotify(string? current, string? latest, bool notify)
        => Assert.Equal(notify, UpdateCheck.ShouldNotify(V(current), V(latest)));

    [Fact]
    public void NoticeFor_BuildsEnvelopeField()
    {
        var n = UpdateCheck.NoticeFor(V("0.3.0"), new UpdateCheck.Cache { Latest = "v0.4.0" });
        Assert.Equal(new UpdateNotice("0.3.0", "0.4.0", "vbaloop update"), n);
        Assert.Null(UpdateCheck.NoticeFor(V("0.3.0"), null));
    }

    // ---- 릴리스 판독 / 체크섬 ----

    private const string ReleaseJson = """
        {
          "tag_name": "v0.4.0",
          "draft": false,
          "prerelease": false,
          "assets": [
            { "name": "vbaloop.exe", "browser_download_url": "https://example.invalid/vbaloop.exe" },
            { "name": "checksums.txt", "browser_download_url": "https://example.invalid/checksums.txt" }
          ]
        }
        """;

    [Fact]
    public void ParseRelease_ReadsTagAndAssets()
    {
        var info = GitHubRelease.Parse(ReleaseJson)!;
        Assert.Equal("v0.4.0", info.Tag);
        Assert.Equal("0.4.0", info.Version);
        Assert.False(info.Prerelease);
        Assert.Equal("https://example.invalid/vbaloop.exe", info.Assets["vbaloop.exe"]);
        Assert.Equal(2, info.Assets.Count);
    }

    [Theory]
    [InlineData("""{"tag_name":"v0.5.0-rc.1","prerelease":false}""", true)] // 태그의 '-'도 prerelease
    [InlineData("""{"tag_name":"v0.5.0","prerelease":true}""", true)]
    public void ParseRelease_DetectsPrerelease(string json, bool pre)
        => Assert.Equal(pre, GitHubRelease.Parse(json)!.Prerelease);

    [Theory]
    [InlineData("""{"tag_name":"nightly"}""")]
    [InlineData("""{"tag_name":"v1.0.0","draft":true}""")]
    [InlineData("""{"message":"Not Found"}""")]
    [InlineData("not json")]
    public void ParseRelease_RejectsUnusable(string json) => Assert.Null(GitHubRelease.Parse(json));

    [Fact]
    public void ExpectedHash_ReadsSha256sumFormat()
    {
        var a = new string('a', 64);
        var b = new string('B', 64);
        var text = $"{a}  vbaloop-0.4.0-win-x64.zip\r\n{b} *vbaloop.exe\r\n";
        Assert.Equal(new string('b', 64), GitHubRelease.ExpectedHash(text, "vbaloop.exe"));
        Assert.Equal(a, GitHubRelease.ExpectedHash(text, "vbaloop-0.4.0-win-x64.zip"));
        Assert.Null(GitHubRelease.ExpectedHash(text, "other.exe"));
        Assert.Null(GitHubRelease.ExpectedHash($"{a}  vbaloop.exe.old\n", "vbaloop.exe")); // 접두 일치 금지
    }

    [Fact]
    public void Sha256Hex_MatchesKnownVector()
    {
        var path = Path.Combine(Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N") + ".bin");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        try
        {
            File.WriteAllText(path, "abc");
            Assert.Equal("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", GitHubRelease.Sha256Hex(path));
        }
        finally { File.Delete(path); }
    }

    // ---- 실행 중 exe 교체 ----

    [Fact]
    public void SelfReplace_MovesOldAside_InstallsNew_AndCleanupRemovesOnlyOld()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var target = Path.Combine(dir, "vbaloop.exe");
            var fresh = Path.Combine(dir, "new.bin");
            File.WriteAllText(target, "old");
            File.WriteAllText(fresh, "new");
            File.WriteAllText(Path.Combine(dir, "keep.old"), "x"); // vbaloop.exe*.old가 아니면 건드리지 않음

            SelfReplace.Replace(fresh, target);
            Assert.Equal("new", File.ReadAllText(target));
            Assert.Equal("old", File.ReadAllText(target + ".old"));

            // 이전 .old가 남아 있어도(지울 수 있으면 지우고) 다시 교체된다
            SelfReplace.Replace(fresh, target);
            Assert.Equal("new", File.ReadAllText(target + ".old"));

            Assert.Equal(1, SelfReplace.CleanupOld(dir));
            Assert.False(File.Exists(target + ".old"));
            Assert.True(File.Exists(Path.Combine(dir, "keep.old")));

            var missing = Path.Combine(dir, "sub", "vbaloop.exe");
            SelfReplace.Replace(fresh, missing);                      // 대상이 없으면 그냥 설치
            Assert.Equal("new", File.ReadAllText(missing));
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }

    [Fact]
    public void SelfReplace_LockedOldCopy_UsesUniqueName()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var target = Path.Combine(dir, "vbaloop.exe");
            var fresh = Path.Combine(dir, "new.bin");
            File.WriteAllText(target, "v2");
            File.WriteAllText(fresh, "v3");
            File.WriteAllText(target + ".old", "v1");
            using (File.Open(target + ".old", FileMode.Open, FileAccess.Read, FileShare.Read)) // 아직 실행 중인 v1 흉내
            {
                SelfReplace.Replace(fresh, target);
            }
            Assert.Equal("v3", File.ReadAllText(target));
            var olds = Directory.GetFiles(dir, "vbaloop.exe*.old").Select(Path.GetFileName).ToArray();
            Assert.Equal(2, olds.Length);
            Assert.Equal(2, SelfReplace.CleanupOld(dir));
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }
}
