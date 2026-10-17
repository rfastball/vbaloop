using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 새 릴리스 알림(설계 §12.4). 명령은 캐시 파일(%LOCALAPPDATA%\vbaloop\update.json)만 읽는다 —
/// 네트워크는 24시간에 한 번, 분리된 자식(<c>vbaloop __update-check</c>)이 3초 제한으로 조회해 캐시를 쓴다.
/// 명령 자체에는 지연이 없다. 정확성(필수 갱신)은 관리 스탬프(§12.2)가 오프라인으로 맡고, 이건 편의다.
/// </summary>
public static class UpdateCheck
{
    public static readonly TimeSpan MaxAge = TimeSpan.FromHours(24);
    public static readonly TimeSpan QueryTimeout = TimeSpan.FromSeconds(3);
    public const string ChildVerb = "__update-check";
    public const string OptOutEnv = "VBALOOP_NO_UPDATE_CHECK";
    public const string UpdateCommand = "vbaloop update";

    public sealed class Cache
    {
        [JsonPropertyName("checkedAt")] public string? CheckedAt { get; set; }
        [JsonPropertyName("latest")] public string? Latest { get; set; }
        [JsonPropertyName("error")] public string? Error { get; set; }
    }

    private static readonly JsonSerializerOptions CacheJson = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string DefaultCachePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop", "update.json");

    public static Cache? Load(string path)
    {
        try
        {
            return File.Exists(path) ? JsonSerializer.Deserialize<Cache>(File.ReadAllText(path)) : null;
        }
        catch
        {
            return null; // 깨진 캐시 = 없음(다음 확인이 덮어쓴다)
        }
    }

    public static void Save(string path, Cache cache)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var tmp = $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(cache, CacheJson), new System.Text.UTF8Encoding(false));
        File.Move(tmp, path, overwrite: true);
    }

    /// <summary>캐시가 maxAge 안인가. checkedAt이 없거나 깨졌거나 미래(시계 되감김)면 stale.</summary>
    public static bool IsFresh(Cache? cache, DateTimeOffset now, TimeSpan? maxAge = null)
    {
        if (cache?.CheckedAt == null) return false;
        if (!DateTimeOffset.TryParse(cache.CheckedAt, null, System.Globalization.DateTimeStyles.RoundtripKind, out var at))
            return false;
        var age = now - at;
        return age >= TimeSpan.FromMinutes(-5) && age < (maxAge ?? MaxAge);
    }

    /// <summary>
    /// 확인을 건너뛸 이유(없으면 null). repo 빌드 exe(rollout.ps1이 갱신 경로), CI, 명시적 옵트아웃.
    /// env는 테스트 주입용.
    /// </summary>
    public static string? SkipReason(bool repoBuild, Func<string, string?>? env = null)
    {
        env ??= Environment.GetEnvironmentVariable;
        if (env(OptOutEnv) is "1" or "true") return OptOutEnv;
        if (!string.IsNullOrEmpty(env("CI")) || !string.IsNullOrEmpty(env("GITHUB_ACTIONS"))
            || !string.IsNullOrEmpty(env("TF_BUILD"))) return "ci";
        if (repoBuild) return "repo-build";
        return null;
    }

    /// <summary>
    /// 알림 판정: latest가 current보다 새롭고, 정식 사용자에게 prerelease를 알리지 않는다
    /// (releases/latest는 원래 prerelease를 주지 않지만 캐시를 손댄 경우까지 방어).
    /// current 판독 불가면 알리지 않는다.
    /// </summary>
    public static bool ShouldNotify(SemVer? current, SemVer? latest)
    {
        if (current == null || latest == null) return false;
        if (latest.IsPrerelease && !current.IsPrerelease) return false;
        return latest.CompareTo(current) > 0;
    }

    /// <summary>캐시 기준 알림(없으면 null). 순수 — 파일·네트워크 없음.</summary>
    public static UpdateNotice? NoticeFor(SemVer? current, Cache? cache)
    {
        var latest = SemVer.ParseOrNull(cache?.Latest);
        return ShouldNotify(current, latest)
            ? new UpdateNotice(current!.ToString(), latest!.ToString(), UpdateCommand)
            : null;
    }

    /// <summary>
    /// 알림 명령(status/doctor/version/init --check)의 진입점: 캐시를 읽어 알림을 돌려주고, stale이면
    /// 분리된 자식을 띄운다(결과는 다음 명령부터 보인다). 어떤 실패도 명령에 영향을 주지 않는다.
    /// </summary>
    public static UpdateNotice? Poll(bool repoBuild)
    {
        try
        {
            if (SkipReason(repoBuild) != null) return null;
            var path = DefaultCachePath;
            var cache = Load(path);
            if (!IsFresh(cache, DateTimeOffset.UtcNow) && Environment.ProcessPath is string exe)
                DetachedProcess.TryStart(exe, ChildVerb);
            return NoticeFor(ExeVersion.Current, cache);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 자식 측(<c>__update-check</c>): releases/latest를 3초 제한으로 조회해 캐시를 쓴다. 실패해도
    /// checkedAt을 기록한다 — 오프라인 PC에서 명령마다 자식이 뜨지 않게(이전 latest는 유지).
    /// 동시에 뜬 자식은 뮤텍스로 하나만 조회한다. 항상 exit 0, 출력 없음.
    /// </summary>
    public static int RunChild()
    {
        try
        {
            using var mx = new Mutex(false, @"Local\vbaloop-update-check");
            bool held;
            try { held = mx.WaitOne(0); }
            catch (AbandonedMutexException) { held = true; }
            if (!held) return 0;
            try
            {
                var path = DefaultCachePath;
                var prev = Load(path);
                if (IsFresh(prev, DateTimeOffset.UtcNow)) return 0; // 다른 자식이 방금 갱신
                var next = new Cache { CheckedAt = DateTimeOffset.UtcNow.ToString("o"), Latest = prev?.Latest };
                try
                {
                    using var http = GitHubRelease.CreateClient(QueryTimeout);
                    var info = GitHubRelease.FetchLatestAsync(http).GetAwaiter().GetResult();
                    if (info != null && !info.Prerelease) next.Latest = info.Version;
                }
                catch (Exception e)
                {
                    next.Error = e.GetType().Name;
                }
                Save(path, next);
            }
            finally
            {
                mx.ReleaseMutex();
            }
        }
        catch
        {
            // 편의 기능 — 조용히 끝낸다
        }
        return 0;
    }
}
