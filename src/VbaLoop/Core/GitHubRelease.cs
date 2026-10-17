using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>
/// GitHub 릴리스 조회와 체크섬 검증(설계 §12.4~12.5). install.ps1과 같은 계약이다:
/// releases/latest(prerelease·draft 제외)를 읽고, vbaloop.exe를 받아 같은 릴리스의 checksums.txt
/// SHA-256과 대조한다. 검증할 수 없으면 설치하지 않는다(fail closed).
/// </summary>
public static class GitHubRelease
{
    public const string Repo = "rfastball/vbaloop";
    public const string ExeAsset = "vbaloop.exe";
    public const string ChecksumsAsset = "checksums.txt";

    public static string LatestApiUrl => $"https://api.github.com/repos/{Repo}/releases/latest";

    public sealed record Info(string Tag, string Version, bool Prerelease, IReadOnlyDictionary<string, string> Assets);

    /// <summary>releases API JSON 판독. tag가 semver가 아니면 null(알림·갱신 대상 아님).</summary>
    public static Info? Parse(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var tag = root.TryGetProperty("tag_name", out var t) ? t.GetString() : null;
            var version = SemVer.ParseOrNull(tag);
            if (tag == null || version == null) return null;
            if (root.TryGetProperty("draft", out var d) && d.ValueKind == JsonValueKind.True) return null;
            bool pre = root.TryGetProperty("prerelease", out var p) && p.ValueKind == JsonValueKind.True;
            var assets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (root.TryGetProperty("assets", out var arr) && arr.ValueKind == JsonValueKind.Array)
                foreach (var a in arr.EnumerateArray())
                {
                    var name = a.TryGetProperty("name", out var n) ? n.GetString() : null;
                    var url = a.TryGetProperty("browser_download_url", out var u) ? u.GetString() : null;
                    if (name != null && url != null) assets[name] = url;
                }
            return new Info(tag, version.ToString(), pre || version.IsPrerelease, assets);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    /// <summary>checksums.txt("&lt;sha256&gt;  &lt;파일명&gt;", sha256sum 형식, '*' 바이너리 표식 허용)에서 fileName의 해시.</summary>
    public static string? ExpectedHash(string checksumsText, string fileName)
    {
        var re = new Regex(@"^\s*([0-9a-fA-F]{64})\s+\*?" + Regex.Escape(fileName) + @"\s*$");
        foreach (var line in checksumsText.Replace("\r\n", "\n").Split('\n'))
        {
            var m = re.Match(line);
            if (m.Success) return m.Groups[1].Value.ToLowerInvariant();
        }
        return null;
    }

    public static string Sha256Hex(string path)
    {
        using var fs = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(fs)).ToLowerInvariant();
    }

    public static HttpClient CreateClient(TimeSpan timeout)
    {
        var http = new HttpClient { Timeout = timeout };
        http.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("vbaloop", ExeVersion.Text == "unknown" ? "0" : ExeVersion.Text));
        http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        return http;
    }

    /// <summary>releases/latest 조회. 네트워크·HTTP 오류는 예외로 전파(호출자가 분류).</summary>
    public static async Task<Info?> FetchLatestAsync(HttpClient http, CancellationToken ct = default)
    {
        using var resp = await http.GetAsync(LatestApiUrl, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();
        return Parse(await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false));
    }
}
