using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 프로젝트 영속 동기화 캐시 &lt;프로젝트루트&gt;\_dbg\sync.json.
/// 세션 메타(session.json)는 세션 수명과 함께 사라지므로, push 해시 캐시가 세션 메타에
/// 있으면 session stop 후 첫 push가 항상 전체 재주입이 된다(SampleRpa 파일럿 실측 마찰).
/// push/pull 이력은 세션이 아니라 프로젝트에 속하므로 여기로 분리한다.
/// forms 맵은 폼 frx 해시 캐시(P5) — frx Export가 비결정적(3096바이트 중 13바이트 변동,
/// 실측)이라 fresh Export 비교 대신 "마지막 동기화된 src 상태"와 비교한다.
/// </summary>
public sealed class SyncCache
{
    [JsonPropertyName("lastPushFingerprint")] public string? LastPushFingerprint { get; set; }
    [JsonPropertyName("lastPushAt")] public string? LastPushAt { get; set; }
    /// <summary>마지막 pull 직후 src 지문 — pull 미푸시 편집 가드(PullGuard)의 기준선.</summary>
    [JsonPropertyName("lastPullFingerprint")] public string? LastPullFingerprint { get; set; }
    [JsonPropertyName("lastPullAt")] public string? LastPullAt { get; set; }
    [JsonPropertyName("forms")] public Dictionary<string, FormSync> Forms { get; set; } = NewFormsMap();

    public sealed class FormSync
    {
        /// <summary>마지막 동기화 시점 src .frm 파일 내용 해시(SHA-256 hex).</summary>
        [JsonPropertyName("frmHash")] public string? FrmHash { get; set; }
        /// <summary>마지막 동기화 시점 src .frx 파일 내용 해시(SHA-256 hex).</summary>
        [JsonPropertyName("frxHash")] public string? FrxHash { get; set; }
    }

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string PathFor(ProjectConfig cfg) => Path.Combine(cfg.DbgDir, "sync.json");

    /// <summary>없거나 깨진 파일은 빈 캐시로 — 캐시는 항상 재구축 가능해야 한다.</summary>
    public static SyncCache Load(string path)
    {
        if (!File.Exists(path)) return new SyncCache();
        try
        {
            var loaded = JsonSerializer.Deserialize<SyncCache>(File.ReadAllText(path));
            if (loaded == null) return new SyncCache();
            // 역직렬화된 Dictionary는 비교자가 기본값이므로 대소문자 무시로 재래핑
            loaded.Forms = new Dictionary<string, FormSync>(loaded.Forms, StringComparer.OrdinalIgnoreCase);
            return loaded;
        }
        catch (JsonException)
        {
            return new SyncCache();
        }
    }

    public void Save(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var tmp = $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(this, Json), CanonicalText.Utf8NoBom);
        File.Move(tmp, path, overwrite: true);
    }

    private static Dictionary<string, FormSync> NewFormsMap()
        => new(StringComparer.OrdinalIgnoreCase);
}
