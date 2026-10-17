using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 세션 메타파일 &lt;프로젝트루트&gt;\_dbg\session.json (설계 §4 공통 계약).
/// 세션 생존 판정 = (pid, excelStartTime) 일치 AND HWND attach 성공.
/// </summary>
public sealed class SessionMeta
{
    [JsonPropertyName("pid")] public int Pid { get; set; }
    /// <summary>Excel 프로세스 StartTime(UTC, "o" 포맷). PID 재사용 오판 방지.</summary>
    [JsonPropertyName("excelStartTime")] public string ExcelStartTime { get; set; } = "";
    [JsonPropertyName("hwnd")] public long Hwnd { get; set; }
    [JsonPropertyName("workbookFullName")] public string WorkbookFullName { get; set; } = "";
    [JsonPropertyName("createdAt")] public string CreatedAt { get; set; } = "";
    [JsonPropertyName("lastCommand")] public string? LastCommand { get; set; }
    /// <summary>
    /// 세션 소유: "tool"(도구가 띄운 Excel) | "external"(session attach로 채택한 사용자 Excel).
    /// 필드가 없는 구버전 메타는 tool. external은 어떤 경로로도 kill/Close/Quit하지 않는다(설계 §11).
    /// </summary>
    [JsonPropertyName("owner")] public string? Owner { get; set; }
    [JsonIgnore] public bool IsExternal => Owner == OwnerExternal;
    public const string OwnerTool = "tool";
    public const string OwnerExternal = "external";
    // push 해시 캐시/타임스탬프는 세션이 아니라 프로젝트 수명 — SyncCache(sync.json)로 분리됨.
    // 구버전 session.json의 잔여 lastPush* 필드는 역직렬화에서 무시된다.
    [JsonPropertyName("poisoned")] public bool Poisoned { get; set; }
    [JsonPropertyName("poisonReason")] public string? PoisonReason { get; set; }
    [JsonPropertyName("poisonedAt")] public string? PoisonedAt { get; set; }
    /// <summary>poison을 유발한 COM HRESULT(진단용).</summary>
    [JsonPropertyName("hResult")] public string? HResult { get; set; }

    private static readonly JsonSerializerOptions MetaJson = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static SessionMeta? Load(string path)
    {
        if (!File.Exists(path)) return null;
        try
        {
            return JsonSerializer.Deserialize<SessionMeta>(File.ReadAllText(path));
        }
        catch (JsonException)
        {
            return null; // 깨진 메타파일은 stale 취급 (session start가 정리)
        }
    }

    public void Save(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        // tmp는 프로세스 유니크 — 고정 이름이면 동시 Save 시 Move 경합(검토 지적)
        var tmp = $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(this, MetaJson), new System.Text.UTF8Encoding(false));
        File.Move(tmp, path, overwrite: true);
    }

    public static void Delete(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }

    public void MarkPoisoned(string path, string reason, int? hResult)
    {
        Poisoned = true;
        PoisonReason = reason;
        PoisonedAt = DateTime.UtcNow.ToString("o");
        HResult = hResult.HasValue ? $"0x{hResult.Value:X8}" : null;
        Save(path);
    }
}
