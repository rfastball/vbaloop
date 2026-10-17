using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 격리 마커 &lt;프로젝트루트&gt;\_dbg\recovery.json (설계 §11). 타임아웃·poison·확인 불가 실패에서
/// 도구가 Excel을 kill해도 안전하다고 증명하지 못했을 때, 프로젝트 mutex를 쥔 채로 기록한다.
/// 마커가 있는 동안 세션 필요 명령은 exit 3으로 거부되고, 해제는 recovery clear(기록된 프로세스
/// 소멸 확인) 또는 session stop --discard뿐이다. 기록된 프로세스가 이미 사라졌으면 게이트가 스스로 지운다.
/// </summary>
public sealed class RecoveryMarker
{
    [JsonPropertyName("pid")] public int Pid { get; set; }
    [JsonPropertyName("excelStartTime")] public string ExcelStartTime { get; set; } = "";
    /// <summary>SessionMeta.Owner와 같은 값("tool" | "external").</summary>
    [JsonPropertyName("owner")] public string Owner { get; set; } = SessionMeta.OwnerTool;
    /// <summary>timeout | poison | stale | session_died | stop 등 — 기계 판독용 짧은 분류.</summary>
    [JsonPropertyName("reason")] public string Reason { get; set; } = "";
    /// <summary>사람용 설명(예외 메시지, 프로브 결과).</summary>
    [JsonPropertyName("detail")] public string? Detail { get; set; }
    [JsonPropertyName("command")] public string? Command { get; set; }
    [JsonPropertyName("workbook")] public string? Workbook { get; set; }
    [JsonPropertyName("createdAt")] public string CreatedAt { get; set; } = "";

    /// <summary>파일은 있으나 읽을 수 없었음 — 격리는 유지(fail closed), 해제는 --force만.</summary>
    [JsonIgnore] public bool Corrupt { get; private set; }
    [JsonIgnore] public bool IsExternal => Owner == SessionMeta.OwnerExternal;

    private static readonly JsonSerializerOptions MarkerJson = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static RecoveryMarker? Load(string path)
    {
        if (!File.Exists(path)) return null;
        try
        {
            return JsonSerializer.Deserialize<RecoveryMarker>(File.ReadAllText(path))
                   ?? new RecoveryMarker { Corrupt = true, Reason = "corrupt" };
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            // 깨진 마커를 "없음"으로 읽으면 격리가 조용히 풀린다 — 존재 자체를 격리로 취급
            return new RecoveryMarker { Corrupt = true, Reason = "corrupt", Detail = ex.Message };
        }
    }

    public void Save(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var tmp = $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(this, MarkerJson), new System.Text.UTF8Encoding(false));
        File.Move(tmp, path, overwrite: true);
    }

    public static void Delete(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }

    public static RecoveryMarker For(SessionMeta meta, string reason, string? detail, string? command) => new()
    {
        Pid = meta.Pid,
        ExcelStartTime = meta.ExcelStartTime,
        Owner = meta.IsExternal ? SessionMeta.OwnerExternal : SessionMeta.OwnerTool,
        Reason = reason,
        Detail = detail,
        Command = command,
        Workbook = meta.WorkbookFullName,
        CreatedAt = DateTime.UtcNow.ToString("o"),
    };

    /// <summary>에러/상태 출력용 한 줄 요약.</summary>
    public string Describe()
        => Corrupt
            ? "quarantine marker corrupt (contents unreadable)"
            : $"Excel PID {Pid}" + (IsExternal ? "(user's Excel, attached)" : "(tool-launched)")
              + $" quarantined — reason {Reason}" + (Command != null ? $", command {Command}" : "")
              + (Detail != null ? $": {Detail}" : "");
}
