using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 도구가 띄운/격리한 Excel PID의 사용자 단위 기록(%LOCALAPPDATA%\vbaloop\excel-pids.json, 설계 §11).
/// 프로젝트 마커는 그 프로젝트 안에서만 보이므로, doctor --kill-orphans가 프로젝트 밖에서도
/// "격리 PID는 절대 kill 안 함"과 "도구가 띄운 PID 우선"을 지키게 하는 보조 기록이다.
/// best-effort: 실패해도 명령은 계속된다. 소멸한 프로세스(신원 Gone) 항목은 기록 시 정리된다.
/// --force로 해제된 격리 PID도 프로세스가 사라질 때까지 quarantined로 남는다(안전 방향).
/// </summary>
public static class ExcelPidRegistry
{
    public const string KindLaunched = "launched";
    public const string KindQuarantined = "quarantined";
    private const int MaxEntries = 200;

    public sealed class Entry
    {
        [JsonPropertyName("pid")] public int Pid { get; set; }
        [JsonPropertyName("excelStartTime")] public string ExcelStartTime { get; set; } = "";
        [JsonPropertyName("kind")] public string Kind { get; set; } = KindLaunched;
        [JsonPropertyName("project")] public string? Project { get; set; }
        [JsonPropertyName("at")] public string At { get; set; } = "";
    }

    private sealed class Doc
    {
        [JsonPropertyName("entries")] public List<Entry> Entries { get; set; } = new();
    }

    private static readonly JsonSerializerOptions DocJson = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop", "excel-pids.json");

    public static List<Entry> Load(string path)
    {
        try
        {
            if (!File.Exists(path)) return new();
            return JsonSerializer.Deserialize<Doc>(File.ReadAllText(path))?.Entries ?? new();
        }
        catch
        {
            return new();
        }
    }

    /// <summary>
    /// (pid, startTime, kind) 기록. isGone은 테스트 주입용 — 기본은 ProcessIdentity로 소멸 판정.
    /// 프로세스 간 경합은 사용자 세션 네임드 뮤텍스로 직렬화한다.
    /// </summary>
    public static void TryRecord(string path, int pid, string startTimeIso, string kind, string? project,
        Func<Entry, bool>? isGone = null)
    {
        isGone ??= e => ProcessIdentity.Identify(e.Pid, e.ExcelStartTime) == IdentityState.Gone;
        try
        {
            using var mx = new Mutex(false, @"Local\vbaloop-excel-pid-registry");
            bool held;
            try { held = mx.WaitOne(2000); }
            catch (AbandonedMutexException) { held = true; }
            try
            {
                var entries = Load(path).Where(e => !isGone(e)).ToList();
                entries.RemoveAll(e => e.Pid == pid && e.ExcelStartTime == startTimeIso && e.Kind == kind);
                entries.Add(new Entry
                {
                    Pid = pid, ExcelStartTime = startTimeIso, Kind = kind, Project = project,
                    At = DateTime.UtcNow.ToString("o"),
                });
                if (entries.Count > MaxEntries) entries = entries.Skip(entries.Count - MaxEntries).ToList();

                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                var tmp = $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
                File.WriteAllText(tmp, JsonSerializer.Serialize(new Doc { Entries = entries }, DocJson),
                    new System.Text.UTF8Encoding(false));
                File.Move(tmp, path, overwrite: true);
            }
            finally
            {
                if (held) mx.ReleaseMutex();
            }
        }
        catch
        {
            // 보조 기록 — 실패가 명령을 깨면 안 된다
        }
    }

    /// <summary>기록 중 (pid, startTime)이 일치하고 kind가 같은 항목이 있는지.</summary>
    public static bool Has(IEnumerable<Entry> entries, int pid, string? startTimeIso, string kind)
        => startTimeIso != null
           && entries.Any(e => e.Pid == pid && e.ExcelStartTime == startTimeIso && e.Kind == kind);
}
