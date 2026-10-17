using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>
/// 워크북 백업 저장소 &lt;프로젝트루트&gt;\_dbg\backups\&lt;id&gt;\ (설계 §4 push 0.5단계 / rollback).
/// id = <c>yyyyMMdd-HHmmss-fff-&lt;kind&gt;</c>(UTC) — 사전순 = 시간순이라 정렬에 메타 파싱이 필요 없다.
/// 슬롯 하나 = 워크북 사본 1개 + metadata.json. metadata.json은 사본 기록이 끝난 뒤 마지막에
/// 원자 기록되므로, metadata 없는 슬롯은 "미완(중단된 백업)"이다 — 복원 대상이 될 수 없고 prune만 된다.
/// 이 클래스는 Excel/COM을 모른다: 사본 기록(SaveCopyAs/File.Copy)은 호출자 몫.
/// </summary>
public static class BackupStore
{
    public const string KindPush = "push";
    public const string KindPreRollback = "pre-rollback";
    public const string MetadataFileName = "metadata.json";

    private const string StampFormat = "yyyyMMdd-HHmmss-fff";
    private static readonly Regex IdPattern = new(@"^(\d{8}-\d{6}-\d{3})-([a-z][a-z-]*)$", RegexOptions.CultureInvariant);

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public sealed class Metadata
    {
        [JsonPropertyName("id")] public string Id { get; set; } = "";
        [JsonPropertyName("kind")] public string Kind { get; set; } = "";
        /// <summary>백업을 만든 명령(push / rollback).</summary>
        [JsonPropertyName("command")] public string Command { get; set; } = "";
        [JsonPropertyName("createdAtUtc")] public string CreatedAtUtc { get; set; } = "";
        /// <summary>백업 시점의 정본 워크북 절대 경로.</summary>
        [JsonPropertyName("workbookPath")] public string WorkbookPath { get; set; } = "";
        /// <summary>슬롯 안의 사본 파일명(= 워크북 파일명).</summary>
        [JsonPropertyName("workbookFile")] public string WorkbookFile { get; set; } = "";
        /// <summary>백업 시점 src 지문(push라면 이번에 주입하려던 src).</summary>
        [JsonPropertyName("srcFingerprint")] public string? SrcFingerprint { get; set; }
        /// <summary>백업 시점 sync.json의 lastPushFingerprint — 사본 워크북이 마지막으로 동기화된 src.</summary>
        [JsonPropertyName("lastPushFingerprint")] public string? LastPushFingerprint { get; set; }
        [JsonPropertyName("exeVersion")] public string? ExeVersion { get; set; }
        [JsonPropertyName("sizeBytes")] public long SizeBytes { get; set; }
        /// <summary>사본의 SHA-256 hex — rollback이 복원 전후로 무결성을 대조한다.</summary>
        [JsonPropertyName("sha256")] public string Sha256 { get; set; } = "";
        /// <summary>pre-rollback 전용: 이 안전 백업 직후 복원된 백업 id.</summary>
        [JsonPropertyName("restoringFrom")] public string? RestoringFrom { get; set; }
    }

    /// <summary>열거 결과 한 건. Meta == null이면 미완 슬롯.</summary>
    public sealed record Entry(string Id, string Dir, DateTime CreatedUtc, string Kind, Metadata? Meta)
    {
        public bool Complete => Meta != null;
        public string? WorkbookCopyPath => Meta == null ? null : Path.Combine(Dir, Meta.WorkbookFile);
    }

    /// <summary>새 슬롯 — 디렉터리는 이미 생성됨. 호출자가 CopyPath에 사본을 쓴 뒤 <see cref="Complete"/>.</summary>
    public sealed record Slot(string Id, string Dir, string Kind, DateTime CreatedUtc, string CopyPath);

    public static string MakeId(DateTime utc, string kind)
        => utc.ToString(StampFormat, CultureInfo.InvariantCulture) + "-" + kind;

    public static bool TryParseId(string id, out DateTime utc, out string kind)
    {
        utc = default;
        kind = "";
        var m = IdPattern.Match(id);
        if (!m.Success) return false;
        if (!DateTime.TryParseExact(m.Groups[1].Value, StampFormat, CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out utc))
            return false;
        kind = m.Groups[2].Value;
        return true;
    }

    /// <summary>
    /// 슬롯 할당: id 충돌(같은 ms에 두 번) 시 1ms씩 밀어 유일 디렉터리를 만든다. 명령 mutex 하에서
    /// 호출되므로 경합은 사실상 없지만 CreateDirectory는 기존 디렉터리에도 성공하므로 존재 검사로 막는다.
    /// </summary>
    public static Slot Allocate(string root, string kind, string workbookPath, DateTime utcNow)
    {
        Directory.CreateDirectory(root);
        var t = utcNow;
        for (int i = 0; i < 1000; i++, t = t.AddMilliseconds(1))
        {
            string id = MakeId(t, kind);
            string dir = Path.Combine(root, id);
            if (Directory.Exists(dir) || File.Exists(dir)) continue;
            Directory.CreateDirectory(dir);
            return new Slot(id, dir, kind, t, Path.Combine(dir, Path.GetFileName(workbookPath)));
        }
        throw new CliError($"백업 슬롯 할당 실패(id 충돌 지속): {root}");
    }

    /// <summary>사본이 기록된 슬롯을 완결: 크기/해시를 채워 metadata.json을 원자 기록.</summary>
    public static Metadata Complete(Slot slot, Metadata partial)
    {
        if (!File.Exists(slot.CopyPath))
            throw new CliError($"백업 사본이 기록되지 않음: {slot.CopyPath}");
        partial.Id = slot.Id;
        partial.Kind = slot.Kind;
        partial.CreatedAtUtc = slot.CreatedUtc.ToString("o", CultureInfo.InvariantCulture);
        partial.WorkbookFile = Path.GetFileName(slot.CopyPath);
        partial.SizeBytes = new FileInfo(slot.CopyPath).Length;
        partial.Sha256 = HashFile(slot.CopyPath);

        string path = Path.Combine(slot.Dir, MetadataFileName);
        string tmp = $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(partial, Json), CanonicalText.Utf8NoBom);
        File.Move(tmp, path, overwrite: true);
        return partial;
    }

    /// <summary>실패한 슬롯 정리(최선 노력) — 미완 슬롯은 어차피 prune 대상이지만 즉시 치운다.</summary>
    public static void Discard(Slot slot)
    {
        try { if (Directory.Exists(slot.Dir)) Directory.Delete(slot.Dir, recursive: true); } catch { /* prune가 회수 */ }
    }

    /// <summary>
    /// 백업 열거(최신순). id 패턴이 아닌 디렉터리는 무시한다 — 사용자가 둔 것일 수 있으므로 prune도 안 한다.
    /// metadata가 없거나 깨졌거나 사본 파일이 없으면 미완(Meta=null)으로 취급.
    /// </summary>
    public static List<Entry> List(string root)
    {
        var list = new List<Entry>();
        if (!Directory.Exists(root)) return list;
        foreach (var dir in Directory.EnumerateDirectories(root))
        {
            string id = Path.GetFileName(dir);
            if (!TryParseId(id, out var utc, out var kind)) continue;
            list.Add(new Entry(id, dir, utc, kind, TryLoadMetadata(dir, id)));
        }
        list.Sort((a, b) => string.CompareOrdinal(b.Id, a.Id));
        return list;
    }

    private static Metadata? TryLoadMetadata(string dir, string id)
    {
        string path = Path.Combine(dir, MetadataFileName);
        if (!File.Exists(path)) return null;
        try
        {
            var m = JsonSerializer.Deserialize<Metadata>(File.ReadAllText(path));
            if (m == null || m.Id != id || string.IsNullOrEmpty(m.WorkbookFile)) return null;
            // 사본 파일명은 슬롯 안의 단일 파일명이어야 한다(메타 변조로 슬롯 밖을 가리키는 것 차단)
            if (Path.GetFileName(m.WorkbookFile) != m.WorkbookFile) return null;
            return File.Exists(Path.Combine(dir, m.WorkbookFile)) ? m : null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    /// <summary>
    /// rollback 대상 선택(순수): --latest면 최신 완결 백업(종류 무관 — 직전 rollback의 pre-rollback
    /// 백업이 최신이면 그것이 곧 "rollback 되돌리기"다), --backup id면 정확 일치 완결 백업.
    /// </summary>
    public static Entry Select(IReadOnlyList<Entry> newestFirst, bool latest, string? id)
    {
        if (latest == (id != null))
            throw new CliError("rollback 대상은 --latest 또는 --backup <id> 중 하나만 지정",
                "vbaloop rollback --list로 백업 목록 확인");

        if (latest)
            return newestFirst.FirstOrDefault(e => e.Complete)
                ?? throw new CliError("복원 가능한 백업이 없음",
                    "push가 백업을 만든다([backup] keep = 0이면 비활성) — vbaloop rollback --list로 확인");

        var hit = newestFirst.FirstOrDefault(e => e.Id == id);
        if (hit == null)
        {
            var known = newestFirst.Where(e => e.Complete).Take(5).Select(e => e.Id).ToList();
            throw new CliError($"백업 id 없음: {id}",
                known.Count > 0 ? $"최근 백업: {string.Join(", ", known)} (전체는 rollback --list)" : "백업이 하나도 없음");
        }
        if (!hit.Complete)
            throw new CliError($"백업 {id}는 미완(metadata.json 또는 사본 없음) — 복원 불가",
                "다른 백업을 선택 (rollback --list)");
        return hit;
    }

    /// <summary>
    /// 보존 정책(순수): 보호 id를 제외한 완결 백업 중 최신 keep개만 남기고, 미완 슬롯은 보호 대상이 아니면
    /// 전부 삭제 대상. keep &lt;= 0이면 아무것도 지우지 않는다(백업 비활성은 "기존 백업 파기"가 아님).
    /// 반환: 삭제할 id 목록(오래된 순서 무관).
    /// </summary>
    public static List<string> PlanPrune(IReadOnlyList<Entry> newestFirst, int keep, IEnumerable<string>? protect = null,
        bool countProtected = false)
    {
        var del = new List<string>();
        if (keep <= 0) return del;
        var guard = new HashSet<string>(protect ?? Array.Empty<string>(), StringComparer.Ordinal);
        int kept = 0;
        foreach (var e in newestFirst.OrderByDescending(e => e.Id, StringComparer.Ordinal))
        {
            if (guard.Contains(e.Id)) { if (countProtected && e.Complete) kept++; continue; }
            if (!e.Complete) { del.Add(e.Id); continue; }
            if (kept < keep) { kept++; continue; }
            del.Add(e.Id);
        }
        return del;
    }

    /// <summary>보존 정책 적용. 개별 삭제 실패(잠금 등)는 비치명 — 실패 목록을 돌려준다.</summary>
    public static (List<string> Deleted, List<string> Failed) Prune(string root, int keep, IEnumerable<string>? protect = null,
        bool countProtected = false)
    {
        var deleted = new List<string>();
        var failed = new List<string>();
        foreach (var id in PlanPrune(List(root), keep, protect, countProtected))
        {
            try
            {
                Directory.Delete(Path.Combine(root, id), recursive: true);
                deleted.Add(id);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                failed.Add(id);
            }
        }
        return (deleted, failed);
    }

    public static string HashFile(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert.ToHexString(SHA256.HashData(fs));
    }
}
