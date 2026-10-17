using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>lease 소유자 레코드 — vh-core.ps1 project lease owner.json 스키마 이식.</summary>
public sealed class LeaseOwner
{
    [JsonPropertyName("schema")] public int Schema { get; set; } = 1;
    [JsonPropertyName("scope")] public string Scope { get; set; } = "project";
    [JsonPropertyName("leaseId")] public string LeaseId { get; set; } = Guid.NewGuid().ToString();
    /// <summary>소유자 신원 토큰: "env:..." | "proc:host:pid:startTime" | "local:host:user:key"</summary>
    [JsonPropertyName("ownerThread")] public string OwnerThread { get; set; } = "";
    [JsonPropertyName("ownerSource")] public string OwnerSource { get; set; } = "weak"; // env|process|weak
    [JsonPropertyName("agentPid")] public int? AgentPid { get; set; }
    [JsonPropertyName("agentStartTime")] public string? AgentStartTime { get; set; }
    [JsonPropertyName("host")] public string Host { get; set; } = "";
    [JsonPropertyName("user")] public string User { get; set; } = "";
    [JsonPropertyName("projectPath")] public string ProjectPath { get; set; } = "";
    [JsonPropertyName("workbookFullName")] public string WorkbookFullName { get; set; } = "";
    [JsonPropertyName("purpose")] public string Purpose { get; set; } = "agent-vba-edit";
    [JsonPropertyName("acquiredAt")] public string AcquiredAt { get; set; } = "";
    [JsonPropertyName("heartbeatAt")] public string HeartbeatAt { get; set; } = "";
    [JsonPropertyName("expiresAt")] public string ExpiresAt { get; set; } = "";
    [JsonPropertyName("lastCommand")] public string? LastCommand { get; set; }
    [JsonPropertyName("lastCommandPid")] public int LastCommandPid { get; set; }
}

/// <summary>
/// 프로젝트 lease 파일(&lt;루트&gt;\_dbg\lease.json)의 원자 조작 + 회수 결정 테이블.
/// vh-core.ps1 이식이되 두 가지 수정:
/// (1) claim이 "빈 파일 생성 후 내용 기록" 2단계가 아니라 **완전한 JSON tmp의 무-덮어쓰기
///     Move 1단계** — 부분 내용 관찰 창(owner null 오판 → 락 강탈)을 제거 (vh 원본의 결함 수정).
/// (2) process-bound owner에도 유휴 백스톱 추가 — 자동 획득 모델에서는 명시 release 의식이
///     없으므로, 장수 호스트가 "존재만" 해도 영구 점유하는 것을 expiresAt+grace로 회수 가능하게.
///     활동 중 에이전트는 매 게이트 명령이 renew하므로 절대 뺏기지 않는 속성은 유지된다.
/// </summary>
public static class LeaseFile
{
    /// <summary>TTL — 게이트 명령마다 heartbeat로 연장 (vh와 동일 값).</summary>
    public const int TtlSeconds = 1800;
    /// <summary>만료 후 이만큼 더 유휴여야 회수 가능 (vh와 동일 값).</summary>
    public const int ReclaimGraceSeconds = 1800;
    /// <summary>owner null/파싱 실패 회수의 최소 파일 연령 — 기록 직후 순간을 강탈로 오판 방지.</summary>
    public const int NullOwnerMinAgeSeconds = 10;

    private static readonly JsonSerializerOptions Json = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string PathFor(ProjectConfig cfg) => Path.Combine(cfg.DbgDir, "lease.json");

    public enum ClaimResult { Claimed, AlreadyHeld }

    /// <summary>원자적 배타 claim. 파일이 이미 있으면 AlreadyHeld.
    /// Directory.CreateDirectory는 멱등이라 뮤텍스가 못 되므로(검토 확정) 무-덮어쓰기 Move 사용.</summary>
    public static ClaimResult TryClaim(string path, LeaseOwner owner)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string tmp = TmpPath(path);
        File.WriteAllText(tmp, JsonSerializer.Serialize(owner, Json), CanonicalText.Utf8NoBom);

        for (int attempt = 0; ; attempt++)
        {
            try
            {
                File.Move(tmp, path); // 대상 존재 시 IOException — 이 성질이 상호배제의 전부
                return ClaimResult.Claimed;
            }
            catch (IOException) when (File.Exists(path))
            {
                TryDelete(tmp);
                return ClaimResult.AlreadyHeld;
            }
            catch (Exception ex) when (attempt < 5 && IsTransient(ex))
            {
                // delete-pending(직전 owner 삭제 후 AV/리더가 핸들 보유) / 공유 위반 — 잠시 후 재시도
                Thread.Sleep(100);
            }
        }
    }

    /// <summary>없거나 깨진 파일은 null — 회수 판정은 IsAutoReclaimable의 최소 연령 가드가 담당.</summary>
    public static LeaseOwner? Read(string path)
    {
        if (!File.Exists(path)) return null;
        try
        {
            return JsonSerializer.Deserialize<LeaseOwner>(File.ReadAllText(path));
        }
        catch (JsonException)
        {
            return null;
        }
        catch (IOException)
        {
            return null;
        }
    }

    /// <summary>보유 중 갱신 — in-place가 아니라 tmp+overwrite Move (torn read → owner null 오판 방지).</summary>
    public static void Save(string path, LeaseOwner owner)
    {
        string tmp = TmpPath(path);
        File.WriteAllText(tmp, JsonSerializer.Serialize(owner, Json), CanonicalText.Utf8NoBom);
        File.Move(tmp, path, overwrite: true);
    }

    public static void Delete(string path)
    {
        if (File.Exists(path)) File.Delete(path);
    }

    /// <summary>heartbeat/만료/명령 갱신 필드 채우기 (호출부가 Save).</summary>
    public static void Touch(LeaseOwner owner, string command, DateTime utcNow)
    {
        owner.HeartbeatAt = utcNow.ToString("o");
        owner.ExpiresAt = utcNow.AddSeconds(TtlSeconds).ToString("o");
        owner.LastCommand = command;
        owner.LastCommandPid = Environment.ProcessId;
    }

    /// <summary>
    /// 회수 결정 테이블 (vh Test-LeaseAutoReclaimable + 유휴 백스톱 수정).
    /// isProcessAlive/clock/파일 연령은 주입 — 단위 테스트 대상.
    /// </summary>
    public static bool IsAutoReclaimable(
        LeaseOwner? owner,
        DateTime utcNow,
        Func<int, string?, bool> isProcessAlive,
        string localHost,
        DateTime? fileLastWriteUtc)
    {
        // owner 없음/깨짐: 최소 연령 가드 — 방금 쓰인 파일(부분 기록/경합 순간)은 강탈하지 않음
        if (owner == null)
            return fileLastWriteUtc == null
                   || (utcNow - fileLastWriteUtc.Value).TotalSeconds >= NullOwnerMinAgeSeconds;

        // 다른 머신의 lease는 절대 자동 회수하지 않음 (break --force만 가능)
        if (!string.Equals(owner.Host, localHost, StringComparison.OrdinalIgnoreCase))
            return false;

        bool idleExpired = IsExpiredBy(owner.ExpiresAt, ReclaimGraceSeconds, utcNow);

        if (owner.OwnerSource == "process" && owner.AgentPid is int pid && pid > 0)
        {
            // 에이전트 프로세스 사망/PID 재사용 → 즉시 회수.
            // 살아있어도 유휴 백스톱(만료+grace) 경과면 회수 — 자동 획득 모델의 영구 점유 방지.
            if (!isProcessAlive(pid, owner.AgentStartTime)) return true;
            return idleExpired;
        }

        // env/weak(또는 pid 없는 process): 프로세스로 판정 불가 — 유휴 백스톱만
        return idleExpired;
    }

    /// <summary>소유 판정 (vh Test-IsProjectLeaseOwner): thread+host+user 3중 일치.</summary>
    public static bool IsOwnedBy(LeaseOwner owner, string ownerThread, string localHost, string localUser)
        => string.Equals(owner.OwnerThread, ownerThread, StringComparison.Ordinal)
           && string.Equals(owner.Host, localHost, StringComparison.OrdinalIgnoreCase)
           && string.Equals(owner.User, localUser, StringComparison.OrdinalIgnoreCase);

    /// <summary>유휴 시간(초) — 마지막 heartbeat 기준. 표시용.</summary>
    public static double IdleSeconds(LeaseOwner owner, DateTime utcNow)
    {
        if (DateTime.TryParse(owner.HeartbeatAt, null,
                System.Globalization.DateTimeStyles.RoundtripKind, out var hb))
            return Math.Max(0, (utcNow - hb.ToUniversalTime()).TotalSeconds);
        return double.NaN;
    }

    private static bool IsExpiredBy(string expiresAtIso, int graceSeconds, DateTime utcNow)
    {
        if (string.IsNullOrWhiteSpace(expiresAtIso)) return true; // 파싱 불가는 만료 취급(보수적)
        if (!DateTime.TryParse(expiresAtIso, null,
                System.Globalization.DateTimeStyles.RoundtripKind, out var exp))
            return true;
        return exp.ToUniversalTime().AddSeconds(graceSeconds) < utcNow;
    }

    private static bool IsTransient(Exception ex)
        => ex is UnauthorizedAccessException
           || (ex is IOException io && (io.HResult & 0xFFFF) == 0x20); // ERROR_SHARING_VIOLATION

    private static string TmpPath(string path)
        => $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";

    private static void TryDelete(string path)
    {
        try { File.Delete(path); } catch { /* tmp 잔재는 무해 */ }
    }
}
