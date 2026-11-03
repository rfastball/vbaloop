using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 협조적 취소 표시 파일(설계 §14.4). LogShim 판독 계약과 같은 폴더(&lt;워크북.Path&gt;\_dbg)에 실행마다
/// 고유한 이름으로 둔다 — 이전 실행의 잔재가 다음 실행을 멈추지 않도록. LogShim은 BeginRun으로 받은
/// runId의 파일만 본다.
/// </summary>
public static class RunCancel
{
    public static string PathFor(string workbookFullName, string runId)
        => Path.Combine(Path.GetDirectoryName(workbookFullName)!, "_dbg", "cancel." + runId);

    /// <summary>취소를 요청한다(멱등). 실패해도 예외를 던지지 않는다 — 호출자는 유예 뒤 강제 경로로 간다.</summary>
    public static bool Request(string flagPath)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(flagPath)!);
            File.WriteAllText(flagPath, DateTime.UtcNow.ToString("O"));
            return true;
        }
        catch
        {
            return false;
        }
    }

    public static void Clear(string flagPath)
    {
        try { if (File.Exists(flagPath)) File.Delete(flagPath); } catch { /* 다음 실행이 오래된 잔재를 지운다 */ }
    }

    /// <summary>
    /// 세션을 놓을 때(격리 해제, session stop, 격리 자동 해제) 남은 표시 파일을 모두 지운다. 강제 경로가 남긴 표시와
    /// LogShim에 남은 runId가 겹치면, 사용자가 계속 쓰는 Excel의 매크로가 LogShim 호출에서 End될 수 있다.
    /// 실행 기록이 있으면(실행 중이거나 고아 취소 대기 중) 건드리지 않는다.
    /// </summary>
    public static void ClearAll(ProjectConfig cfg)
    {
        if (File.Exists(RunLock.RecordPath(cfg))) return;
        foreach (string dir in new[] { cfg.DbgDir, Path.Combine(Path.GetDirectoryName(cfg.WorkbookPath) ?? cfg.RootDir, "_dbg") }.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                if (!Directory.Exists(dir)) continue;
                foreach (var f in Directory.EnumerateFiles(dir, "cancel.*"))
                    File.Delete(f);
            }
            catch { /* 정리 실패 - 하루 뒤 SweepOld */ }
        }
    }

    /// <summary>offset 이후 구간에 이 runId의 LogShim 취소 레코드가 있는가. 파일이 offset보다 짧으면(LogShim.Reset이
    /// 잘라냈다) 처음부터 본다 - runId가 실행마다 고유하므로 안전하다.</summary>
    public static bool SpanCancelled(string jsonlPath, long offset, string runId)
    {
        if (!File.Exists(jsonlPath)) return false;
        using var fs = new FileStream(jsonlPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        if (offset > fs.Length) offset = 0;
        if (offset == fs.Length) return false;
        fs.Seek(offset, SeekOrigin.Begin);
        using var reader = new StreamReader(fs, new UTF8Encoding(false));
        string? line;
        while ((line = reader.ReadLine()) != null)
            if (IsCancelRecord(line, runId)) return true;
        return false;
    }

    public static bool IsCancelRecord(string line, string runId)
    {
        line = line.Trim();
        if (!line.StartsWith('{') || !line.Contains("\"cancelled\"", StringComparison.Ordinal)) return false;
        try
        {
            using var doc = JsonDocument.Parse(line);
            var r = doc.RootElement;
            return r.TryGetProperty("kind", out var k) && k.GetString() == "X"
                && r.TryGetProperty("event", out var e) && e.GetString() == "cancelled"
                && r.TryGetProperty("runId", out var id) && id.GetString() == runId;
        }
        catch (JsonException)
        {
            return false;
        }
    }

    /// <summary>하루 넘은 잔재 표시 파일을 지운다(runId가 달라 무해하지만 쌓이지 않게).</summary>
    public static void SweepOld(string dbgDir, TimeSpan age)
    {
        try
        {
            if (!Directory.Exists(dbgDir)) return;
            foreach (var f in Directory.EnumerateFiles(dbgDir, "cancel.*"))
                if (DateTime.UtcNow - File.GetLastWriteTimeUtc(f) > age)
                    File.Delete(f);
        }
        catch { /* 정리 실패는 무해 */ }
    }
}

/// <summary>실행 중인 run/test의 기록(<c>_dbg\run-active.json</c>). 러너가 쓰고 지운다.</summary>
public sealed class ActiveRun
{
    [JsonPropertyName("runId")] public string RunId { get; set; } = "";
    [JsonPropertyName("command")] public string Command { get; set; } = "";
    [JsonPropertyName("target")] public string Target { get; set; } = "";
    [JsonPropertyName("runnerPid")] public int RunnerPid { get; set; }
    [JsonPropertyName("runnerStart")] public string RunnerStart { get; set; } = "";
    [JsonPropertyName("excelPid")] public int ExcelPid { get; set; }
    [JsonPropertyName("startedUtc")] public string StartedUtc { get; set; } = "";
    [JsonPropertyName("cancelFlag")] public string CancelFlag { get; set; } = "";
    /// <summary>LogShim 기록 파일 - 고아 취소가 매크로의 멈춤(취소 기록)을 확인하는 곳.</summary>
    [JsonPropertyName("jsonl")] public string Jsonl { get; set; } = "";
    /// <summary>러너 없이 남은 기록에 다음 명령이 취소를 요청한 시각(고아 취소 대기의 기준).</summary>
    [JsonPropertyName("orphanCancelUtc")] public string? OrphanCancelUtc { get; set; }
    /// <summary>프로젝트 LogShim이 BeginRun을 받았는가(null = 아직 모름). false면 협조적 취소가 불가능하니
    /// 부모는 유예 없이 강제 경로로 간다(BeginRun이 없는 구버전 LogShim).</summary>
    [JsonPropertyName("cancelSupported")] public bool? CancelSupported { get; set; }

    public string Describe(DateTime nowUtc)
    {
        string elapsed = DateTime.TryParse(StartedUtc, null, System.Globalization.DateTimeStyles.RoundtripKind, out var t)
            ? RunProgress.FormatElapsed(nowUtc - t) : "?";
        return $"{Command} {Target} (runId {RunId}, runner PID {RunnerPid}, Excel PID {ExcelPid}, {elapsed} elapsed)";
    }
}

/// <summary>
/// 러너가 실행 동안 쥐는 프로젝트 실행 잠금(설계 §14.6). 명령 잠금(<see cref="CommandMutex"/>)은 부모만 쥐어
/// 부모가 강제 종료되면 풀리지만 매크로는 계속 돈다 — 러너가 이 잠금을 쥐고, 다른 명령은 이것이 잡혀 있으면
/// 거부한다. 커널 뮤텍스라 러너가 죽으면 자동 해제된다(Abandoned = 비어 있음).
/// </summary>
public sealed class RunLock : IDisposable
{
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    /// <summary>부모·러너가 함께 죽은 실행(트리 kill)의 매크로가 멈추기를 다음 명령이 기다리는 상한.</summary>
    public const int OrphanWaitMs = 30_000;

    private readonly Mutex _mutex;
    private readonly string _recordPath;
    private readonly string _runId;

    private RunLock(Mutex mutex, string recordPath, string runId)
    {
        _mutex = mutex;
        _recordPath = recordPath;
        _runId = runId;
    }

    public static string RecordPath(ProjectConfig cfg) => Path.Combine(cfg.DbgDir, "run-active.json");
    private static string MutexName(ProjectConfig cfg) => @"Local\vbaloop-run-" + CommandMutex.KeyFor(cfg.RootDir);

    /// <summary>러너 측: 잠금을 쥐고 기록을 남긴다. 잡을 수 없으면 null(부모가 명령 잠금으로 이미 직렬화하므로
    /// 정상 경로에선 일어나지 않는다 — 잡지 못해도 실행은 계속한다).</summary>
    public static RunLock? Hold(ProjectConfig cfg, ActiveRun record)
    {
        var mutex = new Mutex(initiallyOwned: false, MutexName(cfg));
        bool held;
        try { held = mutex.WaitOne(0); }
        catch (AbandonedMutexException) { held = true; }
        if (!held)
        {
            mutex.Dispose();
            return null;
        }
        string path = RecordPath(cfg);
        try
        {
            Directory.CreateDirectory(cfg.DbgDir);
            File.WriteAllText(path, JsonSerializer.Serialize(record, Json));
        }
        catch { /* 기록 실패 — 잠금만으로도 거부는 동작한다 */ }
        return new RunLock(mutex, path, record.RunId);
    }

    public void Update(ActiveRun record)
    {
        try { File.WriteAllText(_recordPath, JsonSerializer.Serialize(record, Json)); }
        catch { /* 기록 실패 - 부모는 유예 뒤 강제 경로로 간다 */ }
    }

    public void Dispose()
    {
        try
        {
            if (Load(_recordPath) is { } r && r.RunId == _runId) File.Delete(_recordPath);
        }
        catch { /* 다음 명령이 잔재로 정리 */ }
        try { _mutex.ReleaseMutex(); } catch (ApplicationException) { /* 소유 아님 */ }
        _mutex.Dispose();
    }

    /// <summary>부모 측: 러너를 kill한 뒤 그 실행의 기록을 지운다(다른 실행의 기록이면 둔다). 취소 표시는 남긴다 —
    /// 격리된 Excel 안에서 아직 도는 매크로가 다음 LogShim 호출에서 멈추도록.</summary>
    public static void Forget(ProjectConfig cfg, string runId)
    {
        string path = RecordPath(cfg);
        try
        {
            if (Load(path) is { } r && r.RunId == runId) File.Delete(path);
        }
        catch { /* 다음 명령이 잔재로 정리 */ }
    }

    public static ActiveRun? Load(string path)
    {
        try
        {
            return File.Exists(path) ? JsonSerializer.Deserialize<ActiveRun>(File.ReadAllText(path)) : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 명령 측(명령 잠금을 쥔 뒤): 다른 러너가 실행 중이면 거부한다. 러너가 이미 죽었는데 기록이 남았으면
    /// (부모·러너 트리 kill) 그 실행의 매크로는 Excel 안에서 계속 돌 수 있으므로 취소 표시를 남기고 기록을 지운다 —
    /// 매크로는 다음 LogShim 확인 지점에서 멈춘다.
    /// </summary>
    public static void ThrowIfActive(ProjectConfig cfg)
    {
        if (Mutex.TryOpenExisting(MutexName(cfg), out var existing))
        {
            using (existing)
            {
                bool free;
                try { free = existing.WaitOne(0); }
                catch (AbandonedMutexException) { free = true; }
                if (free)
                {
                    existing.ReleaseMutex();
                }
                else
                {
                    var rec = Load(RecordPath(cfg));
                    throw new CliError(
                        "A vbaloop run is still executing a macro in this project"
                        + (rec != null ? ": " + rec.Describe(DateTime.UtcNow) : ""),
                        "Wait for it to finish, or stop the command that started it (Ctrl+C / stop the task) "
                        + "- it then cancels the macro at its next LogShim call",
                        environment: true);
                }
            }
        }
        CancelOrphan(cfg);
    }

    /// <summary>
    /// 러너 없이 남은 기록(부모·러너 트리 kill) → 그 실행의 매크로에 취소를 요청하고, 매크로가 멈췄다는 근거
    /// (이 runId의 취소 기록, Excel 소멸)가 나오거나 <see cref="OrphanWaitMs"/>가 지날 때까지 명령을 거부한다 —
    /// 아직 바쁜 Excel에 다음 명령이 끼어들지 않게. 상한이 지나면 기록을 지우고 평소의 세션 복구 규칙에 맡긴다.
    /// </summary>
    private static void CancelOrphan(ProjectConfig cfg)
    {
        string path = RecordPath(cfg);
        var rec = Load(path);
        if (rec == null)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { /* 손상 기록 */ }
            return;
        }
        if (RunnerAlive(rec)) return; // 잠금은 비었는데 러너가 산 경우(기록 직전/직후 경합) — 건드리지 않는다

        var now = DateTime.UtcNow;
        bool first = !DateTime.TryParse(rec.OrphanCancelUtc, null, System.Globalization.DateTimeStyles.RoundtripKind, out var since);
        if (first)
        {
            since = now;
            if (rec.CancelFlag.Length > 0) RunCancel.Request(rec.CancelFlag);
            rec.OrphanCancelUtc = now.ToString("O");
            try { File.WriteAllText(path, JsonSerializer.Serialize(rec, Json)); } catch { /* 다음 명령이 다시 시작 */ }
        }
        bool gone = !ProcessAlive(rec.ExcelPid)
            || (rec.Jsonl.Length > 0 && SafeSpanCancelled(rec.Jsonl, rec.RunId));
        if (gone && rec.CancelFlag.Length > 0) RunCancel.Clear(rec.CancelFlag); // 멈출 매크로가 없다
        bool stopped = gone || rec.CancelSupported == false; // false: 멈출 수단이 없다 - 기다려도 소용없다
        double waited = (now - since).TotalMilliseconds;
        if (stopped || waited >= OrphanWaitMs)
        {
            try { File.Delete(path); } catch { /* 다음 명령이 재시도 */ }
            return;
        }
        throw new CliError(
            $"The previous vbaloop run was killed while its macro was still running in Excel (PID {rec.ExcelPid}): "
            + rec.Describe(now) + " - it was asked to stop at its next LogShim call",
            $"Retry in a few seconds (after {OrphanWaitMs / 1000}s the tool stops waiting and the normal session recovery rules apply)",
            environment: true);
    }

    private static bool SafeSpanCancelled(string jsonl, string runId)
    {
        try { return RunCancel.SpanCancelled(jsonl, 0, runId); }
        catch { return false; }
    }

    private static bool ProcessAlive(int pid)
    {
        if (pid <= 0) return false;
        try
        {
            using var p = Process.GetProcessById(pid);
            return !p.HasExited;
        }
        catch
        {
            return false;
        }
    }

    private static bool RunnerAlive(ActiveRun rec)
        => ProcessIdentity.Identify(rec.RunnerPid, rec.RunnerStart) != IdentityState.Gone;
}

/// <summary>실행 진행 표시(설계 §14.5) — 순수 포맷.</summary>
public static class RunProgress
{
    public static string FormatElapsed(TimeSpan t)
        => t.TotalHours >= 1 ? $"{(int)t.TotalHours}h{t.Minutes:00}m{t.Seconds:00}s"
         : t.TotalMinutes >= 1 ? $"{(int)t.TotalMinutes}m{t.Seconds:00}s"
         : $"{Math.Max(0, (int)t.TotalSeconds)}s";

    public static string Line(string command, string target, TimeSpan elapsed, long records, TimeSpan? sinceLast)
        => $"vbaloop: {command} {target} still running - {FormatElapsed(elapsed)} elapsed, {records} LogShim record(s)"
           + (sinceLast is { } s ? $", last {FormatElapsed(s)} ago" : "");
}
