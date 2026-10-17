using System.Runtime.InteropServices;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>세션 필요 명령의 워크북 사용 의도. Read면 일회용 인스턴스가 정본을 ReadOnly:=True로 연다.</summary>
internal enum ScopeAccess { Write, Read }

/// <summary>
/// 세션 필요 명령(push/pull/test/run/diff)이 쓸 Excel 확보와 정리(설계 §4 세션 모드).
/// mutex+lease 획득 후에 호출할 것 — 다중 에이전트 동시 기동 레이스 방지.
/// <list type="bullet">
/// <item>persistent 모드: 구 auto-ensure 그대로(세션 채택 또는 기동, 명령 사이에 유지).</item>
/// <item>ephemeral 모드: 살아 있는 명시 세션(session start/attach)이 있으면 그것을 쓰고, 없으면 숨은
/// Excel을 이 명령 전용으로 띄운다. Dispose가 저장 없이 닫는다 — push는 성공 경로에서 스스로 저장한다.
/// 도구가 띄운 인스턴스라 비정상 상태의 정리도 격리 규칙(설계 §11)만 따르면 된다.</item>
/// </list>
/// </summary>
internal sealed class ExcelScope : IDisposable
{
    private readonly ProjectConfig _cfg;
    private readonly ToolLog _log;
    private readonly Outcome _o;
    private bool _disposed;

    public dynamic App { get; }
    public SessionMeta Meta { get; }
    public bool Ephemeral => Meta.Ephemeral;

    private ExcelScope(ProjectConfig cfg, ToolLog log, Outcome o, EnsureResult r)
    {
        _cfg = cfg;
        _log = log;
        _o = o;
        App = r.App;
        Meta = r.Meta;
    }

    public static ExcelScope Acquire(ProjectConfig cfg, ToolLog log, Outcome o, string command, ScopeAccess access)
    {
        if (cfg.SessionMode == SessionMode.Persistent)
        {
            var r = SessionService.EnsureActive(cfg, log, command);
            if (r.Action == "started")
            {
                o.Session = new { action = "started", pid = r.Meta.Pid, workbook = r.Meta.WorkbookFullName };
                o.HumanLines.Add($"Session auto-started (PID {r.Meta.Pid})");
            }
            return new ExcelScope(cfg, log, o, r);
        }

        var adopted = SessionService.AdoptExisting(cfg, log, command);
        if (adopted != null)
            return new ExcelScope(cfg, log, o, adopted);

        SweepLeftovers(cfg, log);
        string metaPath = Path.Combine(cfg.EphemeralDir, Guid.NewGuid().ToString("N") + ".json");
        bool readOnly = access == ScopeAccess.Read;
        var e = SessionService.LaunchEphemeral(cfg, log, command, readOnly, metaPath);
        o.Session = new { action = "ephemeral", pid = e.Meta.Pid, workbook = e.Meta.WorkbookFullName, readOnly };
        return new ExcelScope(cfg, log, o, e);
    }

    /// <summary>
    /// 이전 명령이 비정상 종료(하드 kill, 정리 실패)하며 남긴 일회용 인스턴스를 정리한다. CommandMutex
    /// 안에서만 불리므로 남은 메타는 전부 끝난 명령의 것이다. 정리는 격리 규칙 그대로 — 사용자 작업이
    /// 보이면 kill하지 않고 격리(exit 3)한다.
    /// </summary>
    private static void SweepLeftovers(ProjectConfig cfg, ToolLog log)
    {
        if (!Directory.Exists(cfg.EphemeralDir)) return;
        foreach (var file in Directory.EnumerateFiles(cfg.EphemeralDir, "*.json"))
        {
            var meta = SessionMeta.Load(file);
            if (meta != null && ProcessIdentity.Identify(meta.Pid, meta.ExcelStartTime) != IdentityState.Gone)
            {
                var (res, marker, why) = RecoveryGate.Resolve(cfg, meta, "ephemeral_leftover",
                    "ephemeral Excel left over by a previous command", meta.LastCommand, log);
                log.Write("warn", "ephemeral_leftover", new { pid = meta.Pid, result = res.ToString(), why });
                TryDelete(file);
                if (res == RecoveryGate.Resolution.Quarantined)
                    throw RecoveryGate.Quarantined(marker!);
                continue;
            }
            TryDelete(file);
        }
    }

    /// <summary>
    /// 일회용 인스턴스 정리(명시 세션은 no-op). 정상 경로는 정본을 저장 없이 닫고 Quit — 남은 워크북이
    /// 있으면(사용자가 숨은 인스턴스로 파일을 연 경우) 창을 보이게 하고 그대로 둔다. 러너·가드가 이미
    /// poison/kill/격리했거나 Quit이 실패하면 격리 규칙으로 정리한다.
    /// </summary>
    public void Dispose()
    {
        if (_disposed || !Ephemeral) return;
        _disposed = true;

        var identity = ProcessIdentity.Identify(Meta.Pid, Meta.ExcelStartTime);
        var marker = RecoveryMarker.Load(_cfg.RecoveryMarkerPath);
        if (identity == IdentityState.Gone || (marker != null && marker.Pid == Meta.Pid))
        {
            TryDelete(Meta.FilePath); // 이미 사라졌거나 격리됨 — 마커가 신원을 들고 있다
            return;
        }

        // 러너/supervisor가 같은 파일에 poison을 기록했거나(세션 사망 처리로) 파일을 지웠을 수 있다
        var current = SessionMeta.Load(Meta.FilePath);
        bool healthy = current is { Poisoned: false } && !Meta.Poisoned;
        if (healthy && QuitClean())
        {
            TryDelete(Meta.FilePath);
            return;
        }

        var (res, mk, why) = RecoveryGate.Resolve(_cfg, Meta, "ephemeral_cleanup",
            healthy ? "ephemeral Excel Quit failed" : current?.PoisonReason ?? Meta.PoisonReason, Meta.LastCommand, _log);
        TryDelete(Meta.FilePath);
        if (res == RecoveryGate.Resolution.Quarantined)
        {
            _o.HumanLines.Add($"Note: ephemeral Excel {RecoveryGate.Explain(res, mk, why)} — session commands exit 3 until cleared");
            _o.Hint ??= RecoveryGate.HintFor(mk!);
        }
    }

    /// <summary>저장 없이 닫고 Quit. 프로세스 종료까지 확인하면 true(남은 워크북을 넘겨준 경우 포함).</summary>
    private bool QuitClean()
    {
        var filter = OleMessageFilter.Register();
        try
        {
            App.DisplayAlerts = false;
            if (SessionService.IsWorkbookAlive(App, Meta.WorkbookFullName))
            {
                dynamic wb = SessionService.FindWorkbook(App, Meta.WorkbookFullName);
                ComRetry.Invoke(() => { wb.Close(false); });
            }

            var foreign = new List<string>();
            int count = Convert.ToInt32(App.Workbooks.Count);
            for (int i = 1; i <= count; i++)
                foreign.Add(Convert.ToString(App.Workbooks.Item(i).FullName) ?? "?");
            if (foreign.Count > 0)
            {
                // 숨은 인스턴스에 사용자가 연 파일이 들어온 경우: kill하지 않고 사용자 Excel로 넘긴다
                App.EnableEvents = true;
                App.DisplayAlerts = true;
                App.Visible = true;
                _log.Write("warn", "ephemeral_handed_over", new { pid = Meta.Pid, foreign });
                _o.HumanLines.Add($"Note: ephemeral Excel (PID {Meta.Pid}) has non-tool workbooks open; not closed, made visible: "
                                  + string.Join(", ", foreign));
                return true;
            }

            ComRetry.Invoke(() => { App.Quit(); });
            try { Marshal.FinalReleaseComObject(App); } catch { /* 이미 분리됨 */ }
        }
        catch (Exception ex)
        {
            _log.Write("warn", "ephemeral_quit_failed", new { pid = Meta.Pid, error = ex.Message });
            return false;
        }
        finally
        {
            filter.Revoke();
        }

        // Quit한 Excel은 외부 참조(RCW)가 남아 있는 동안 종료되지 않고, 명령 도중에는 그 참조를 다 풀 수 없다
        // (실측 2026-10-17: GC 후 5초를 기다려도 11회 전부 잔존 — 명령 시간의 대부분이 이 대기였다). 워크북 0개를
        // 방금 확인하고 Quit까지 보낸 도구 소유 인스턴스라 신원이 맞으면 kill해도 잃을 것이 없다 — dist 임시
        // Excel과 같은 처분. 짧은 유예는 Quit이 끝난 드문 경우를 kill 없이 넘기기 위함이다.
        long deadline = Environment.TickCount64 + 300;
        while (Environment.TickCount64 < deadline)
        {
            if (ProcessIdentity.Identify(Meta.Pid, Meta.ExcelStartTime) == IdentityState.Gone)
            {
                _log.Write("info", "ephemeral_closed", new { pid = Meta.Pid, how = "quit" });
                return true;
            }
            Thread.Sleep(100);
        }
        bool gone = RecoveryGate.KillVerified(Meta.Pid, Meta.ExcelStartTime, _log);
        _log.Write(gone ? "info" : "warn", "ephemeral_closed", new { pid = Meta.Pid, how = "kill_after_quit", gone });
        return gone;
    }

    private static void TryDelete(string path)
    {
        try { if (path.Length > 0 && File.Exists(path)) File.Delete(path); } catch { /* 다음 명령의 sweep이 정리 */ }
    }
}
