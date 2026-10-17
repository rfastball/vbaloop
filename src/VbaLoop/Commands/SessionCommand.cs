using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>vbaloop session start|attach|stop [--discard] (설계 §4, 격리 모델·attach §11).</summary>
public static class SessionCommand
{
    public static Outcome Run(CliArgs opts)
    {
        string sub = opts.Positional.FirstOrDefault()
            ?? throw new CliError("session subcommand required", "session start | session attach | session stop");
        return sub switch
        {
            "start" => Start(opts),
            "attach" => Attach(opts),
            "stop" => Stop(opts),
            _ => throw new CliError($"Unknown session subcommand: {sub}", "session start | session attach | session stop"),
        };
    }

    private static Outcome Start(CliArgs opts)
    {
        var o = new Outcome { Command = "session start" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "session start");
        o.LogRecords = log.Records;

        // 게이트: 동시 명령 직렬화(mutex) + 프로젝트 소유권(lease 자동 획득)
        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "session start", log);

        // 핵심 로직은 SessionService.EnsureActive로 이동(auto-ensure와 공유) — 여기는
        // 결과를 Outcome으로 변환만 한다. payload 스키마는 기존과 동일 유지.
        var r = SessionService.EnsureActive(cfg, log, "session start");
        o.HumanLines.Add(r.Action == "adopted"
            ? $"Existing session adopted (PID {r.Meta.Pid}{(r.Meta.IsExternal ? ", attached user's Excel" : "")}, no-op)"
            : $"Session started (PID {r.Meta.Pid}, {Path.GetFileName(r.Meta.WorkbookFullName)})");
        if (cfg.SessionMode == SessionMode.Ephemeral && r.Action == "started")
            o.HumanLines.Add("Ephemeral mode, but while this session is alive push/pull/test/run/diff use this Excel — run session stop to end it");
        o.Payload = new { action = r.Action, pid = r.Meta.Pid, hwnd = r.Meta.Hwnd, workbook = r.Meta.WorkbookFullName };
        return o;
    }

    private static Outcome Stop(CliArgs opts)
    {
        var o = new Outcome { Command = "session stop" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "session stop");
        o.LogRecords = log.Records;

        // stop 특례: 내 lease거나 없으면 진행, 남의 살아있는 lease면 exit 2
        // (B의 stop 한 방이 A의 Excel을 종료시키는 것 방지). 성공 종결 시 자동 해제.
        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "session stop", log);

        if (opts.Discard)
            return Discard(o, cfg, log);

        // 격리 중이면 일반 stop도 거부(exit 3) — 해제는 recovery clear 또는 --discard.
        // 기록된 프로세스가 이미 사라졌으면 게이트가 마커를 지우고 통과시킨다.
        RecoveryGate.Enforce(cfg, log);

        var info = SessionService.Inspect(cfg);
        switch (info.State)
        {
            case SessionState.Busy:
                throw SessionService.BusyError(info);

            case SessionState.None:
                LeaseService.ReleaseIfMine(cfg, log);
                o.HumanLines.Add("No session (no-op)");
                o.Payload = new { action = "none" };
                return o;

            case SessionState.Stale:
            case SessionState.Poisoned:
                // 자동 kill 폐지(설계 §11): 살아있는 프로세스는 격리 규칙이 허용할 때만 정리
                if (info.Meta != null)
                    SessionService.DisposeUnhealthy(cfg, info, "session stop", log);
                SessionMeta.Delete(cfg.SessionMetaPath);
                log.Write("info", "session_cleaned", new { prevState = info.State.ToString() });
                LeaseService.ReleaseIfMine(cfg, log);
                o.HumanLines.Add($"{info.State.ToString().ToLowerInvariant()} session cleaned up");
                o.Payload = new { action = "cleaned", prevState = info.State.ToString().ToLowerInvariant() };
                return o;
        }

        dynamic app = info.App!;
        var meta = info.Meta!;

        // attach(external) 세션: 분리만 — 사용자 Excel을 저장/닫기/종료하지 않는다(push가 이미 저장한 것만 남음)
        if (meta.IsExternal)
        {
            SessionMeta.Delete(cfg.SessionMetaPath);
            log.Write("info", "session_detached", new { pid = meta.Pid });
            LeaseService.ReleaseIfMine(cfg, log);
            o.HumanLines.Add($"Session detached (user's Excel PID {meta.Pid} left as is — not saved/closed)");
            o.Payload = new { action = "detached", pid = meta.Pid, saved = false };
            return o;
        }

        var filter = OleMessageFilter.Register();
        try
        {
            app.DisplayAlerts = false;

            // 대상 워크북 Save 후 Close — dirty를 먼저 해소해 NUIDialog 여지 차단
            bool closedOurs = false;
            string? saveDrift = null;
            try
            {
                dynamic wb = SessionService.FindWorkbook(app, meta.WorkbookFullName);
                if (!opts.NoSave)
                {
                    ComRetry.Invoke(() => { wb.Save(); });
                    // push와 같은 조용한 SaveAs 오탈출 감지 — Close는 막지 않음
                    string actualFull = Convert.ToString(wb.FullName) ?? "";
                    if (!PathDrift.SamePath(actualFull, cfg.WorkbookPath))
                    {
                        saveDrift = actualFull;
                        log.Write("warn", "session_stop_saveas_drift",
                            new { expected = cfg.WorkbookPath, actual = actualFull });
                    }
                }
                ComRetry.Invoke(() => { wb.Close(false); });
                closedOurs = true;
                log.Write("info", "workbook_closed", new { saved = !opts.NoSave });
            }
            catch (CliError)
            {
                // 이미 닫혀 있음 — Quit 판단으로 진행
                log.Write("warn", "workbook_already_closed", null);
            }

            // 도구가 열지 않은(foreign) 워크북이 남아 있으면 Quit 거부 (사람 작업 보호)
            var foreign = new List<string>();
            int count = Convert.ToInt32(app.Workbooks.Count);
            for (int i = 1; i <= count; i++)
                foreign.Add(Convert.ToString(app.Workbooks.Item(i).FullName) ?? "?");

            if (foreign.Count > 0)
            {
                // 동명 워크북이 다른 경로에 열려 있으면 자기 워크북의 FullName 드리프트
                // (SaveAs 오탈출) 의심 — foreign 오인일 가능성을 안내
                var driftSuspects = PathDrift.FindDriftSuspects(foreign, cfg.WorkbookPath);
                SessionMeta.Delete(cfg.SessionMetaPath);
                log.Write("warn", "quit_refused_foreign_workbooks", new { foreign, driftSuspects });
                LeaseService.ReleaseIfMine(cfg, log);
                o.Payload = new
                {
                    action = "quit_refused",
                    closedWorkbook = closedOurs,
                    foreignWorkbooks = foreign,
                    driftSuspects,
                    saveDrift,
                };
                o.Fail($"Quit refused: {foreign.Count} non-tool workbooks open — Excel left running",
                    driftSuspects.Count > 0
                        ? $"Same-named workbook open at a different path ({driftSuspects[0]}) — suspected FullName drift (SaveAs escape), verify the canonical workbook is current"
                        : "Do not close Excel yourself; ask the user to save/close their workbooks and close Excel manually");
                return o;
            }

            ComRetry.Invoke(() => { app.Quit(); });
            SessionMeta.Delete(cfg.SessionMetaPath);
            log.Write("info", "session_stopped", new { pid = meta.Pid });
            LeaseService.ReleaseIfMine(cfg, log); // 작업 종료와 lease 수명 결합(자동 획득 모델의 짝)
            o.HumanLines.Add($"Session ended (PID {meta.Pid})");
            if (saveDrift != null)
                o.HumanLines.Add($"Warning: saved to {saveDrift} instead of the canonical workbook (suspected SaveAs escape) — verify the canonical workbook is current");
            o.Payload = new { action = "stopped", pid = meta.Pid, saved = !opts.NoSave, saveDrift };
            return o;
        }
        finally
        {
            filter.Revoke();
        }
    }

    /// <summary>
    /// session attach (설계 §11 L5): 사용자가 이미 자기 Excel에서 연 정본 워크북을 external 세션으로
    /// 채택한다. 이후 push/pull/test/run은 그 Excel에서 동작하고, stop은 분리만 한다. 사용자 Excel의
    /// 설정(DisplayAlerts/EnableEvents 등)은 바꾸지 않는다.
    /// </summary>
    private static Outcome Attach(CliArgs opts)
    {
        var o = new Outcome { Command = "session attach" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "session attach");
        o.LogRecords = log.Records;

        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "session attach", log);
        RecoveryGate.Enforce(cfg, log);

        if (!File.Exists(cfg.WorkbookPath))
            throw new CliError($"Workbook file not found: {cfg.WorkbookPath}", "Check the [excel].workbook path in vbaloop.toml");

        var info = SessionService.Inspect(cfg);
        switch (info.State)
        {
            case SessionState.Active when info.Meta!.IsExternal:
                o.HumanLines.Add($"Already attached (user's Excel PID {info.Meta.Pid}, no-op)");
                o.Payload = new { action = "adopted", pid = info.Meta.Pid, workbook = info.Meta.WorkbookFullName, owner = SessionMeta.OwnerExternal };
                return o;
            case SessionState.Active:
                o.Fail($"A tool-launched session (PID {info.Meta!.Pid}) is active — the canonical workbook is open in it",
                    "Keep working in that session, or run vbaloop session stop, have the user open the workbook in their Excel, then (after the user confirms) session attach");
                return o;
            case SessionState.Busy:
                throw SessionService.BusyError(info);
            case SessionState.Stale:
            case SessionState.Poisoned:
                if (info.Meta != null)
                    SessionService.DisposeUnhealthy(cfg, info, "session attach", log);
                SessionMeta.Delete(cfg.SessionMetaPath);
                break;
        }

        var scan = ExcelAttach.ScanOpenWorkbooks();
        var pick = AttachMatch.Pick(scan.Candidates, cfg.WorkbookPath);
        log.Write("info", "attach_scan", new
        {
            kind = pick.Kind.ToString(), candidates = scan.Candidates.Count, busy = scan.BusyPids, errors = scan.Errors,
        });
        switch (pick.Kind)
        {
            case AttachPickKind.NotOpen when scan.BusyPids.Count > 0:
                throw new CliError(
                    $"Excel (PID {string.Join(", ", scan.BusyPids)}) is busy; could not check open workbooks",
                    "Excel is in cell edit mode or showing a dialog; retry session attach shortly", environment: true);
            case AttachPickKind.NotOpen:
                o.Payload = new { action = "not_found", suspects = pick.Suspects, errors = scan.Errors };
                o.Fail($"No Excel found with the canonical workbook open: {cfg.WorkbookPath}",
                    pick.Suspects.Count > 0
                        ? $"Same-named workbook open at a different path/URL ({pick.Suspects[0]}) — check the vbaloop.toml path or the OneDrive sync path"
                        : "Retry after the user opens the canonical workbook in Excel — push/test without attach launches a separate Excel");
                return o;
            case AttachPickKind.ReadOnlyOnly:
                o.Fail($"Not adopted: canonical workbook is open read-only only (PID {pick.Target!.Pid}) — split-brain risk, push saves would leak to another location",
                    "The user should clear read-only or check the process holding the write lock (vbaloop doctor)");
                return o;
            case AttachPickKind.Ambiguous:
                o.Fail($"Canonical workbook is open writable in multiple Excel instances: {string.Join("; ", pick.Suspects)}",
                    "The user should close all but one, then retry");
                return o;
        }

        var target = pick.Target!;
        dynamic app = scan.Apps[target.Pid];
        string start = ProcessIdentity.GetStartTimeIso(target.Pid)
            ?? throw new CliError($"Could not read StartTime of PID {target.Pid}", environment: true);
        long hwnd = Convert.ToInt64(app.Hwnd);
        var meta = new SessionMeta
        {
            Pid = target.Pid,
            ExcelStartTime = start,
            Hwnd = hwnd,
            WorkbookFullName = target.FullName,
            CreatedAt = DateTime.UtcNow.ToString("o"),
            LastCommand = "session attach",
            Owner = SessionMeta.OwnerExternal,
        };
        meta.Save(cfg.SessionMetaPath);
        log.Write("info", "session_attached", new { pid = meta.Pid, hwnd, workbook = meta.WorkbookFullName });
        o.HumanLines.Add($"User's Excel adopted (PID {meta.Pid}, {Path.GetFileName(meta.WorkbookFullName)}) — "
                         + "push/pull/test/run operate in this Excel; session stop only detaches (no save/close)");
        o.Payload = new { action = "attached", pid = meta.Pid, hwnd, workbook = meta.WorkbookFullName, owner = SessionMeta.OwnerExternal };
        return o;
    }

    /// <summary>
    /// session stop --discard: 저장 없이 세션을 버린다(설계 §11). 격리 해제 경로이기도 하다.
    /// external(attach) 세션은 분리만 — 사용자 Excel은 절대 종료하지 않는다. 도구 기동 세션은
    /// PID+StartTime 신원이 일치할 때만 트리 kill하고, 프로브가 세션 외 워크북을 보여주면 거부한다.
    /// </summary>
    private static Outcome Discard(Outcome o, ProjectConfig cfg, ToolLog log)
    {
        var meta = SessionMeta.Load(cfg.SessionMetaPath);
        var marker = RecoveryMarker.Load(cfg.RecoveryMarkerPath);
        if (meta == null && marker == null)
        {
            LeaseService.ReleaseIfMine(cfg, log);
            o.HumanLines.Add("No session (no-op)");
            o.Payload = new { action = "none" };
            return o;
        }

        // 대상 신원: 격리 마커가 우선(격리된 바로 그 프로세스), 없으면 세션 메타
        bool fromMarker = marker is { Corrupt: false };
        int pid = fromMarker ? marker!.Pid : meta?.Pid ?? 0;
        string start = fromMarker ? marker!.ExcelStartTime : meta?.ExcelStartTime ?? "";
        bool external = fromMarker ? marker!.IsExternal : meta?.IsExternal ?? false;
        string workbook = (fromMarker ? marker!.Workbook : meta?.WorkbookFullName) ?? cfg.WorkbookPath;

        var identity = ProcessIdentity.Identify(pid, start);
        ProbeResult? probe = !external && identity == IdentityState.Match
            ? SessionProbe.Run(pid, start)
            : null;
        var d = RecoveryPolicy.DecideDiscard(external, identity, probe, workbook);
        log.Write("info", "session_discard_decision", new
        {
            pid, action = d.Action.ToString(), why = d.Why, identity = identity.ToString(),
            probe = probe?.Kind.ToString(), workbooks = probe?.Workbooks,
        });

        switch (d.Action)
        {
            case DiscardAction.Refuse:
                o.Payload = new { action = "refused", pid, reason = d.Why, workbooks = probe?.Workbooks };
                o.Fail($"session stop --discard refused: {d.Why}",
                    "Do not kill Excel yourself; inform the user. After the user saves/closes the work in that Excel, run vbaloop recovery clear");
                return o;

            case DiscardAction.Kill:
                if (!RecoveryGate.KillVerified(pid, start, log))
                {
                    if (marker == null && meta != null)
                    {
                        RecoveryMarker.For(meta, "discard_failed", "process still alive after kill", "session stop --discard")
                            .Save(cfg.RecoveryMarkerPath);
                        ExcelPidRegistry.TryRecord(ExcelPidRegistry.DefaultPath, meta.Pid, meta.ExcelStartTime,
                            ExcelPidRegistry.KindQuarantined, cfg.RootDir);
                    }
                    o.Payload = new { action = "kill_failed", pid };
                    o.Fail($"Could not confirm Excel PID {pid} exited — quarantine kept",
                        $"Ask the user to check EXCEL.EXE (PID {pid}) in Task Manager and close it, then run vbaloop recovery clear", environment: true);
                    return o;
                }
                break;
        }

        SessionMeta.Delete(cfg.SessionMetaPath);
        RecoveryMarker.Delete(cfg.RecoveryMarkerPath);
        LeaseService.ReleaseIfMine(cfg, log);
        string action = d.Action switch
        {
            DiscardAction.Kill => "discarded",
            DiscardAction.Detach => "detached",
            _ => "cleaned",
        };
        log.Write("warn", "session_discarded", new { pid, action });
        o.HumanLines.Add(d.Action switch
        {
            DiscardAction.Kill => $"Session discarded: Excel PID {pid} killed (not saved), quarantine lifted",
            DiscardAction.Detach => $"Session detached: user's Excel PID {pid} left as is (not saved/closed), quarantine lifted",
            _ => $"Session cleaned up: {d.Why}, quarantine lifted",
        });
        o.Payload = new { action, pid, reason = d.Why };
        return o;
    }
}
