using System.Runtime.InteropServices;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>Busy: 살아있는 우리 Excel이 호출을 거절 중(셀 편집/모달) — 정리 대상이 아니라 재시도 대상(설계 §11).</summary>
public enum SessionState { None, Active, Poisoned, Stale, Busy }

public sealed class SessionInfo
{
    public SessionState State { get; init; }
    public SessionMeta? Meta { get; init; }
    public dynamic? App { get; init; }
    public string? Detail { get; init; }
}

/// <summary>EnsureActive 결과 — Action은 "adopted"(기존 세션 채택) 또는 "started"(신규 기동).</summary>
public sealed record EnsureResult(dynamic App, SessionMeta Meta, string Action);

/// <summary>세션 상태 판정/획득/오염 처리의 단일 창구 (설계 §4 공통 계약).</summary>
public static class SessionService
{
    /// <summary>
    /// 세션 자동 확보(auto-ensure): Active면 채택, None/Stale/Poisoned면 정리 후 신규 기동.
    /// session start 명령과 세션 필요 명령(push/pull/test/run) 진입부가 공유한다 — 세션은
    /// 어차피 한 번은 열려야 하므로 부재를 에러(exit 3 왕복)로 만들지 않고 도구가 스스로 연다.
    /// 기동 실패 CliError는 environment=true로 전파되어 기존 exit 3 프로토콜 유지.
    /// OleMessageFilter는 내부에서 완결(호출자 filter 등록 전에 반환됨 — 중첩 없음).
    /// 자식 러너는 이 경로 금지 — 부모가 보장하므로 AcquireOrThrow를 사용한다.
    /// </summary>
    public static EnsureResult EnsureActive(ProjectConfig cfg, ToolLog log, string command)
    {
        var adopted = AdoptExisting(cfg, log, command);
        if (adopted != null) return adopted;

        var meta = LaunchAndOpen(cfg, log, command, cfg.Visible, readOnly: false, cfg.SessionMetaPath,
            ephemeral: false, out dynamic app);
        log.Write("info", "session_started",
            new { pid = meta.Pid, hwnd = meta.Hwnd, workbook = meta.WorkbookFullName });
        return new EnsureResult(app, meta, "started");
    }

    /// <summary>
    /// EnsureActive의 앞 절반: 격리 게이트 → 살아 있는 세션(session.json)이면 채택, busy면 exit 3,
    /// stale/poisoned면 격리 규칙으로 정리하고 메타를 지운 뒤 null(=기동 필요). 일회용 모드도
    /// 이것으로 명시 세션(session start/attach)을 먼저 찾는다.
    /// </summary>
    public static EnsureResult? AdoptExisting(ProjectConfig cfg, ToolLog log, string command)
    {
        if (!File.Exists(cfg.WorkbookPath))
            throw new CliError($"Workbook file not found: {cfg.WorkbookPath}",
                "Check the [excel].workbook path in vbaloop.toml");

        // 격리 게이트(설계 §11): 마커가 있으면 Inspect(=COM 탐침) 전에 exit 3
        RecoveryGate.Enforce(cfg, log);

        var info = Inspect(cfg);

        // 멱등: 살아있으면 기존 세션 채택(no-op) — 고아 인스턴스와 불필요한 Excel 재기동을 피한다(설계 §4).
        if (info.State == SessionState.Active)
        {
            var m = info.Meta!;
            m.LastCommand = command;
            m.Save(cfg.SessionMetaPath);
            log.Write("info", "session_adopted", new { pid = m.Pid });
            return new EnsureResult(info.App!, m, "adopted");
        }

        if (info.State == SessionState.Busy)
            throw BusyError(info);

        if (info.State is SessionState.Poisoned or SessionState.Stale)
        {
            // 자동 kill 폐지(설계 §11): 프로세스가 살아있으면 "도구 기동 + 신원 일치 + 세션 워크북
            // 외 워크북 없음 확인"일 때만 정리하고, 그 밖의 불확실성은 격리(exit 3)로 멈춘다.
            if (info.Meta != null)
                DisposeUnhealthy(cfg, info, command, log);
            SessionMeta.Delete(cfg.SessionMetaPath);
            log.Write("info", "stale_meta_cleaned", new { prevState = info.State.ToString() });
        }
        return null;
    }

    /// <summary>
    /// 일회용 인스턴스 기동(설계 §4 세션 모드): 숨은 Excel에 정본을 열고 메타를 metaPath에 기록한다.
    /// readOnly면 ReadOnly:=True로 열어 사용자가 정본을 열어 둔 상태에서도 동작하고 저장은 불가능하다
    /// (pull/diff/test/run). 쓰기(push)는 read-only 폴백을 거부한다. 정리는 호출자(ExcelScope) 몫.
    /// </summary>
    public static EnsureResult LaunchEphemeral(ProjectConfig cfg, ToolLog log, string command, bool readOnly,
        string metaPath)
    {
        var meta = LaunchAndOpen(cfg, log, command, visible: false, readOnly, metaPath, ephemeral: true,
            out dynamic app);
        log.Write("info", "ephemeral_started",
            new { pid = meta.Pid, hwnd = meta.Hwnd, workbook = meta.WorkbookFullName, readOnly });
        return new EnsureResult(app, meta, "ephemeral");
    }

    private static SessionMeta LaunchAndOpen(ProjectConfig cfg, ToolLog log, string command, bool visible,
        bool readOnly, string metaPath, bool ephemeral, out dynamic appOut)
    {
        var filter = OleMessageFilter.Register();
        try
        {
            var launch = ExcelLauncher.CreateNew(visible);
            dynamic app = launch.App;
            try
            {
                // 오픈 정책(설계 §4):
                // 링크/보안/이벤트가 start를 블로킹하거나 Workbook_Open이 실행되는 것 차단
                app.DisplayAlerts = false;
                app.EnableEvents = false;
                app.AskToUpdateLinks = false;

                int prevSecurity = Convert.ToInt32(app.AutomationSecurity);
                app.AutomationSecurity = 1; // msoAutomationSecurityLow
                dynamic wb;
                try
                {
                    wb = app.Workbooks.Open(cfg.WorkbookPath, 0, readOnly); // UpdateLinks:=0
                }
                finally
                {
                    app.AutomationSecurity = prevSecurity;
                }

                // ReadOnly:=False 명시는 잠긴 파일에서 오류를 내지 않는다 — Excel이 조용히
                // read-only로 폴백하고, 이후 Save가 기본 저장 위치로 SaveAs 오탈출해 split-brain을
                // 만든다. 쓰기 의도로는 read-only 세션을 성립시키지 않는다.
                if (!readOnly && Convert.ToBoolean(wb.ReadOnly))
                {
                    log.Write("error", "session_readonly_refused", new { workbook = cfg.WorkbookPath });
                    throw new CliError(
                        $"Workbook opened read-only — another process may hold the file write lock (including the user having it open in Excel): {cfg.WorkbookPath}",
                        "If the user has this workbook open in Excel, after the user confirms adopt that Excel with vbaloop session attach"
                        + " — otherwise diagnose the lock-holding process with vbaloop doctor (the user decides whether to clean it up)",
                        environment: true);
                }

                var meta = new SessionMeta
                {
                    Pid = launch.Pid,
                    ExcelStartTime = launch.StartTimeIso,
                    Hwnd = launch.Hwnd,
                    WorkbookFullName = Convert.ToString(wb.FullName) ?? cfg.WorkbookPath,
                    CreatedAt = DateTime.UtcNow.ToString("o"),
                    LastCommand = command,
                    Ephemeral = ephemeral,
                };
                meta.Save(metaPath);
                appOut = app;
                return meta;
            }
            catch
            {
                try { app.Quit(); } catch { /* 이미 죽었으면 무시 */ }
                throw;
            }
        }
        finally
        {
            filter.Revoke();
        }
    }

    /// <summary>
    /// 비정상(Stale/Poisoned) 세션 처분 — 호출 뒤 호출자가 메타를 지운다. external이 단지 stale
    /// (사용자가 정본을 닫음/Excel 응답 불가)이면 조용히 분리한다: 사용자 Excel은 건드리지 않고
    /// 다음 확보는 도구 소유 세션으로 진행된다. 그 밖에는 격리 규칙(설계 §11) — 격리면 exit 3.
    /// </summary>
    public static void DisposeUnhealthy(ProjectConfig cfg, SessionInfo info, string command, ToolLog log)
    {
        var meta = info.Meta!;
        if (meta.IsExternal && info.State == SessionState.Stale)
        {
            log.Write("warn", "external_session_detached", new { pid = meta.Pid, detail = info.Detail });
            return;
        }
        var (res, marker, _) = RecoveryGate.Resolve(cfg, meta,
            info.State == SessionState.Poisoned ? "poison" : "stale", info.Detail, command, log);
        if (res == RecoveryGate.Resolution.Quarantined)
            throw RecoveryGate.Quarantined(marker!);
    }

    /// <summary>부작용 없이 상태 판정. Active면 attach된 Application을 함께 반환.</summary>
    /// <param name="metaPath">생략 시 session.json. 일회용 인스턴스는 자기 메타 파일을 준다.</param>
    public static SessionInfo Inspect(ProjectConfig cfg, string? metaPath = null)
    {
        var meta = SessionMeta.Load(metaPath ?? cfg.SessionMetaPath);
        if (meta == null)
            return new SessionInfo { State = SessionState.None };
        if (meta.Poisoned)
            return new SessionInfo
            {
                State = SessionState.Poisoned,
                Meta = meta,
                Detail = meta.PoisonReason,
            };

        var att = ExcelAttach.TryAttach(meta);
        if (att.App == null)
            return new SessionInfo
            {
                State = att.Busy ? SessionState.Busy : SessionState.Stale,
                Meta = meta,
                Detail = att.FailReason,
            };

        return new SessionInfo { State = SessionState.Active, Meta = meta, App = att.App };
    }

    /// <summary>Active 세션 필수 명령(pull/push/test)의 공통 전제. lastCommand 기록.</summary>
    public static (dynamic App, SessionMeta Meta) AcquireOrThrow(ProjectConfig cfg, string command,
        ToolLog? log = null, string? metaPath = null)
    {
        RecoveryGate.Enforce(cfg, log);
        var info = Inspect(cfg, metaPath);
        switch (info.State)
        {
            case SessionState.None:
                throw new CliError("No active session", "session start", environment: true);
            case SessionState.Poisoned:
                throw new CliError($"Session poisoned: {info.Detail}", "session start", environment: true);
            case SessionState.Stale:
                throw new CliError($"Session stale: {info.Detail}", "session start", environment: true);
            case SessionState.Busy:
                throw BusyError(info);
        }

        var meta = info.Meta!;
        meta.LastCommand = command;
        meta.Save(meta.FilePath);
        return (info.App!, meta);
    }

    /// <summary>busy 세션 — 재시도 가능한 환경 오류(exit 3). stale로 취급해 정리하지 않는다.</summary>
    public static CliError BusyError(SessionInfo info)
        => new($"Session Excel (PID {info.Meta?.Pid}) is busy — rejecting calls (cell edit mode or a dialog is open): {info.Detail}",
            "Excel is in cell edit mode or showing a dialog; retry the same command shortly", environment: true);

    /// <summary>세션의 대상 워크북 객체를 찾는다. 닫혀 있으면 CliError.</summary>
    public static dynamic FindWorkbook(dynamic app, string workbookFullName)
    {
        int count = Convert.ToInt32(app.Workbooks.Count);
        for (int i = 1; i <= count; i++)
        {
            dynamic wb = app.Workbooks.Item(i);
            string full = Convert.ToString(wb.FullName) ?? "";
            if (string.Equals(full, workbookFullName, StringComparison.OrdinalIgnoreCase))
                return wb;
        }
        throw new CliError($"Target workbook is not open in the session: {workbookFullName}", "session start",
            environment: true);
    }

    /// <summary>
    /// 워크북 생존 프로브 — 예외를 절대 던지지 않는다. Excel 사망(RPC 단절), 워크북만 닫힘,
    /// RCW 무효화 전부 false. 러너 루프의 세션 사망 감지용 — FindWorkbook은 CliError를 던져
    /// 러너 프로토콜(exit 4)을 깨므로 재사용하지 않는다. ComRetry 미경유(재시도 낭비 방지).
    /// </summary>
    public static bool IsWorkbookAlive(dynamic app, string workbookFullName)
    {
        try
        {
            int count = Convert.ToInt32(app.Workbooks.Count);
            for (int i = 1; i <= count; i++)
            {
                string full = Convert.ToString(app.Workbooks.Item(i).FullName) ?? "";
                if (string.Equals(full, workbookFullName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>VBProject 획득. VBIDE 신뢰 액세스 미설정 시 안내 hint.</summary>
    public static dynamic GetVbProject(dynamic wb)
    {
        try
        {
            return wb.VBProject;
        }
        catch (COMException ex)
        {
            throw new CliError(
                $"VBProject access failed (0x{ex.HResult:X8}) — VBIDE trusted access (AccessVBOM) may not be enabled",
                "Excel Options > Trust Center > Macro Settings > check 'Trust access to the VBA project object model'");
        }
    }

    /// <summary>
    /// 치명적 COM 오류(연결 끊김 계열)면 메타파일에 poison 기록 후 true.
    /// 이후 명령은 session start 재실행을 요구하게 된다.
    /// </summary>
    public static bool PoisonIfFatal(ProjectConfig cfg, SessionMeta? meta, Exception ex, ToolLog log)
    {
        var hr = (ex as COMException)?.HResult ?? (ex as InvalidComObjectException)?.HResult;
        bool fatal = ex is InvalidComObjectException
                     || (hr.HasValue && ComRetry.IsPoison(hr.Value));
        if (!fatal || meta == null) return false;

        meta.MarkPoisoned($"{meta.LastCommand}: {ex.Message}", hr);
        log.Write("error", "session_poisoned", new { reason = ex.Message, hResult = hr });
        return true;
    }
}
