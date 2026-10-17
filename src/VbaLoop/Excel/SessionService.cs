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
        if (!File.Exists(cfg.WorkbookPath))
            throw new CliError($"워크북 파일 없음: {cfg.WorkbookPath}",
                "vbaloop.toml의 [excel].workbook 경로 확인");

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

        var filter = OleMessageFilter.Register();
        try
        {
            var launch = ExcelLauncher.CreateNew(cfg.Visible);
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
                    wb = app.Workbooks.Open(cfg.WorkbookPath, 0, false); // UpdateLinks:=0, ReadOnly:=False
                }
                finally
                {
                    app.AutomationSecurity = prevSecurity;
                }

                // ReadOnly:=False 명시는 잠긴 파일에서 오류를 내지 않는다 — Excel이 조용히
                // read-only로 폴백하고, 이후 Save가 기본 저장 위치로 SaveAs 오탈출해 split-brain을
                // 만든다. read-only로는 세션을 성립시키지 않는다.
                if (Convert.ToBoolean(wb.ReadOnly))
                {
                    log.Write("error", "session_readonly_refused", new { workbook = cfg.WorkbookPath });
                    throw new CliError(
                        $"워크북이 read-only로 열림 — 다른 프로세스가 파일 쓰기 잠금 보유 의심(사용자가 Excel에서 열어둔 경우 포함): {cfg.WorkbookPath}",
                        "사용자가 이 워크북을 Excel에서 열어둔 것이면 사용자 확인 후 vbaloop session attach로 그 Excel을 채택"
                        + " — 아니면 vbaloop doctor로 잠금 보유 프로세스 진단(정리 여부는 사용자가 판단)",
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
                };
                meta.Save(cfg.SessionMetaPath);
                log.Write("info", "session_started",
                    new { pid = meta.Pid, hwnd = meta.Hwnd, workbook = meta.WorkbookFullName });
                return new EnsureResult(app, meta, "started");
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
    public static SessionInfo Inspect(ProjectConfig cfg)
    {
        var meta = SessionMeta.Load(cfg.SessionMetaPath);
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
        ToolLog? log = null)
    {
        RecoveryGate.Enforce(cfg, log);
        var info = Inspect(cfg);
        switch (info.State)
        {
            case SessionState.None:
                throw new CliError("활성 세션이 없음", "session start", environment: true);
            case SessionState.Poisoned:
                throw new CliError($"세션이 오염됨: {info.Detail}", "session start", environment: true);
            case SessionState.Stale:
                throw new CliError($"세션이 stale: {info.Detail}", "session start", environment: true);
            case SessionState.Busy:
                throw BusyError(info);
        }

        var meta = info.Meta!;
        meta.LastCommand = command;
        meta.Save(cfg.SessionMetaPath);
        return (info.App!, meta);
    }

    /// <summary>busy 세션 — 재시도 가능한 환경 오류(exit 3). stale로 취급해 정리하지 않는다.</summary>
    public static CliError BusyError(SessionInfo info)
        => new($"세션 Excel(PID {info.Meta?.Pid})이 busy — 호출 거절 중(셀 편집 중이거나 대화상자가 열려 있음): {info.Detail}",
            "Excel에서 셀 편집/열린 대화상자를 마친 뒤 같은 명령 재시도", environment: true);

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
        throw new CliError($"세션에 대상 워크북이 열려 있지 않음: {workbookFullName}", "session start",
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
                $"VBProject 접근 실패 (0x{ex.HResult:X8}) — VBIDE 신뢰 액세스(AccessVBOM) 미설정 가능성",
                "Excel 옵션 > 보안 센터 > 매크로 설정 > 'VBA 프로젝트 개체 모델에 안전하게 액세스' 체크");
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

        meta.MarkPoisoned(cfg.SessionMetaPath, $"{meta.LastCommand}: {ex.Message}", hr);
        log.Write("error", "session_poisoned", new { reason = ex.Message, hResult = hr });
        return true;
    }
}
