using System.Diagnostics;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// 격리 모델의 실행부 (설계 §11). 판정은 Core.RecoveryPolicy, 여기서는 마커 기록/해제·프로브·kill만.
/// 모든 진입점은 프로젝트 CommandMutex를 쥔 호출자에서만 부른다 — 마커 기록과 다음 명령의
/// 게이트 검사 사이에 다른 명령이 끼지 못하게 하기 위함.
/// </summary>
public static class RecoveryGate
{
    public enum Resolution { ProcessGone, Killed, Quarantined }

    /// <summary>
    /// 세션 필요 명령의 진입 게이트. 마커가 없으면 통과. 기록된 프로세스가 이미 사라졌으면
    /// 마커를 지우고 통과(kill 없음 — 지킬 대상이 없다). 그 밖에는 exit 3 CliError.
    /// 격리된 프로세스에는 COM 탐침도 하지 않는다(재프로브는 마커 기록 시점에 이미 실패한 판정의 반복).
    /// </summary>
    public static void Enforce(ProjectConfig cfg, ToolLog? log)
    {
        var m = RecoveryMarker.Load(cfg.RecoveryMarkerPath);
        if (m == null) return;
        if (!m.Corrupt && ProcessIdentity.Identify(m.Pid, m.ExcelStartTime) == IdentityState.Gone)
        {
            RecoveryMarker.Delete(cfg.RecoveryMarkerPath);
            log?.Write("info", "recovery_auto_cleared", new { pid = m.Pid, reason = m.Reason });
            return;
        }
        log?.Write("error", "recovery_quarantined", new { pid = m.Pid, reason = m.Reason, corrupt = m.Corrupt });
        throw Quarantined(m);
    }

    public static CliError Quarantined(RecoveryMarker m)
        => new($"세션 격리 중 — {m.Describe()}", HintFor(m), environment: true);

    /// <summary>
    /// 격리 해제 안내. --discard는 미저장 변경을 버리는 파괴적 동작이라 "사용자 확인 후"를 명시한다 —
    /// 에이전트가 격리를 만나자마자 스스로 kill하는 것을 막는 것이 격리 모델의 요점이다.
    /// </summary>
    public static string HintFor(RecoveryMarker m)
    {
        if (m.Corrupt)
            return "_dbg/recovery.json을 확인한 뒤 vbaloop recovery clear --force";
        if (m.IsExternal)
            return $"사용자 Excel(PID {m.Pid})의 대화상자/실행 중 매크로를 사용자가 정리한 뒤 "
                   + "vbaloop session stop --discard(분리만, Excel은 그대로) — 상태 확인은 vbaloop recovery status";
        return $"사용자에게 Excel(PID {m.Pid})에 저장할 작업이 없는지 확인 → 사용자가 직접 닫으면 vbaloop recovery clear, "
               + "버려도 된다고 하면 vbaloop session stop --discard(미저장 변경 유실) — 상태 확인은 vbaloop recovery status";
    }

    /// <summary>
    /// 비정상 세션(타임아웃·poison·stale·사망)의 정리 시도. RecoveryPolicy.DecideAuto가 Kill을
    /// 허용할 때만 kill하고, 격리면 마커를 기록한다. 예외를 던지지 않는다 — 호출자가 결과로 분기.
    /// </summary>
    public static (Resolution Result, RecoveryMarker? Marker, string Why) Resolve(
        ProjectConfig cfg, SessionMeta meta, string reason, string? detail, string? command, ToolLog log)
    {
        var identity = ProcessIdentity.Identify(meta.Pid, meta.ExcelStartTime);
        ProbeResult? probe = identity == IdentityState.Match && !meta.IsExternal
            ? SessionProbe.Run(meta.Pid, meta.ExcelStartTime)
            : null;
        var d = RecoveryPolicy.DecideAuto(meta.IsExternal, identity, probe, meta.WorkbookFullName);
        log.Write("info", "recovery_decision", new
        {
            pid = meta.Pid, reason, action = d.Action.ToString(), why = d.Why,
            identity = identity.ToString(), probe = probe?.Kind.ToString(), workbooks = probe?.Workbooks,
        });

        string why = d.Why;
        switch (d.Action)
        {
            case AutoAction.Clear:
                return (Resolution.ProcessGone, null, why);
            case AutoAction.Kill:
                if (KillVerified(meta.Pid, meta.ExcelStartTime, log))
                    return (Resolution.Killed, null, why);
                why = $"{why}, 그러나 kill 후에도 프로세스가 남아 있음";
                break;
        }

        var marker = RecoveryMarker.For(meta, reason,
            string.IsNullOrEmpty(detail) ? why : $"{detail} / {why}", command);
        marker.Save(cfg.RecoveryMarkerPath);
        ExcelPidRegistry.TryRecord(ExcelPidRegistry.DefaultPath, meta.Pid, meta.ExcelStartTime,
            ExcelPidRegistry.KindQuarantined, cfg.RootDir);
        log.Write("error", "recovery_marker_written", new { pid = meta.Pid, reason, why });
        ComCleanup.SkipOnExit = true; // 응답 없는 Excel의 RCW 해제는 무기한 블록 가능
        return (Resolution.Quarantined, marker, why);
    }

    /// <summary>Resolve 결과를 실패 메시지 꼬리로 — 에이전트가 kill 여부를 오해하지 않도록 명시.</summary>
    public static string Explain(Resolution r, RecoveryMarker? marker, string why) => r switch
    {
        Resolution.ProcessGone => "Excel 프로세스도 종료됨 — 다음 명령이 세션을 자동 재기동",
        Resolution.Killed => $"세션 Excel을 정리함({why}) — 다음 명령이 세션을 자동 재기동",
        _ => $"Excel(PID {marker?.Pid})은 kill하지 않고 격리함({why})",
    };

    public static string HintAfter(Resolution r, RecoveryMarker? marker)
        => r == Resolution.Quarantined && marker != null ? HintFor(marker) : "session start";

    /// <summary>
    /// 신원(PID+StartTime) Match일 때만 트리 kill 후 소멸을 확인. Gone이면 true(이미 없음),
    /// Unknown이면 kill하지 않고 false.
    /// </summary>
    public static bool KillVerified(int pid, string startTimeIso, ToolLog log)
    {
        var before = ProcessIdentity.Identify(pid, startTimeIso);
        if (before == IdentityState.Gone) return true;
        if (before != IdentityState.Match) return false;
        try
        {
            using var p = Process.GetProcessById(pid);
            p.Kill(entireProcessTree: true);
            p.WaitForExit(10_000);
        }
        catch (Exception ex)
        {
            log.Write("error", "excel_kill_failed", new { pid, error = ex.Message });
        }
        bool gone = ProcessIdentity.Identify(pid, startTimeIso) == IdentityState.Gone;
        log.Write(gone ? "warn" : "error", gone ? "excel_killed" : "excel_kill_unconfirmed", new { pid });
        return gone;
    }
}
