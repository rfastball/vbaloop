using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop recovery status|clear [--force] (설계 §11). 격리 마커의 조회와 해제.
/// clear는 kill하지 않는다 — 기록된 프로세스가 사라졌을 때만 해제하고, --force는 살아있어도
/// 경고와 함께 해제한다(그 Excel은 도구가 잊고 사용자 프로세스로 남는다). 버리는(kill) 경로는
/// session stop --discard 하나뿐이다.
/// </summary>
public static class RecoveryCommand
{
    public static Outcome Run(CliArgs opts)
    {
        string sub = opts.Positional.FirstOrDefault() ?? "status";
        return sub switch
        {
            "status" => Status(opts),
            "clear" => Clear(opts),
            _ => throw new CliError($"미지 recovery 하위 명령: {sub}", "recovery status 또는 recovery clear [--force]"),
        };
    }

    private static Outcome Status(CliArgs opts)
    {
        var o = new Outcome { Command = "recovery status" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var m = RecoveryMarker.Load(cfg.RecoveryMarkerPath);
        if (m == null)
        {
            o.HumanLines.Add("격리 없음");
            o.Payload = new { quarantined = false };
            return o;
        }
        var identity = m.Corrupt ? IdentityState.Unknown : ProcessIdentity.Identify(m.Pid, m.ExcelStartTime);
        o.HumanLines.Add(m.Describe());
        o.HumanLines.Add(identity == IdentityState.Gone
            ? "기록된 Excel 프로세스는 이미 종료됨 — recovery clear 또는 다음 세션 명령이 자동 해제"
            : "기록된 Excel 프로세스가 살아있음");
        o.Hint = identity == IdentityState.Gone ? "vbaloop recovery clear" : RecoveryGate.HintFor(m);
        o.Payload = new { quarantined = true, processState = identity.ToString().ToLowerInvariant(), marker = Describe(m) };
        return o;
    }

    private static Outcome Clear(CliArgs opts)
    {
        var o = new Outcome { Command = "recovery clear" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "recovery clear");
        o.LogRecords = log.Records;

        // mutex: 마커 기록(타임아웃/poison 경로)과 해제가 교차하지 않게. lease는 잡지 않는다 —
        // 소멸 확인 해제는 누가 해도 안전하고, --force는 사람의 명시 판단이다.
        using var cmdLock = CommandMutex.Acquire(cfg);

        var m = RecoveryMarker.Load(cfg.RecoveryMarkerPath);
        if (m == null)
        {
            o.HumanLines.Add("격리 없음 (no-op)");
            o.Payload = new { action = "none" };
            return o;
        }

        var identity = m.Corrupt ? IdentityState.Unknown : ProcessIdentity.Identify(m.Pid, m.ExcelStartTime);
        var d = RecoveryPolicy.DecideClear(m.Corrupt, identity, opts.Force);
        if (d.Action == ClearAction.Refuse)
        {
            log.Write("warn", "recovery_clear_refused", new { pid = m.Pid, why = d.Why });
            o.Payload = new { action = "refused", reason = d.Why, marker = Describe(m) };
            o.Fail($"격리 해제 거부: {d.Why}", RecoveryGate.HintFor(m), environment: true);
            return o;
        }

        RecoveryMarker.Delete(cfg.RecoveryMarkerPath);
        // 세션 메타가 같은 프로세스를 가리키면 함께 잊는다 — --force 후 다음 명령이 같은 Excel을
        // 다시 탐침/격리하지 않게(사용자 소유로 남긴다는 결정을 존중).
        var meta = SessionMeta.Load(cfg.SessionMetaPath);
        if (meta != null && (m.Corrupt || meta.Pid == m.Pid))
            SessionMeta.Delete(cfg.SessionMetaPath);

        bool forced = d.Action == ClearAction.ForceClear;
        log.Write(forced ? "warn" : "info", forced ? "recovery_force_cleared" : "recovery_cleared",
            new { pid = m.Pid, reason = m.Reason, why = d.Why });
        o.HumanLines.Add(forced
            ? $"경고: {d.Why} — Excel PID {m.Pid}는 종료되지 않았고 도구가 더 이상 관리하지 않음(필요하면 직접 저장/종료)"
            : $"격리 해제: {d.Why}");
        o.Payload = new { action = forced ? "force_cleared" : "cleared", pid = m.Pid, reason = d.Why };
        return o;
    }

    private static object Describe(RecoveryMarker m) => new
    {
        pid = m.Pid,
        excelStartTime = m.ExcelStartTime,
        owner = m.Owner,
        reason = m.Reason,
        detail = m.Detail,
        command = m.Command,
        workbook = m.Workbook,
        createdAt = m.CreatedAt,
        corrupt = m.Corrupt,
    };

    /// <summary>status 명령과 공유하는 마커 요약.</summary>
    public static object? DescribeOrNull(RecoveryMarker? m) => m == null ? null : Describe(m);
}
