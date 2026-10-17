using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop lease acquire|release|status|break — 다중 에이전트 동시 접근 보호.
/// 변경성 명령(session/pull/push/test/run)은 자동 획득하므로 acquire는 선점용 선택지다.
/// </summary>
public static class LeaseCommand
{
    public static Outcome Run(CliArgs opts)
    {
        string sub = opts.Positional.FirstOrDefault()
            ?? throw new CliError("lease 하위 명령 필요", "lease acquire|release|status|break --force");
        return sub switch
        {
            "acquire" => Acquire(opts),
            "release" => Release(opts),
            "status" => Status(opts),
            "break" => Break(opts),
            _ => throw new CliError($"미지 lease 하위 명령: {sub}",
                "lease acquire|release|status|break --force"),
        };
    }

    private static Outcome Acquire(CliArgs opts)
    {
        var o = new Outcome { Command = "lease acquire" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "lease acquire");
        o.LogRecords = log.Records;

        var outcome = LeaseService.AcquireOrRenew(cfg, "lease acquire", log);
        var state = LeaseService.Inspect(cfg);
        o.HumanLines.Add(outcome == LeaseService.AcquireOutcome.Claimed
            ? $"lease 획득: {state.Owner?.OwnerThread}"
            : $"lease 갱신 (이미 보유): {state.Owner?.OwnerThread}");
        o.Payload = new
        {
            action = outcome.ToString().ToLowerInvariant(),
            ownerThread = state.Owner?.OwnerThread,
            ownerSource = state.Owner?.OwnerSource,
            expiresAt = state.Owner?.ExpiresAt,
        };
        return o;
    }

    private static Outcome Release(CliArgs opts)
    {
        var o = new Outcome { Command = "lease release" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "lease release");
        o.LogRecords = log.Records;

        var state = LeaseService.Inspect(cfg);
        if (state.Owner == null && !state.FilePresent)
        {
            o.HumanLines.Add("lease 없음 (no-op)");
            o.Payload = new { action = "none" };
            return o;
        }
        if (!state.Mine)
        {
            // release는 owner만 가능 — 타 소유자 제거는 break --force가 유일 수단
            o.Payload = new { action = "refused", ownerThread = state.Owner?.OwnerThread };
            o.Fail($"내 lease가 아님: {state.Owner?.OwnerThread ?? "(owner 판독 불가)"}",
                "타 소유자 lease 제거는 vbaloop lease break --force");
            return o;
        }
        LeaseFile.Delete(LeaseFile.PathFor(cfg));
        log.Write("info", "lease_released", new { ownerThread = state.Owner!.OwnerThread });
        o.HumanLines.Add("lease 해제됨");
        o.Payload = new { action = "released" };
        return o;
    }

    private static Outcome Status(CliArgs opts)
    {
        var o = new Outcome { Command = "lease status" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var state = LeaseService.Inspect(cfg);

        o.HumanLines.Add(DescribeState(state));
        o.Payload = new
        {
            state = state.Owner == null ? (state.FilePresent ? "corrupt" : "none")
                : state.Mine ? "owned" : state.Reclaimable ? "reclaimable" : "busy",
            ownerThread = state.Owner?.OwnerThread,
            ownerSource = state.Owner?.OwnerSource,
            host = state.Owner?.Host,
            lastCommand = state.Owner?.LastCommand,
            heartbeatAt = state.Owner?.HeartbeatAt,
            expiresAt = state.Owner?.ExpiresAt,
            idleSeconds = double.IsNaN(state.IdleSeconds) ? (double?)null : state.IdleSeconds,
            reclaimable = state.Reclaimable,
        };
        return o;
    }

    private static Outcome Break(CliArgs opts)
    {
        var o = new Outcome { Command = "lease break" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "lease break");
        o.LogRecords = log.Records;

        if (!opts.Force)
            throw new CliError("lease break는 소유권을 무시하고 강제 제거함", "--force로 의사를 명시");

        var state = LeaseService.Inspect(cfg);
        if (!state.FilePresent)
        {
            o.HumanLines.Add("lease 없음 (no-op)");
            o.Payload = new { action = "none" };
            return o;
        }
        LeaseFile.Delete(LeaseFile.PathFor(cfg));
        log.Write("warn", "lease_broken", new { prevOwner = state.Owner?.OwnerThread });
        o.HumanLines.Add($"lease 강제 제거됨 (이전 소유: {state.Owner?.OwnerThread ?? "판독 불가"})");
        o.Payload = new { action = "broken", prevOwner = state.Owner?.OwnerThread };
        return o;
    }

    public static string DescribeState(LeaseService.LeaseState state)
    {
        if (state.Owner == null)
            return state.FilePresent ? "lease   : 파일은 있으나 판독 불가 (회수 대상)" : "lease   : 없음";
        string idle = double.IsNaN(state.IdleSeconds) ? "?" : $"{(int)state.IdleSeconds}s";
        string kind = state.Mine ? "내 소유" : state.Reclaimable ? "타 소유(회수 가능)" : "타 소유(활성)";
        return $"lease   : {kind} — {state.Owner.OwnerThread} (유휴 {idle}, 마지막 {state.Owner.LastCommand})";
    }
}
