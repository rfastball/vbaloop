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
            ?? throw new CliError("lease subcommand required", "lease acquire|release|status|break --force");
        return sub switch
        {
            "acquire" => Acquire(opts),
            "release" => Release(opts),
            "status" => Status(opts),
            "break" => Break(opts),
            _ => throw new CliError($"Unknown lease subcommand: {sub}",
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
            ? $"lease acquired: {state.Owner?.OwnerThread}"
            : $"lease renewed (already held): {state.Owner?.OwnerThread}");
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
            o.HumanLines.Add("no lease (no-op)");
            o.Payload = new { action = "none" };
            return o;
        }
        if (!state.Mine)
        {
            // release는 owner만 가능 — 타 소유자 제거는 break --force가 유일 수단
            o.Payload = new { action = "refused", ownerThread = state.Owner?.OwnerThread };
            o.Fail($"Not my lease: {state.Owner?.OwnerThread ?? "(owner unreadable)"}",
                "To remove another owner's lease use vbaloop lease break --force");
            return o;
        }
        LeaseFile.Delete(LeaseFile.PathFor(cfg));
        log.Write("info", "lease_released", new { ownerThread = state.Owner!.OwnerThread });
        o.HumanLines.Add("lease released");
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
            throw new CliError("lease break removes the lease regardless of ownership", "Pass --force to confirm intent");

        var state = LeaseService.Inspect(cfg);
        if (!state.FilePresent)
        {
            o.HumanLines.Add("no lease (no-op)");
            o.Payload = new { action = "none" };
            return o;
        }
        LeaseFile.Delete(LeaseFile.PathFor(cfg));
        log.Write("warn", "lease_broken", new { prevOwner = state.Owner?.OwnerThread });
        o.HumanLines.Add($"lease force-removed (previous owner: {state.Owner?.OwnerThread ?? "unreadable"})");
        o.Payload = new { action = "broken", prevOwner = state.Owner?.OwnerThread };
        return o;
    }

    public static string DescribeState(LeaseService.LeaseState state)
    {
        if (state.Owner == null)
            return state.FilePresent ? "lease   : file present but unreadable (reclaimable)" : "lease   : none";
        string idle = double.IsNaN(state.IdleSeconds) ? "?" : $"{(int)state.IdleSeconds}s";
        string kind = state.Mine ? "mine" : state.Reclaimable ? "other's (reclaimable)" : "other's (active)";
        return $"lease   : {kind} — {state.Owner.OwnerThread} (idle {idle}, last {state.Owner.LastCommand})";
    }
}
