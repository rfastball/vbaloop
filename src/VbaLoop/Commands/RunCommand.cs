using System.Text.Json;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop run Module.Proc (설계 §8): 임의 매크로 단건 실행. test와 같은 자식 러너/워치독/
/// 타임아웃 감독을 공유 — 임의 매크로는 테스트보다 인터랙티브 상태를 만날 확률이 높아
/// 워치독(P1) 이후로 순서가 묶여 있던 명령.
/// </summary>
public static class RunCommand
{
    private const int DefaultTimeoutSec = 120;

    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "run" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "run");
        o.LogRecords = log.Records;

        string macro = opts.Positional.FirstOrDefault()
            ?? throw new CliError("macro to run is required", "vbaloop run Module.Proc");
        if (!MacroName.TryParse(macro, out _, out _, out var nameError))
            throw new CliError(nameError!, "Format: vbaloop run Module.Proc");

        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "run", log);

        // 세션 확보 (auto-ensure: 없으면 스스로 기동)
        // 일회용 모드는 정본을 읽기 전용으로 연다 — 테스트가 바꾼 시트·상태는 저장되지 않고 인스턴스와 함께 사라진다
        using var ses = ExcelScope.Acquire(cfg, log, o, "run", ScopeAccess.Read);
        var meta = ses.Meta;

        string runId = Guid.NewGuid().ToString("N");
        int timeoutSec = opts.TimeoutSec ?? DefaultTimeoutSec;

        var sup = RunnerSupervisor.Run(cfg, meta, log, "run", runId, timeoutSec,
            new[] { "--macro", macro });

        if (sup.TimedOut)
        {
            o.Fail($"run timeout ({timeoutSec}s): child runner terminated, session marked poisoned — {sup.RecoveryTail}",
                sup.SessionHint, environment: true);
            return o;
        }
        if (sup.ExitCode == 3)
        {
            log.Write("error", "runner_poisoned", new { runId });
            o.Fail((sup.Error ?? "fatal COM error in child runner") + $" — {sup.RecoveryTail}",
                sup.SessionHint, environment: true);
            return o;
        }
        if (sup.ExitCode != 0 || sup.Root == null)
        {
            log.Write("error", "runner_failed", new { runId, exitCode = sup.ExitCode, stderr = sup.Stderr });
            o.Fail(sup.Error ?? $"child runner failed (exit {sup.ExitCode}): {RunnerSupervisor.Truncate(sup.Stderr, 300)}",
                sup.Hint);
            return o;
        }

        var root = sup.Root.Value;
        if (!root.TryGetProperty("run", out var run))
        {
            o.Fail("no run result in runner output");
            return o;
        }
        object? payload = root.TryGetProperty("watchdog", out var wd)
            ? new { run, sessionDied = sup.SessionDiedAfterModal, watchdog = wd }
            : new { run, sessionDied = sup.SessionDiedAfterModal, watchdog = (object?)null };
        o.Payload = payload;

        string result = run.TryGetProperty("result", out var r) ? r.GetString() ?? "?" : "?";
        string? detail = run.TryGetProperty("detail", out var d) && d.ValueKind == JsonValueKind.String
            ? d.GetString() : null;
        string returnType = run.TryGetProperty("returnType", out var rt) ? rt.GetString() ?? "?" : "?";
        string returnValue = run.TryGetProperty("returnValue", out var rv) ? rv.ToString() : "";
        long elapsed = run.TryGetProperty("elapsedMs", out var el) ? el.GetInt64() : 0;
        int failedAsserts = run.TryGetProperty("failedAsserts", out var fa) ? fa.GetInt32() : 0;

        o.HumanLines.Add($"run {macro}: {result} ({elapsed}ms)" +
                         (returnType != "Empty" ? $" → {returnValue} ({returnType})" : ""));
        if (failedAsserts > 0)
            o.HumanLines.Add($"Note: {failedAsserts} failed assertion(s) in range (unrelated to run success, see payload.run)");

        if (sup.SessionDiedAfterModal)
            o.HumanLines.Add($"Warning: Excel session ended after dismissing a compile error modal — {sup.RestartNote}");
        if (sup.Marker != null)
            o.HumanLines.Add($"Warning: {sup.RecoveryTail} — session commands exit 3 until cleared");

        log.Write("info", "run_done", new { runId, macro, result, elapsed, sessionDied = sup.SessionDiedAfterModal });
        if (result != "ok")
            o.Fail($"run {macro} failed" + (detail != null ? $": {detail}" : "")
                   + (sup.SessionDiedAfterModal ? $" (session ended, {sup.RestartNote})" : ""),
                sup.SessionDiedAfterModal ? sup.SessionHint : null);
        return o;
    }
}
