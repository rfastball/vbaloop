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
            ?? throw new CliError("실행할 매크로 필요", "vbaloop run Module.Proc");
        if (!MacroName.TryParse(macro, out _, out _, out var nameError))
            throw new CliError(nameError!, "형식: vbaloop run Module.Proc");

        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "run", log);

        // 세션 확보 (auto-ensure: 없으면 스스로 기동)
        var meta = SessionEnsure.ForCommand(cfg, log, o, "run").Meta;

        string runId = Guid.NewGuid().ToString("N");
        int timeoutSec = opts.TimeoutSec ?? DefaultTimeoutSec;

        var sup = RunnerSupervisor.Run(cfg, meta, log, "run", runId, timeoutSec,
            new[] { "--macro", macro });

        if (sup.TimedOut)
        {
            o.Fail($"run 타임아웃({timeoutSec}s): 자식 러너 종료, 세션 오염 처리 — {sup.RecoveryTail}",
                sup.SessionHint, environment: true);
            return o;
        }
        if (sup.ExitCode == 3)
        {
            log.Write("error", "runner_poisoned", new { runId });
            o.Fail((sup.Error ?? "자식 러너에서 치명적 COM 오류") + $" — {sup.RecoveryTail}",
                sup.SessionHint, environment: true);
            return o;
        }
        if (sup.ExitCode != 0 || sup.Root == null)
        {
            log.Write("error", "runner_failed", new { runId, exitCode = sup.ExitCode, stderr = sup.Stderr });
            o.Fail(sup.Error ?? $"자식 러너 실패 (exit {sup.ExitCode}): {RunnerSupervisor.Truncate(sup.Stderr, 300)}",
                sup.Hint);
            return o;
        }

        var root = sup.Root.Value;
        if (!root.TryGetProperty("run", out var run))
        {
            o.Fail("러너 출력에 run 결과가 없음");
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
            o.HumanLines.Add($"참고: 구간 내 실패 어서션 {failedAsserts}건 (run 성패와 무관, payload.run 참조)");

        if (sup.SessionDiedAfterModal)
            o.HumanLines.Add("주의: 컴파일 오류 모달 해제로 Excel 세션이 종료됨 — 후속 명령 전에 session start 필요");
        if (sup.Marker != null)
            o.HumanLines.Add($"주의: {sup.RecoveryTail} — 해제 전까지 세션 명령은 exit 3");

        log.Write("info", "run_done", new { runId, macro, result, elapsed, sessionDied = sup.SessionDiedAfterModal });
        if (result != "ok")
            o.Fail($"run {macro} 실패" + (detail != null ? $": {detail}" : "")
                   + (sup.SessionDiedAfterModal ? " (세션 종료됨, session start 필요)" : ""),
                sup.SessionDiedAfterModal ? sup.SessionHint : null);
        return o;
    }
}
