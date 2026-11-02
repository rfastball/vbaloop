using System.Text.Json;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop test (설계 §4): 자식 러너를 RunnerSupervisor로 감독. 모달은 러너 내 워치독이
/// kill 없이 해제(P1), 워치독 관할 밖 상황은 타임아웃 → 러너 kill → poisoned + 격리 규칙(설계 §11).
/// </summary>
public static class TestCommand
{
    private const int DefaultTimeoutSec = 120;

    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "test" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "test");
        o.LogRecords = log.Records;

        // 게이트: 자식 러너는 lease를 건드리지 않는다 — 부모가 배치 전체를 대표해 보유
        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "test", log);

        // 자식을 띄우기 전에 세션을 확보 (auto-ensure: 없으면 스스로 기동)
        // 일회용 모드는 정본을 읽기 전용으로 연다 — 테스트가 바꾼 시트·상태는 저장되지 않고 인스턴스와 함께 사라진다
        using var ses = ExcelScope.Acquire(cfg, log, o, "test", ScopeAccess.Read);
        var meta = ses.Meta;

        string runId = Guid.NewGuid().ToString("N");
        int timeoutSec = opts.TimeoutSec ?? DefaultTimeoutSec;

        var sup = RunnerSupervisor.Run(cfg, meta, log, "test", runId, timeoutSec,
            Array.Empty<string>());

        if (sup.TimedOut)
        {
            o.Fail($"test timeout ({timeoutSec}s): child runner terminated, session marked poisoned — {sup.RecoveryTail}",
                sup.SessionHint, environment: true);
            return o;
        }

        if (sup.ExitCode == 3) // poison — 러너가 메타파일에 이미 기록함
        {
            log.Write("error", "runner_poisoned", new { runId });
            o.Payload = sup.Root != null ? BuildPayload(sup.Root.Value, sessionDied: false, crash: null) : null;
            o.Fail((sup.Error ?? "fatal COM error in child runner") + $" — {sup.RecoveryTail}",
                sup.SessionHint, environment: true);
            return o;
        }
        if ((sup.ExitCode != 0 || sup.Root == null) && sup.Crash != null)
        {
            // 러너가 결과를 못 낸 채 Excel이 사라졌다 — 러너 오류보다 크래시가 먼저다
            o.Fail(sup.CrashDetail("test")! + $" (session ended, {sup.RestartNote})", sup.DeathHint);
            if (sup.RelaunchWarning != null) o.HumanLines.Add(sup.RelaunchWarning);
            o.Payload = new { sessionDied = true, excelCrash = sup.Crash };
            return o;
        }
        if (sup.ExitCode != 0 || sup.Root == null)
        {
            log.Write("error", "runner_failed", new { runId, exitCode = sup.ExitCode, stderr = sup.Stderr });
            o.Fail(sup.Error ?? $"child runner failed (exit {sup.ExitCode}): {RunnerSupervisor.Truncate(sup.Stderr, 300)}",
                sup.Hint);
            return o;
        }

        // 정상 완료 — 판정 집계
        var root = sup.Root.Value;
        var summaryEl = root.GetProperty("summary");
        int pass = summaryEl.GetProperty("pass").GetInt32();
        int fail = summaryEl.GetProperty("fail").GetInt32();
        int skipped = summaryEl.TryGetProperty("skipped", out var sk) ? sk.GetInt32() : 0;
        bool sessionDied = sup.SessionDiedAfterModal; // 러너 self-report + supervisor 백스톱 결합
        o.Payload = BuildPayload(root, sessionDied, sup.Crash);

        o.HumanLines.Add($"test: pass {pass}, fail {fail}" + (skipped > 0 ? $", skipped {skipped}" : ""));
        if (root.TryGetProperty("watchdog", out var wdArr) && wdArr.ValueKind == JsonValueKind.Array)
        {
            int modalCount = wdArr.EnumerateArray()
                .Count(e => e.TryGetProperty("kind", out var k) && k.GetString() == "vbe_modal");
            if (modalCount > 0)
                o.HumanLines.Add(sessionDied
                    ? $"{modalCount} modal(s) auto-dismissed — session ended, the first FAIL's modal text is the real signal (details in payload.watchdog)"
                    : $"{modalCount} modal(s) auto-dismissed — session kept (details in payload.watchdog)");
        }
        foreach (var t in root.GetProperty("tests").EnumerateArray())
        {
            string result = t.GetProperty("result").GetString() ?? "?";
            string name = $"{t.GetProperty("module").GetString()}.{t.GetProperty("name").GetString()}";
            string? detail = t.TryGetProperty("detail", out var d) && d.ValueKind == JsonValueKind.String
                ? d.GetString() : null;
            if (result == "fail")
                o.HumanLines.Add($"  FAIL {name}" + (detail != null ? $" — {detail}" : ""));
        }
        if (skipped > 0)
            o.HumanLines.Add($"  SKIP {skipped} — not run because the session ended");

        if (sessionDied)
            o.HumanLines.Add(sup.DeathWarning);
        if (sup.RelaunchWarning != null)
            o.HumanLines.Add(sup.RelaunchWarning);
        if (sup.Marker != null)
            o.HumanLines.Add($"Warning: {sup.RecoveryTail} — session commands exit 3 until cleared");

        log.Write("info", "test_done", new { runId, pass, fail, skipped, sessionDied });
        if (fail > 0)
            o.Fail((sup.Crash != null ? (sup.ExcelExit == ExcelExitKind.Exited ? "Excel exited during the run; " : "Excel crashed during the run; ") : "")
                   + $"{fail} test failure(s) (pass {pass}" + (skipped > 0 ? $", skipped {skipped}" : "")
                   + ") — see payload.tests for failure records"
                   + (sessionDied ? $" (session ended, {sup.RestartNote})" : ""),
                sessionDied ? sup.DeathHint : null);
        else if (skipped > 0)
            o.Fail((sup.Crash != null ? sup.CrashDetail("test")! + "; " : "")
                   + $"{skipped} test(s) not run because the session ended (pass {pass}) — {sup.RestartNote}, rerun required",
                sup.DeathHint);
        else if (sessionDied)
        {
            o.Hint = sup.DeathHint; // 전부 실행·통과했으나 마지막에 세션이 죽은 경우
            if (sup.Crash != null) // 테스트 밖(정리 중)에 사라졌어도 성공으로 넘기지 않는다
                o.Fail(sup.CrashDetail("test")! + $" after pass {pass} (session ended, {sup.RestartNote})", sup.DeathHint);
        }
        return o;
    }

    private static object? BuildPayload(JsonElement root, bool sessionDied, ExcelCrashInfo? crash)
    {
        if (!root.TryGetProperty("tests", out var tests) || !root.TryGetProperty("summary", out var summary))
            return null;
        return root.TryGetProperty("watchdog", out var wd)
            ? new { tests, summary, sessionDied, excelCrash = crash, watchdog = wd }
            : new { tests, summary, sessionDied, excelCrash = crash, watchdog = (object?)null };
    }
}
