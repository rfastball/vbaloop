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
        var meta = SessionEnsure.ForCommand(cfg, log, o, "test").Meta;

        string runId = Guid.NewGuid().ToString("N");
        int timeoutSec = opts.TimeoutSec ?? DefaultTimeoutSec;

        var sup = RunnerSupervisor.Run(cfg, meta, log, "test", runId, timeoutSec,
            Array.Empty<string>());

        if (sup.TimedOut)
        {
            o.Fail($"test 타임아웃({timeoutSec}s): 자식 러너 종료, 세션 오염 처리 — {sup.RecoveryTail}",
                sup.SessionHint, environment: true);
            return o;
        }

        if (sup.ExitCode == 3) // poison — 러너가 메타파일에 이미 기록함
        {
            log.Write("error", "runner_poisoned", new { runId });
            o.Payload = sup.Root != null ? BuildPayload(sup.Root.Value, sessionDied: false) : null;
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

        // 정상 완료 — 판정 집계
        var root = sup.Root.Value;
        var summaryEl = root.GetProperty("summary");
        int pass = summaryEl.GetProperty("pass").GetInt32();
        int fail = summaryEl.GetProperty("fail").GetInt32();
        int skipped = summaryEl.TryGetProperty("skipped", out var sk) ? sk.GetInt32() : 0;
        bool sessionDied = sup.SessionDiedAfterModal; // 러너 self-report + supervisor 백스톱 결합
        o.Payload = BuildPayload(root, sessionDied);

        o.HumanLines.Add($"test: pass {pass}, fail {fail}" + (skipped > 0 ? $", skipped {skipped}" : ""));
        if (root.TryGetProperty("watchdog", out var wdArr) && wdArr.ValueKind == JsonValueKind.Array)
        {
            int modalCount = wdArr.EnumerateArray()
                .Count(e => e.TryGetProperty("kind", out var k) && k.GetString() == "vbe_modal");
            if (modalCount > 0)
                o.HumanLines.Add(sessionDied
                    ? $"모달 자동 해제 {modalCount}건 — 세션 종료됨, 첫 FAIL의 모달 본문이 진짜 신호 (상세는 payload.watchdog)"
                    : $"모달 자동 해제 {modalCount}건 — 세션 유지됨 (상세는 payload.watchdog)");
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
            o.HumanLines.Add($"  SKIP {skipped}건 — 세션 종료로 미실행");

        if (sessionDied)
            o.HumanLines.Add("주의: 컴파일 오류 모달 해제로 Excel 세션이 종료됨 — 후속 명령 전에 session start 필요");
        if (sup.Marker != null)
            o.HumanLines.Add($"주의: {sup.RecoveryTail} — 해제 전까지 세션 명령은 exit 3");

        log.Write("info", "test_done", new { runId, pass, fail, skipped, sessionDied });
        if (fail > 0)
            o.Fail($"테스트 실패 {fail}건 (pass {pass}" + (skipped > 0 ? $", skipped {skipped}" : "")
                   + ") — 실패 레코드는 payload.tests 참조"
                   + (sessionDied ? " (세션 종료됨, session start 필요)" : ""),
                sessionDied ? sup.SessionHint : null);
        else if (skipped > 0)
            o.Fail($"세션 종료로 테스트 {skipped}건 미실행 (pass {pass}) — session start 후 재실행 필요",
                sup.SessionHint);
        else if (sessionDied)
            o.Hint = sup.SessionHint; // 전부 실행·통과했으나 마지막에 세션이 죽은 경우
        return o;
    }

    private static object? BuildPayload(JsonElement root, bool sessionDied)
    {
        if (!root.TryGetProperty("tests", out var tests) || !root.TryGetProperty("summary", out var summary))
            return null;
        return root.TryGetProperty("watchdog", out var wd)
            ? new { tests, summary, sessionDied, watchdog = wd }
            : new { tests, summary, sessionDied, watchdog = (object?)null };
    }
}
