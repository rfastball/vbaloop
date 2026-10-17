using System.Diagnostics;
using System.Text;
using System.Text.Json;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// 자식 __runner 프로세스 감독 공용부(test/run 공유, 설계 §4 test 실행 모델):
/// 스폰 → 벽시계 타임아웃 감시(초과 시 자식 러너 kill → poisoned 기록 → Excel은 격리 규칙,
/// 설계 §11) → 비정상 종료 시 post-mortem 모달 스윕 → 러너 JSON 파싱.
/// </summary>
internal static class RunnerSupervisor
{
    public sealed record Result(
        bool TimedOut, int ExitCode, JsonElement? Root, string? Error, string? Hint, string Stderr,
        bool SessionDiedAfterModal = false,
        RecoveryGate.Resolution? Recovery = null, RecoveryMarker? Marker = null, string? RecoveryWhy = null)
    {
        /// <summary>일회용 인스턴스에서 실행했음 — 세션 재기동이 필요 없다(다음 명령이 새로 띄운다).</summary>
        public bool Ephemeral { get; init; }

        /// <summary>세션 관련 후속 안내 — 격리됐으면 해제 안내, 아니면 session start(일회용이면 재실행).</summary>
        public string SessionHint => Marker != null ? RecoveryGate.HintFor(Marker)
            : Ephemeral ? "같은 명령 재실행(일회용 Excel은 명령마다 새로 뜬다)" : "session start";

        /// <summary>세션 사망 뒤 할 일을 사람용 문장 꼬리로.</summary>
        public string RestartNote => Ephemeral ? "재실행하면 새 일회용 Excel로 진행" : "session start 필요";

        /// <summary>타임아웃/poison 메시지 꼬리 — kill 여부를 명시.</summary>
        public string RecoveryTail => Recovery is { } r ? RecoveryGate.Explain(r, Marker, RecoveryWhy ?? "") : "";
    }

    public static Result Run(
        ProjectConfig cfg, SessionMeta meta, ToolLog log,
        string commandName, string runId, int timeoutSec,
        IEnumerable<string> extraRunnerArgs)
        => RunCore(cfg, meta, log, commandName, runId, timeoutSec, extraRunnerArgs) with { Ephemeral = meta.Ephemeral };

    private static Result RunCore(
        ProjectConfig cfg, SessionMeta meta, ToolLog log,
        string commandName, string runId, int timeoutSec,
        IEnumerable<string> extraRunnerArgs)
    {
        string exe = Environment.ProcessPath
            ?? throw new CliError("자기 exe 경로를 확인할 수 없음");
        string toml = cfg.TomlPath; // --config로 받은 경로일 수 있다 — 조립하면 자식이 설정을 못 찾는다

        var psi = new ProcessStartInfo
        {
            FileName = exe,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        psi.ArgumentList.Add("__runner");
        psi.ArgumentList.Add("--config");
        psi.ArgumentList.Add(toml);
        psi.ArgumentList.Add("--runid");
        psi.ArgumentList.Add(runId);
        psi.ArgumentList.Add("--session-meta"); // session.json 또는 일회용 인스턴스 메타
        psi.ArgumentList.Add(meta.FilePath);
        foreach (var a in extraRunnerArgs)
            psi.ArgumentList.Add(a);

        log.Write("info", "runner_started", new { command = commandName, runId, timeoutSec });
        using var proc = Process.Start(psi)
            ?? throw new CliError("자식 러너 프로세스 시작 실패");
        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();

        if (!proc.WaitForExit(timeoutSec * 1000))
        {
            // 워치독이 못 알아보는 모달(MsgBox 등)/행의 최종 폴백. 러너(우리 프로세스)만 kill —
            // Excel은 사용자 작업이 없다고 증명될 때만 정리하고 아니면 격리(설계 §11, 자동 kill 폐지).
            try { proc.Kill(entireProcessTree: true); } catch { /* 이미 종료 */ }
            string reason = $"{commandName} 타임아웃({timeoutSec}s) — 워치독 관할 밖 모달/행 가능성";
            meta.MarkPoisoned(reason, null);
            log.Write("error", "runner_timeout", new { command = commandName, runId, timeoutSec });
            var (res, marker, why) = RecoveryGate.Resolve(cfg, meta, "timeout", reason, commandName, log);
            return new Result(true, -1, null, null, null, "", Recovery: res, Marker: marker, RecoveryWhy: why);
        }

        string stdout = stdoutTask.GetAwaiter().GetResult();
        string stderr = stderrTask.GetAwaiter().GetResult();
        var (root, error, hint) = Parse(stdout);

        if (proc.ExitCode == 3)
        {
            // poison: 러너가 메타에 poisoned를 기록함. 격리 판정은 mutex를 쥔 부모가 한다.
            var (pres, pmarker, pwhy) = RecoveryGate.Resolve(cfg, meta, "poison", error, commandName, log);
            return new Result(false, 3, root, error, hint, stderr,
                Recovery: pres, Marker: pmarker, RecoveryWhy: pwhy);
        }

        // 안전망 스윕: 러너가 비정상 종료했거나(모달 잔존), 정상 종료라도 워치독이 뭔가
        // 해제했으면(VBE 닫기 후속 "저장?" 프롬프트가 세션을 영구 차단할 수 있음) 잔존
        // 프롬프트를 정리한다 — 러너의 링거 드레인을 놓친 late 프롬프트의 부모측 백스톱.
        bool abnormal = proc.ExitCode != 0 && proc.ExitCode != 3;
        bool dismissed = root != null && root.Value.TryGetProperty("watchdog", out var wd)
            && wd.ValueKind == JsonValueKind.Array && wd.GetArrayLength() > 0;
        // attach(external) 세션은 사용자 Excel이다 — 사용자의 대화상자를 닫을 수 있는 사후 스윕은 하지 않는다
        if ((abnormal || dismissed) && !meta.IsExternal)
        {
            var swept = ModalHunter.SweepOnce(meta.Pid);
            if (swept.Count > 0)
                log.Write("warn", "postmortem_modal_sweep", new { swept });
        }

        // 컴파일 오류 해제는 VBE 닫기가 불가피하고, dirty 프로젝트에서 그것이 Excel 종료를
        // 유발할 수 있다(2026-07-08 실측 — 런타임 오류의 종료/End 경로는 VBE를 안 닫아 무사).
        // 1차 신호는 러너의 self-report(sessionDied — 안정화 폴링 포함 in-loop 판정).
        // 러너가 보고하지 않았고 모달 해제가 있었으면 백스톱: 종료는 지연되므로(러너 직후엔
        // 프로세스가 아직 살아있음) 짧게 안정화를 기다린 뒤 실제 attach 가능 여부로 판정한다.
        // 죽었으면 poison이 아니라 clean stale로 안내. 메타 삭제는 감지 경로와 무관하게 여기서.
        bool sessionDied = root != null
            && root.Value.TryGetProperty("sessionDied", out var sd)
            && sd.ValueKind == JsonValueKind.True;
        if (!sessionDied && dismissed)
        {
            Thread.Sleep(800); // XLMAIN 소멸/프로세스 종료가 완료되도록
            // Poisoned는 사망이 아니라 poison 경로(exit 3) 소관 — 메타의 poison 마커 보존
            var state = SessionService.Inspect(cfg, meta.FilePath).State;
            sessionDied = state is SessionState.Stale or SessionState.None;
        }
        RecoveryGate.Resolution? dres = null;
        RecoveryMarker? dmarker = null;
        string? dwhy = null;
        if (sessionDied)
        {
            // 워크북만 닫히고 프로세스는 창 없이 잔존할 수 있다 — 이 좀비가 워크북 파일
            // 잠금을 쥐면 다음 세션이 read-only로 열리는 버그의 트리거.
            // 창 없는 좀비는 격리 규칙의 Windowless로 정리되고, 다른 워크북이 보이면 격리(설계 §11).
            (dres, dmarker, dwhy) = RecoveryGate.Resolve(cfg, meta, "session_died",
                "컴파일 오류 모달 해제(VBE 닫기) 후 세션 워크북 소실", commandName, log);
            SessionMeta.Delete(meta.FilePath);
            log.Write("warn", "session_died_after_modal",
                new { note = "컴파일 오류 모달 해제(VBE 닫기)가 Excel 종료를 유발 — session start 필요", recovery = dres.ToString() });
        }

        return new Result(false, proc.ExitCode, root, error, hint, stderr, sessionDied,
            Recovery: dres, Marker: dmarker, RecoveryWhy: dwhy);
    }

    private static (JsonElement? Root, string? Error, string? Hint) Parse(string stdout)
    {
        int brace = stdout.IndexOf('{');
        if (brace < 0)
            return (null, "자식 러너가 JSON을 출력하지 않음", null);
        try
        {
            var doc = JsonDocument.Parse(stdout[brace..]);
            var root = doc.RootElement.Clone();
            string? error = root.TryGetProperty("error", out var e) && e.ValueKind == JsonValueKind.String
                ? e.GetString() : null;
            string? hint = root.TryGetProperty("hint", out var h) && h.ValueKind == JsonValueKind.String
                ? h.GetString() : null;
            return (root, error, hint);
        }
        catch (JsonException ex)
        {
            return (null, $"자식 러너 출력 파싱 실패: {ex.Message}", null);
        }
    }

    public static string Truncate(string s, int max)
        => s.Length <= max ? s : s[..max] + "...";
}
