using System.Diagnostics;
using System.Text;
using System.Text.Json;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>run/test 대기 한도(설계 §14.5). 호출마다 넘기는 예산은 선택(--timeout)이고, 기본은 프로젝트 설정의
/// 안전 상한뿐이다 — 버려진 실행 정리용이지 소요시간 추정이 아니다.</summary>
public sealed record RunLimits(int? TimeoutSec, int MaxSec)
{
    public long DeadlineMs => (long)Math.Min(TimeoutSec ?? int.MaxValue, MaxSec) * 1000;

    public string Describe() => TimeoutSec is int t && t <= MaxSec
        ? $"timeout ({t}s)"
        : $"safety cap reached ([run] maxSeconds = {MaxSec}s)";
}

/// <summary>
/// 자식 __runner 프로세스 감독 공용부(test/run 공유, 설계 §4 test 실행 모델, §14):
/// 스폰 → 완료까지 대기하며 취소 계기 감시(Ctrl+C, --timeout, 안전 상한) → 취소는 먼저 LogShim 협조적
/// 취소를 요청하고 유예 안에 멈추지 않으면 강제 경로(자식 러너 kill → poisoned 기록 → Excel은 격리 규칙,
/// 설계 §11) → 막는 대화상자는 협조적 취소가 불가능하므로 곧바로 강제 경로 → 비정상 종료 시 post-mortem
/// 모달 스윕 → 러너 JSON 파싱. 대기 중 진행 상황을 stderr에 주기적으로 쓴다.
/// </summary>
internal static class RunnerSupervisor
{
    private const int ProgressIntervalMs = 30_000;

    public sealed record Result(
        bool Aborted, int ExitCode, JsonElement? Root, string? Error, string? Hint, string Stderr,
        bool SessionDiedAfterModal = false,
        RecoveryGate.Resolution? Recovery = null, RecoveryMarker? Marker = null, string? RecoveryWhy = null,
        ExcelExitKind ExcelExit = ExcelExitKind.Alive, ExcelCrashInfo? Crash = null)
    {
        /// <summary>일회용 인스턴스에서 실행했음 — 세션 재기동이 필요 없다(다음 명령이 새로 띄운다).</summary>
        public bool Ephemeral { get; init; }

        /// <summary>강제 경로로 끝낸 이유(Aborted일 때) — 실패 메시지의 머리.</summary>
        public string? AbortReason { get; init; }

        /// <summary>협조적 취소를 요청했던 이유. 러너가 정상 반환했으면 결과는 "cancelled"일 수 있다.</summary>
        public string? CancelReason { get; init; }

        /// <summary>세션 관련 후속 안내 — 격리됐으면 해제 안내, 아니면 session start(일회용이면 재실행).</summary>
        public string SessionHint => Marker != null ? RecoveryGate.HintFor(Marker)
            : Ephemeral ? "rerun the same command (ephemeral Excel starts fresh for each command)" : "session start";

        /// <summary>컴파일 오류 모달로 세션이 죽은 뒤의 hint — 세션 문제가 아니라 src 문제임을 먼저 말한다
        /// (세션 재기동만 하고 같은 코드로 재실행하는 루프 방지).</summary>
        public string CompileErrorHint => "the project does not compile (the modal text is the error; it names no location, "
            + "so start with the modules you just changed) — fix src, push, then " + SessionHint;

        /// <summary>세션 사망 뒤 할 일을 사람용 문장 꼬리로.</summary>
        public string RestartNote => Ephemeral ? "rerunning proceeds with a new ephemeral Excel" : "session start required";

        /// <summary>세션 사망의 원인별 hint — 프로세스 소멸(크래시·Quit)이면 크래시 안내, 아니면 컴파일 오류 모달.</summary>
        public string DeathHint => Crash == null ? CompileErrorHint
            : ExcelExit == ExcelExitKind.Exited
                ? "the code ended Excel (Application.Quit or similar) — remove that from what vbaloop runs, push, then " + SessionHint
                : ExcelCrash.Hint(Crash.Relaunched, SessionHint);

        /// <summary>세션 사망 경고 한 줄(사람용).</summary>
        public string DeathWarning => Crash == null
            ? $"Warning: Excel session ended after dismissing a compile error modal — {RestartNote}"
            : ExcelExit == ExcelExitKind.Exited
                ? $"Warning: Excel (PID {Crash.Pid}) exited during the run (no crash signs) — {RestartNote}"
                : $"Warning: Excel (PID {Crash.Pid}) crashed during the run — {RestartNote}";

        /// <summary>프로세스 소멸 설명(what = 매크로/명령). 러너가 소멸 전에 반환해 부모 백스톱만 본 경우에도
        /// 실패 메시지가 크래시를 먼저 말하도록 쓴다. 소멸이 없으면 null.</summary>
        public string? CrashDetail(string what) => Crash == null ? null
            : ExcelExit == ExcelExitKind.Exited ? ExcelCrash.ExitedDetail(what, Crash.Pid)
            : ExcelCrash.Detail(what, Crash.Pid, Crash.ExitCode);

        /// <summary>재기동 Excel 경고(없으면 null).</summary>
        public string? RelaunchWarning => Crash is { Relaunched.Count: > 0 } c
            ? "Warning: " + ExcelCrash.RelaunchNote(c.Relaunched, WorkbookPath, c.WorkbookLocked) : null;

        /// <summary>RelaunchWarning 문구용 정본 경로.</summary>
        public string WorkbookPath { get; init; } = "";

        /// <summary>타임아웃/poison 메시지 꼬리 — kill 여부를 명시.</summary>
        public string RecoveryTail => Recovery is { } r ? RecoveryGate.Explain(r, Marker, RecoveryWhy ?? "") : "";
    }

    public static Result Run(
        ProjectConfig cfg, SessionMeta meta, ToolLog log,
        string commandName, string runId, RunLimits limits,
        IEnumerable<string> extraRunnerArgs)
        => RunCore(cfg, meta, log, commandName, runId, limits, extraRunnerArgs)
            with { Ephemeral = meta.Ephemeral, WorkbookPath = meta.WorkbookFullName };

    private static Result RunCore(
        ProjectConfig cfg, SessionMeta meta, ToolLog log,
        string commandName, string runId, RunLimits limits,
        IEnumerable<string> extraRunnerArgs)
    {
        string exe = Environment.ProcessPath
            ?? throw new CliError("cannot determine own exe path");
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
        psi.ArgumentList.Add("--parent-pid"); // 부모가 강제 종료되면 러너가 취소를 요청한다(설계 §14.5)
        psi.ArgumentList.Add(Environment.ProcessId.ToString());
        foreach (var a in extraRunnerArgs)
            psi.ArgumentList.Add(a);

        // 러너가 끝난 뒤에도 종료 코드를 읽을 수 있게 세션 Excel 핸들을 먼저 쥔다(크래시 백스톱)
        using var exitWatch = ExcelExitWatch.Open(meta.Pid, meta.ExcelStartTime);

        log.Write("info", "runner_started", new
        {
            command = commandName, runId, timeoutSec = limits.TimeoutSec, maxSeconds = limits.MaxSec,
        });
        string jsonl = Path.Combine(Path.GetDirectoryName(meta.WorkbookFullName)!, "_dbg",
            Path.GetFileNameWithoutExtension(meta.WorkbookFullName) + ".jsonl");
        var progress = new LogProgress(jsonl);
        var dialogs = new BlockingDialogTracker(BlockingHandles(meta.Pid));
        string cancelFlag = RunCancel.PathFor(meta.WorkbookFullName, runId);
        int ctrlC = 0;
        ConsoleCancelEventHandler onCtrlC = (_, e) => { e.Cancel = true; Interlocked.Increment(ref ctrlC); };
        Console.CancelKeyPress += onCtrlC;

        using var proc = Process.Start(psi)
            ?? throw new CliError("failed to start child runner process");
        var stdoutTask = proc.StandardOutput.ReadToEndAsync();
        var stderrTask = proc.StandardError.ReadToEndAsync();

        string? cancelReason = null, abortReason = null, abortKind = null;
        try
        {
            var sw = Stopwatch.StartNew();
            long cancelAt = -1, nextProgress = ProgressIntervalMs;
            long logLen = SafeLength(jsonl), lastGrowth = 0;
            long cancelScan = logLen; // 이 실행이 쓴 기록부터 취소 레코드를 찾는다
            bool macroStopped = false;
            while (!proc.WaitForExit(250))
            {
                long now = sw.ElapsedMilliseconds;
                long len = SafeLength(jsonl);
                if (len != logLen) { logLen = len; lastGrowth = now; }
                int presses = Volatile.Read(ref ctrlC);
                if (cancelReason == null)
                {
                    if (presses > 0) cancelReason = "interrupted (Ctrl+C)";
                    else if (now >= limits.DeadlineMs) cancelReason = limits.Describe();
                }
                if (cancelReason != null && cancelAt < 0)
                {
                    cancelAt = now;
                    RunCancel.Request(cancelFlag);
                    log.Write("warn", "run_cancel_requested", new { command = commandName, runId, reason = cancelReason });
                    Console.Error.WriteLine($"vbaloop: {cancelReason} - asking the macro to stop at its next LogShim call "
                        + $"(waiting up to {RunnerMain.CancelGraceMs / 1000}s" + (presses > 0 ? ", Ctrl+C again forces it)" : ")"));
                }
                if (abortReason == null && cancelAt >= 0)
                {
                    if (!macroStopped && RunCancel.ScanForCancel(jsonl, ref cancelScan, runId))
                    {
                        // 매크로는 멈췄다 - 러너의 뒷정리(§14.8, 계산 모드 복원의 재계산 등)와 결과 출력에 유예를 새로 준다
                        macroStopped = true;
                        cancelAt = now;
                    }
                    if (presses > 1)
                    {
                        abortKind = "cancel";
                        abortReason = $"{cancelReason}; forced by a second Ctrl+C";
                    }
                    else if (now - cancelAt >= RunnerMain.CancelGraceMs)
                    {
                        abortKind = "cancel";
                        abortReason = macroStopped
                            ? $"{cancelReason}; the macro stopped at a LogShim check point but the runner did not finish within {RunnerMain.CancelGraceMs / 1000}s after it"
                            : $"{cancelReason}; the macro did not stop at a LogShim check point within {RunnerMain.CancelGraceMs / 1000}s";
                    }
                }
                // 기록이 계속 늘면 막힌 게 아니다(Activate에서 일하는 모달 진행 폼 등)
                if (abortReason == null && dialogs.Observe(BlockingHandles(meta.Pid), now) is IntPtr blocked
                    && now - lastGrowth >= dialogs.PersistMs)
                {
                    // 대화상자에 막힌 VBA는 LogShim에 닿지 않는다 — 협조적 취소 없이 곧바로 강제 경로
                    abortKind = "dialog";
                    abortReason = $"blocked by a dialog for {dialogs.PersistMs / 1000}s ({Truncate(ModalHunter.CollectInfo(blocked), 300)})"
                                  + " - headless runs cannot answer MsgBox/InputBox/UserForm/file dialogs";
                }
                if (abortReason != null) break;
                if (now >= nextProgress)
                {
                    nextProgress += ProgressIntervalMs;
                    var (records, sinceLast) = progress.Read();
                    Console.Error.WriteLine(RunProgress.Line(commandName, commandName == "test" ? "Test_*" : TargetOf(extraRunnerArgs),
                        sw.Elapsed, records, sinceLast));
                }
            }
        }
        finally
        {
            Console.CancelKeyPress -= onCtrlC;
        }

        if (abortReason != null)
        {
            // 러너(우리 프로세스)만 kill — Excel은 사용자 작업이 없다고 증명될 때만 정리하고 아니면 격리
            // (설계 §11, §14.3). 취소 표시는 남겨 둔다: 아직 도는 매크로가 다음 LogShim 호출에서 멈추도록.
            try { proc.Kill(entireProcessTree: true); } catch { /* 이미 종료 */ }
            try { proc.WaitForExit(5000); } catch { /* 이미 종료 */ }
            RunCancel.Request(cancelFlag); // 대화상자 경로 포함 — 누가 대화상자에 답하면 다음 LogShim 호출에서 멈춘다
            RunLock.Forget(cfg, runId);
            string reason = $"{commandName} {abortReason}";
            meta.MarkPoisoned(reason, null);
            log.Write("error", "runner_aborted", new { command = commandName, runId, kind = abortKind, reason = abortReason });
            var (res, marker, why) = RecoveryGate.Resolve(cfg, meta, abortKind == "dialog" ? "dialog" : "timeout",
                reason, commandName, log);
            if (ProcessIdentity.Identify(meta.Pid, meta.ExcelStartTime) == IdentityState.Gone)
                RunCancel.Clear(cancelFlag); // Excel을 정리했다 — 멈출 매크로가 없다
            return new Result(true, -1, null, null, null, "", Recovery: res, Marker: marker, RecoveryWhy: why)
                { AbortReason = abortReason, CancelReason = cancelReason };
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

        // Excel 프로세스 자체의 소멸(크래시·Quit): 러너 self-report가 1차, 모달 해제가 없었는데 프로세스가 사라졌으면
        // 부모의 핸들로 백스톱(러너가 소멸 전에 반환한 경우). 모달 사망은 정상 종료라 아래 기존 경로가 맡는다.
        var excelExit = ParseExit(root);
        if (excelExit == ExcelExitKind.Alive && !sessionDied && !dismissed)
            excelExit = exitWatch.Classify(0);
        if (excelExit != ExcelExitKind.Alive)
            return Crashed(cfg, meta, log, commandName, exitWatch, excelExit,
                new Result(false, proc.ExitCode, root, error, hint, stderr, SessionDiedAfterModal: true));

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
                "session workbook lost after compile error modal dismissal (VBE closed)", commandName, log);
            SessionMeta.Delete(meta.FilePath);
            log.Write("warn", "session_died_after_modal",
                new { note = "dismissing the compile error modal (closing VBE) caused Excel to exit — session start required", recovery = dres.ToString() });
        }

        return new Result(false, proc.ExitCode, root, error, hint, stderr, sessionDied,
            Recovery: dres, Marker: dmarker, RecoveryWhy: dwhy) { CancelReason = cancelReason };
    }

    private static long SafeLength(string path)
    {
        try { return File.Exists(path) ? new FileInfo(path).Length : 0; }
        catch { return 0; }
    }

    private static IEnumerable<IntPtr> BlockingHandles(int pid)
    {
        try { return ModalHunter.FindBlockingDialogs(pid).Select(d => d.Hwnd).ToList(); }
        catch { return Array.Empty<IntPtr>(); }
    }

    private static string TargetOf(IEnumerable<string> runnerArgs)
    {
        var list = runnerArgs.ToList();
        int i = list.IndexOf("--macro");
        return i >= 0 && i + 1 < list.Count ? list[i + 1] : "";
    }

    /// <summary>LogShim 기록 파일의 증가분 — 진행 표시용(레코드 수, 마지막 기록 이후 경과).</summary>
    private sealed class LogProgress(string path)
    {
        private readonly long _start = SafeLength(path);
        private long _pos = -1;
        private long _records;
        private DateTime? _lastWriteUtc;

        public (long Records, TimeSpan? SinceLast) Read()
        {
            try
            {
                if (_pos < 0) _pos = _start;
                if (File.Exists(path))
                {
                    using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    if (fs.Length < _pos) _pos = 0; // LogShim.Reset이 잘라냈다
                    fs.Seek(_pos, SeekOrigin.Begin);
                    var buf = new byte[64 * 1024];
                    int n;
                    while ((n = fs.Read(buf, 0, buf.Length)) > 0)
                        for (int i = 0; i < n; i++)
                            if (buf[i] == (byte)'\n') _records++;
                    _pos = fs.Length;
                    if (_records > 0) _lastWriteUtc = File.GetLastWriteTimeUtc(path);
                }
            }
            catch { /* 진행 표시는 최선 노력 */ }
            return (_records, _lastWriteUtc is DateTime t ? DateTime.UtcNow - t : null);
        }

        private static long SafeLength(string p)
        {
            try { return File.Exists(p) ? new FileInfo(p).Length : 0; }
            catch { return 0; }
        }
    }

    private static ExcelExitKind ParseExit(JsonElement? root)
    {
        if (root == null || !root.Value.TryGetProperty("excelExit", out var x) || x.ValueKind != JsonValueKind.String)
            return ExcelExitKind.Alive;
        return x.GetString() == "exited" ? ExcelExitKind.Exited : ExcelExitKind.Crashed;
    }

    /// <summary>
    /// 세션 Excel이 실행 도중 사라졌다. 크래시면 Office가 다시 띄운 Excel(부모 = 크래시한 PID)을 찾아 보고만 한다 —
    /// 보이는 창이고 사용자용 복구 상태(정본 + 복구 사본)를 들고 있을 수 있어 kill하지 않는다(설계 §11).
    /// 기록된 프로세스는 이미 없으므로 격리 규칙은 Clear로 끝나고, 세션 메타는 지운다(다음 명령이 새로 띄운다).
    /// </summary>
    private static Result Crashed(ProjectConfig cfg, SessionMeta meta, ToolLog log, string commandName,
        ExcelExitWatch exitWatch, ExcelExitKind kind, Result r)
    {
        var info = exitWatch.Describe(kind, meta.WorkbookFullName);
        log.Write(kind == ExcelExitKind.Crashed ? "error" : "warn",
            kind == ExcelExitKind.Crashed ? "excel_crashed" : "excel_exited", new
            {
                command = commandName,
                pid = info.Pid,
                exitCode = ExcelCrash.FormatExitCode(info.ExitCode),
                relaunched = info.Relaunched,
                workbookLocked = info.WorkbookLocked,
                workbook = meta.WorkbookFullName,
            });
        var (res, marker, why) = RecoveryGate.Resolve(cfg, meta,
            kind == ExcelExitKind.Crashed ? "excel_crashed" : "excel_exited",
            $"Excel process exited during {commandName}", commandName, log);
        SessionMeta.Delete(meta.FilePath);
        return r with { Recovery = res, Marker = marker, RecoveryWhy = why, ExcelExit = kind, Crash = info };
    }

    private static (JsonElement? Root, string? Error, string? Hint) Parse(string stdout)
    {
        int brace = stdout.IndexOf('{');
        if (brace < 0)
            return (null, "child runner produced no JSON output", null);
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
            return (null, $"failed to parse child runner output: {ex.Message}", null);
        }
    }

    public static string Truncate(string s, int max)
        => s.Length <= max ? s : s[..max] + "...";
}
