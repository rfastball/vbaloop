using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop __runner — test의 자식 러너 프로세스(설계 §4 test).
/// Application.Run은 블로킹 COM 호출이라 인프로세스 타임아웃이 불가능하므로,
/// 부모가 벽시계 타임아웃을 감시하고 이 프로세스는 스스로 재attach해 테스트를 순차 실행한다.
/// 출력: stdout에 JSON 한 덩어리. exit 0=실행 완료(실패·취소 포함), 3=poison, 4=셋업 실패.
/// 실행 동안 프로젝트 실행 잠금을 쥐고(설계 §14.6), LogShim.BeginRun으로 runId를 알려 협조적 취소를 받는다(§14.4).
/// 부모가 사라지면 취소를 요청하고, 유예 안에 멈추지 않으면 격리 규칙으로 정리한 뒤 끝낸다(§14.5).
/// </summary>
public static class RunnerMain
{
    private static readonly JsonSerializerOptions MarkerJson = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>취소 요청 뒤 매크로가 LogShim 확인 지점에 닿기를 기다리는 시간(부모·러너 공통).</summary>
    public const int CancelGraceMs = 30_000;

    public static int Run(CliArgs opts)
    {
        try
        {
            return Execute(opts);
        }
        catch (CliError e)
        {
            PrintResult(new { ok = false, poison = false, error = e.Message, hint = e.Hint });
            return 4;
        }
        catch (Exception e)
        {
            PrintResult(new { ok = false, poison = false, error = $"{e.GetType().Name}: {e.Message}", hint = (string?)null });
            return 4;
        }
    }

    private static int Execute(CliArgs opts)
    {
        // 콘솔 Ctrl+C는 부모와 러너에 함께 온다 - 러너는 죽지 않고 취소만 요청한다(부모가 결과를 기다린다).
        // 세션 확보 전에 걸어 둔다: 시작 중의 Ctrl+C도 죽지 않고 실행 전 취소로 처리한다.
        string? flagForCtrlC = null;
        int ctrlC = 0;
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            Interlocked.Exchange(ref ctrlC, 1);
            if (Volatile.Read(ref flagForCtrlC) is string f) RunCancel.Request(f);
        };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        string runId = opts.RunId ?? Guid.NewGuid().ToString("N");
        string command = opts.Macro != null ? "run" : "test";
        var (app, meta) = SessionService.AcquireOrThrow(cfg, command, null, opts.SessionMeta);
        var filter = OleMessageFilter.Register();
        string cancelFlag = RunCancel.PathFor(meta.WorkbookFullName, runId);
        RunCancel.SweepOld(Path.GetDirectoryName(cancelFlag)!, TimeSpan.FromDays(1));
        var active = new ActiveRun
        {
            RunId = runId, Command = command, Target = opts.Macro ?? "Test_*",
            RunnerPid = Environment.ProcessId, RunnerStart = ProcessIdentity.GetStartTimeIso(Environment.ProcessId) ?? "",
            ExcelPid = meta.Pid, StartedUtc = DateTime.UtcNow.ToString("O"), CancelFlag = cancelFlag,
            Jsonl = Path.Combine(Path.GetDirectoryName(meta.WorkbookFullName)!, "_dbg",
                Path.GetFileNameWithoutExtension(meta.WorkbookFullName) + ".jsonl"),
        };
        using var runLock = RunLock.Hold(cfg, active);
        Volatile.Write(ref flagForCtrlC, cancelFlag);
        if (Volatile.Read(ref ctrlC) != 0) RunCancel.Request(cancelFlag);
        using var mainDone = new ManualResetEventSlim(false);
        var watchState = new StrongBox<int>(0); // 0 실행 중, 1 본 스레드 완료, 2 감시 스레드가 정리를 맡음
        StartParentWatch(opts.ParentPid, cancelFlag, mainDone, watchState, cfg, meta, command, runId);
        try
        {
            // 모달 워치독: 컴파일/런타임 오류 모달을 kill 없이 해제해 해당 테스트만 fail 처리.
            // 워치독 이벤트는 stdout/jsonl에 직접 쓰지 않고 큐로만 전달(프로토콜 보호).
            using var watchdog = new ModalWatchdog(meta.Pid);
            var watchdogEvents = new List<object>();
            // Excel 프로세스 소멸(크래시) 판정용 — 핸들을 실행 전에 쥐어야 소멸 뒤에도 종료 코드를 읽는다
            using var exitWatch = ExcelExitWatch.Open(meta.Pid, meta.ExcelStartTime);

            dynamic wb = SessionService.FindWorkbook(app, meta.WorkbookFullName);
            dynamic proj = SessionService.GetVbProject(wb);
            string wbName = Convert.ToString(wb.Name) ?? "";

            // 판독 계약(설계 §4): LogShim이 기록하는 <워크북.Path>\_dbg\<베이스명>.jsonl
            string wbDir = Path.GetDirectoryName(meta.WorkbookFullName)!;
            string jsonlPath = Path.Combine(wbDir, "_dbg",
                Path.GetFileNameWithoutExtension(meta.WorkbookFullName) + ".jsonl");

            var cancel = new CancelState(runId, runLock, active, cancelFlag);

            // --macro: 테스트 열거 대신 단건 매크로 실행 (run 명령의 자식 모드)
            if (opts.Macro != null)
                return RunMacro(opts.Macro, runId, cfg, app, meta, watchdog, exitWatch, watchdogEvents, wbName, jsonlPath, cancel);

            var testList = new List<(string Module, string Sub)>();
            List<VbaProjectOps.ComponentInfo> comps = VbaProjectOps.ListComponents(proj);
            foreach (var ci in comps.Where(c => c.Type == VbaProjectOps.TypeStdModule))
            {
                string body = VbaProjectOps.ReadModuleBody(ci.Comp.CodeModule);
                if (body.Length == 0) continue;
                foreach (var sub in TestEnumerator.FindTestSubs(CanonicalText.SplitLines(body)))
                    testList.Add((ci.Name, sub));
            }

            var tests = new List<object>();
            int pass = 0, fail = 0, skipped = 0;
            // 배치 지속 플래그: VBE 닫기 해제가 한 번이라도 있었으면 이후 RPC 단절은
            // poison이 아니라 세션 사망 후보 (drain 단위가 아닌 배치 단위 — save_prompt가
            // 이전 테스트 drain에 소비된 경우에도 유효해야 함)
            bool vbeCloseHappened = false;
            bool sessionDied = false;
            bool cancelled = false;
            ExcelExitKind excelExit = ExcelExitKind.Alive; // 프로세스 소멸(크래시/종료) — 모달 사망과 구분
            int next = 0; // 다음 미처리 테스트 인덱스 (사망 중단 시 잔여 skipped 채우기용)

            for (; next < testList.Count && !sessionDied; next++)
            {
                var (module, sub) = testList[next];
                string full = $"{module}.{sub}";
                if (File.Exists(cancelFlag))
                {
                    cancelled = true; // 테스트 사이에 온 취소 - 남은 테스트는 아래에서 skipped
                    break;
                }
                // 테스트의 End·런타임 오류 End는 LogShim의 runId까지 지운다 - 매 테스트 전에 다시 알린다
                cancel.Begin(app, wbName);
                AppendMarker(jsonlPath, "test_start", full, runId, null);
                long offset = File.Exists(jsonlPath) ? new FileInfo(jsonlPath).Length : 0;

                string? error = null;
                Exception? caught = null;
                watchdog.SetLabel(full);
                try
                {
                    // 항상 모듈 한정 호출 — 비한정 호출은 동명 프로시저 존재 시 1004 모호
                    ComRetry.Invoke(() => { app.Run($"'{wbName}'!{full}"); });
                }
                catch (COMException ex) when (ComRetry.IsPoison(ex.HResult)
                                              && !vbeCloseHappened && exitWatch.ClassifyAfter(ex) != ExcelExitKind.Alive)
                {
                    // 프로세스가 사라졌다 — poison이 아니라 크래시/종료이고 이 테스트가 진짜 신호(아래 판정에서 fail)
                    error = ex.Message;
                    caught = ex;
                }
                catch (COMException ex) when (ComRetry.IsPoison(ex.HResult))
                {
                    // 트리거 C: VBE 닫기 이후의 RPC 단절은 세션 사망의 다른 얼굴 — poison(exit 3,
                    // MarkPoisoned)이 아니라 clean stale로 처리, 이 테스트는 오탐이므로 skipped
                    if (vbeCloseHappened)
                    {
                        sessionDied = true;
                        skipped++;
                        AppendMarker(jsonlPath, "test_end", full, runId, RunnerVerdict.Skipped);
                        tests.Add(new
                        {
                            name = sub, module, result = RunnerVerdict.Skipped,
                            detail = $"not run because the session ended (error was a false positive: {ex.Message})",
                        });
                        next++;
                        break;
                    }
                    meta.MarkPoisoned($"test {full}: {ex.Message}", ex.HResult);
                    PrintResult(new
                    {
                        ok = false,
                        poison = true,
                        sessionDied = false,
                        error = $"test {full}: fatal COM error: {ex.Message}",
                        hint = "session start",
                        tests,
                        summary = new { pass, fail, skipped },
                        watchdog = watchdogEvents,
                    });
                    return 3;
                }
                catch (Exception ex)
                {
                    error = ex.Message; // VBA 런타임 에러 → 해당 테스트 fail
                    caught = ex;
                }
                finally
                {
                    watchdog.SetLabel(null);
                }

                // Run 반환은 워치독의 VBE 닫기 이후이므로 이 시점 큐에는 관련 이벤트가 있다
                var drained = watchdog.Drain();
                var myModals = drained.Where(e => e.Kind == "vbe_modal" && e.Label == full).ToList();
                foreach (var e in drained)
                    watchdogEvents.Add(ToEventPayload(e, staleFor: e.Kind == "vbe_modal" && e.Label != full ? full : null));

                bool vbeCloseInDrain = drained.Any(
                    e => e.Kind == "vbe_modal" && ModalHunter.RequiresVbeClose(e.Dismissal));
                if (vbeCloseInDrain) vbeCloseHappened = true;

                // 세션 사망 probe — 트리거 A: VBE 닫기 해제 발생(라벨 무관, Excel 종료가 지연되므로
                // 안정화 폴링). 트리거 B: non-poison 예외 발생(오탐 1004 가능성, 즉시 1회).
                // result 확정 전에 판정하므로 사후 재분류가 필요 없다.
                // 트리거 D: Excel 프로세스 자체가 사라짐(크래시·Quit). 크래시 의심 오류만 소멸을 기다리고 나머지는
                // 즉시 확인한다. 모달 사망(VBE 닫기)은 트리거 A 소관 — 정상 종료라 Quit으로 오인하지 않도록 제외.
                string? crashDetail = null;
                if (!vbeCloseInDrain)
                {
                    var kind = caught != null ? exitWatch.ClassifyAfter(caught) : exitWatch.Classify(0);
                    if (kind != ExcelExitKind.Alive)
                    {
                        excelExit = kind;
                        crashDetail = ExitDetail(kind, full, exitWatch);
                    }
                }
                bool deadNow = crashDetail != null;
                if (!deadNow && vbeCloseInDrain)
                    deadNow = ProbeDeathStabilized(app, meta.WorkbookFullName);
                else if (!deadNow && error != null)
                    deadNow = !SessionService.IsWorkbookAlive(app, meta.WorkbookFullName);

                var failedAsserts = ReadFailedAsserts(jsonlPath, offset);
                string? modalDetail = myModals.Count > 0
                    ? $"modal auto-dismissed ({myModals[0].Dismissal}): {Truncate(myModals[0].Body, 500)}"
                    : null;
                var (result, detail) = RunnerVerdict.Decide(error, failedAsserts, modalDetail, deadNow, crashDetail);
                if (!deadNow && SpanCancelled(jsonlPath, offset, runId))
                {
                    // LogShim이 End로 멈췄다 - 오류(0x800A9C68)는 취소의 흔적이지 테스트 실패가 아니다
                    cancelled = true;
                    result = RunnerVerdict.Skipped;
                    detail = "cancelled while running (stopped at a LogShim check point)";
                }
                if (result == RunnerVerdict.Pass) pass++;
                else if (result == RunnerVerdict.Fail) fail++;
                else skipped++;
                AppendMarker(jsonlPath, "test_end", full, runId, result);
                tests.Add(new { name = sub, module, result, detail });

                if (deadNow)
                {
                    sessionDied = true;
                    next++;
                    break;
                }
                if (cancelled)
                {
                    next++;
                    break;
                }
            }

            // 사망·취소 중단 시 잔여 테스트는 skipped로 채움 — 죽은 세션에 던져 오탐을 쌓지 않는다
            for (; (sessionDied || cancelled) && next < testList.Count; next++)
            {
                var (module, sub) = testList[next];
                skipped++;
                tests.Add(new
                {
                    name = sub, module, result = RunnerVerdict.Skipped,
                    detail = sessionDied ? "not run because the session ended" : "not run because the run was cancelled",
                });
            }

            // 마지막 해제의 3초 가드 윈도가 남았으면 저장 프롬프트 정리까지 대기 후 종료
            watchdog.LingerForGuardWindow();
            foreach (var e in watchdog.Drain())
            {
                if (e.Kind == "vbe_modal" && ModalHunter.RequiresVbeClose(e.Dismissal))
                    vbeCloseHappened = true;
                watchdogEvents.Add(ToEventPayload(e, staleFor: null));
            }

            // 마지막 테스트에서 모달이 났고 안정화 창(2s)을 넘겨 죽는 경우의 self-report
            if (vbeCloseHappened && !sessionDied)
                sessionDied = !SessionService.IsWorkbookAlive(app, meta.WorkbookFullName);

            PrintResult(new
            {
                ok = true,
                poison = false,
                sessionDied,
                cancelled,
                cancelSupported = cancel.Supported,
                excelExit = ExitName(excelExit),
                excelExitCode = excelExit != ExcelExitKind.Alive ? exitWatch.ExitCode : null,
                error = (string?)null,
                hint = (string?)null,
                tests,
                summary = new { pass, fail, skipped },
                watchdog = watchdogEvents,
            });
            return 0;
        }
        finally
        {
            // 감시 스레드가 이미 정리(poison·격리)를 맡았으면 그 스레드가 프로세스를 끝낸다 - 여기서 결과를 쓰거나
            // 잠금을 풀며 끼어들지 않는다
            if (Interlocked.CompareExchange(ref watchState.Value, 1, 0) == 2)
                Thread.Sleep(Timeout.Infinite);
            mainDone.Set();
            RunCancel.Clear(cancelFlag);
            filter.Revoke();
        }
    }

    /// <summary>협조적 취소 상태: LogShim.BeginRun 지원 여부를 실행 기록에 남겨 부모가 유예 여부를 정하게 한다.</summary>
    private sealed class CancelState(string runId, RunLock? runLock, ActiveRun record, string flag)
    {
        public bool? Supported { get; private set; }

        /// <summary>이미 취소가 요청됐는가(시작 중의 Ctrl+C, 부모 사망 등).</summary>
        public bool Requested => File.Exists(flag);

        /// <summary>LogShim.BeginRun 호출. 없는 구버전 LogShim(또는 LogShim 없음)이면 이후 호출하지 않는다.</summary>
        public void Begin(dynamic app, string wbName)
        {
            if (Supported == false) return;
            bool ok;
            try
            {
                ComRetry.Invoke(() => { app.Run($"'{wbName}'!LogShim.BeginRun", runId); });
                ok = true;
            }
            catch (Exception)
            {
                ok = false;
            }
            if (Supported != ok)
            {
                Supported = ok;
                record.CancelSupported = ok;
                runLock?.Update(record);
            }
        }
    }

    /// <summary>
    /// 부모 감시(설계 §14.5): 부모가 정리할 틈 없이 사라지면(강제 종료) 취소를 요청하고, 유예 안에 실행이 끝나지
    /// 않으면 부모가 하던 정리(poison 기록 + 격리 규칙)를 대신하고 끝낸다 — 매크로가 끝날 때까지 잠금을 쥔 채
    /// 아무도 결과를 읽지 않는 러너로 남지 않게.
    /// </summary>
    private static void StartParentWatch(int? parentPid, string cancelFlag, ManualResetEventSlim mainDone,
        StrongBox<int> watchState, ProjectConfig cfg, SessionMeta meta, string command, string runId)
    {
        if (parentPid is not int ppid) return;
        System.Diagnostics.Process? parent;
        try { parent = System.Diagnostics.Process.GetProcessById(ppid); }
        catch (ArgumentException) { parent = null; } // 이미 사라짐
        var thread = new Thread(() =>
        {
            ToolLog? log = null;
            try
            {
                parent?.WaitForExit();
                if (mainDone.IsSet) return;
                RunCancel.Request(cancelFlag);
                log = new ToolLog(cfg.ToolLogPath, command);
                log.Write("warn", "runner_parent_gone", new { parentPid = ppid, cancelFlag });
                if (mainDone.Wait(CancelGraceMs)) return;
            }
            catch
            {
                return; // 감시 실패 - 실행은 매크로가 끝나면 정상 종료한다
            }
            // 본 스레드와 경합하지 않게 정리를 한쪽이 맡는다(본 스레드가 방금 끝났으면 그쪽이 정상 종료한다)
            if (Interlocked.CompareExchange(ref watchState.Value, 2, 0) != 0) return;
            try
            {
                string reason = $"{command}: the vbaloop process that started it ended and the macro did not stop at a LogShim check point";
                meta.MarkPoisoned(reason, null);
                var (res, _, why) = RecoveryGate.Resolve(cfg, meta, "parent_gone", reason, command, log);
                log?.Write("error", "runner_abandoned", new { recovery = res.ToString(), why });
                RunLock.Forget(cfg, runId);
                if (ProcessIdentity.Identify(meta.Pid, meta.ExcelStartTime) == IdentityState.Gone)
                    RunCancel.Clear(cancelFlag); // Excel을 정리했다 - 멈출 매크로가 없다
            }
            catch
            {
                // 정리 실패도 끝낸다 - 잠금을 쥔 채 아무도 읽지 않는 러너로 남지 않게
            }
            // 본 스레드가 바쁜 Excel의 COM 호출에 묶여 있다 - 정상 종료는 RCW 정리에서 멈출 수 있다
            ComCleanup.HardExit(3);
        }) { IsBackground = true, Name = "vbaloop-parent-watch" };
        thread.Start();
    }

    /// <summary>offset 이후 구간에 이 runId의 LogShim 취소 레코드가 있는가(<see cref="RunCancel.SpanCancelled"/>).</summary>
    public static bool SpanCancelled(string jsonlPath, long offset, string runId)
        => RunCancel.SpanCancelled(jsonlPath, offset, runId);

    public static bool IsCancelRecord(string line, string runId) => RunCancel.IsCancelRecord(line, runId);

    /// <summary>
    /// 워크북 사망 안정화 probe: VBE 닫기 해제 후 Excel 종료는 지연되므로(supervisor는 800ms
    /// 단발) 250ms 간격 최대 2초 폴링. 죽음 확인 즉시 true, 창 내내 살아있으면 false —
    /// VBE 닫기가 항상 죽이는 것은 아니다(저장 프롬프트 취소 성공 경로).
    /// </summary>
    private static bool ProbeDeathStabilized(dynamic app, string workbookFullName, int maxMs = 2000)
    {
        long deadline = Environment.TickCount64 + maxMs;
        while (Environment.TickCount64 < deadline)
        {
            if (!SessionService.IsWorkbookAlive(app, workbookFullName)) return true;
            Thread.Sleep(250);
        }
        return !SessionService.IsWorkbookAlive(app, workbookFullName);
    }

    /// <summary>단건 매크로 실행(설계 §8 run). 판정: 런타임 에러 없음 AND 모달 없음 = ok.
    /// 어서션 실패는 run의 성패를 가르지 않고 정보로만 노출(테스트가 아니므로).</summary>
    private static int RunMacro(
        string macroSpec, string runId, ProjectConfig cfg, dynamic app, SessionMeta meta,
        ModalWatchdog watchdog, ExcelExitWatch exitWatch, List<object> watchdogEvents, string wbName, string jsonlPath,
        CancelState cancel)
    {
        if (!MacroName.TryParse(macroSpec, out var module, out var proc, out var nameError))
            throw new CliError(nameError!, "Format: vbaloop run Module.Proc");
        string full = $"{module}.{proc}";

        cancel.Begin(app, wbName);
        AppendMarker(jsonlPath, "run_start", full, runId, null);
        long offset = File.Exists(jsonlPath) ? new FileInfo(jsonlPath).Length : 0;
        bool cancelledBeforeStart = cancel.Requested; // 시작 중에 온 취소 - 매크로를 부르지 않는다

        string? error = null;
        Exception? caught = null;
        object? rawReturn = null;
        var sw = System.Diagnostics.Stopwatch.StartNew();
        watchdog.SetLabel(full);
        try
        {
            // 항상 모듈 한정 호출
            if (!cancelledBeforeStart)
                rawReturn = ComRetry.Invoke<object?>(() => app.Run($"'{wbName}'!{full}"));
        }
        catch (COMException ex) when (ComRetry.IsPoison(ex.HResult) && exitWatch.ClassifyAfter(ex) != ExcelExitKind.Alive)
        {
            // 프로세스가 사라졌다 — poison이 아니라 크래시/종료(아래에서 sessionDied로 보고)
            error = ex.Message;
            caught = ex;
        }
        catch (COMException ex) when (ComRetry.IsPoison(ex.HResult))
        {
            meta.MarkPoisoned($"run {full}: {ex.Message}", ex.HResult);
            PrintResult(new
            {
                ok = false,
                poison = true,
                sessionDied = false,
                error = $"run {full}: fatal COM error: {ex.Message}",
                hint = "session start",
                watchdog = watchdogEvents,
            });
            return 3;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            caught = ex;
        }
        finally
        {
            sw.Stop();
            watchdog.SetLabel(null);
        }

        var drained = watchdog.Drain();
        var myModals = drained.Where(e => e.Kind == "vbe_modal" && e.Label == full).ToList();
        foreach (var e in drained)
            watchdogEvents.Add(ToEventPayload(e, staleFor: e.Kind == "vbe_modal" && e.Label != full ? full : null));
        bool vbeCloseHappened = drained.Any(
            e => e.Kind == "vbe_modal" && ModalHunter.RequiresVbeClose(e.Dismissal));

        // Excel 프로세스 자체가 사라졌는지(크래시·Quit) — 모달 사망(VBE 닫기)은 아래 안정화 probe 소관
        var excelExit = vbeCloseHappened ? ExcelExitKind.Alive
            : caught != null ? exitWatch.ClassifyAfter(caught) : exitWatch.Classify(0);
        string? crashDetail = ExitDetail(excelExit, full, exitWatch);

        var (records, failedAsserts) = ReadSpan(jsonlPath, offset);
        bool modalHit = myModals.Count > 0;
        // LogShim이 End로 멈춘 취소 - 오류(0x800A9C68)는 취소의 흔적이다
        bool cancelled = cancelledBeforeStart
            || (crashDetail == null && !modalHit && SpanCancelled(jsonlPath, offset, runId));
        string result = cancelled ? "cancelled"
            : error == null && !modalHit && crashDetail == null ? "ok" : "fail";
        AppendMarker(jsonlPath, "run_end", full, runId, result);

        string? modalDetail = modalHit
            ? $"modal auto-dismissed ({myModals[0].Dismissal}): {Truncate(myModals[0].Body, 500)}"
            : null;
        string? detail = cancelledBeforeStart ? "cancelled before the macro started"
            : cancelled ? "cancelled at a LogShim check point (the runner asked the macro to stop)"
            : crashDetail != null
            ? crashDetail + (modalDetail != null ? $" | {modalDetail}" : "") + (error != null ? $" | COM: {error}" : "")
            : modalDetail != null ? modalDetail + (error != null ? $" | COM: {error}" : "")
            : error != null ? $"runtime error: {error}" : null;

        var (returnValue, returnType) = SerializeReturn(rawReturn);

        watchdog.LingerForGuardWindow();
        foreach (var e in watchdog.Drain())
        {
            if (e.Kind == "vbe_modal" && ModalHunter.RequiresVbeClose(e.Dismissal))
                vbeCloseHappened = true;
            watchdogEvents.Add(ToEventPayload(e, staleFor: null));
        }

        // test 모드와 대칭: VBE 닫기 해제가 있었으면 안정화 probe로 사망을 self-report
        bool sessionDied = excelExit != ExcelExitKind.Alive
                           || (vbeCloseHappened && ProbeDeathStabilized(app, meta.WorkbookFullName));

        PrintResult(new
        {
            ok = result == "ok",
            poison = false,
            sessionDied,
            cancelled,
            cancelSupported = cancel.Supported,
            excelExit = ExitName(excelExit),
            excelExitCode = excelExit != ExcelExitKind.Alive ? exitWatch.ExitCode : null,
            error = result == "ok" ? null : detail,
            hint = (string?)null,
            run = new
            {
                macro = full,
                module,
                proc,
                result,
                returnValue,
                returnType,
                elapsedMs = sw.ElapsedMilliseconds,
                detail,
                failedAsserts = failedAsserts.Count,
                failedAssertSamples = failedAsserts.Take(5),
                records,
            },
            watchdog = watchdogEvents,
        });
        return 0;
    }

    /// <summary>프로세스 소멸 종류별 실패 detail. 살아 있으면 null.</summary>
    private static string? ExitDetail(ExcelExitKind kind, string full, ExcelExitWatch w) => kind switch
    {
        ExcelExitKind.Crashed => ExcelCrash.Detail(full, w.Pid, w.ExitCode),
        ExcelExitKind.Exited => ExcelCrash.ExitedDetail(full, w.Pid),
        _ => null,
    };

    /// <summary>러너 JSON의 excelExit 값(supervisor가 읽는다). 살아 있으면 null.</summary>
    private static string? ExitName(ExcelExitKind kind) => kind switch
    {
        ExcelExitKind.Crashed => "crashed",
        ExcelExitKind.Exited => "exited",
        _ => null,
    };

    /// <summary>Application.Run 반환값 직렬화 — 원시 타입만. COM 객체는 러너가 GC 정리
    /// 경로 밖이므로 즉시 해제하고 타입명만 보고(검토 지적: RCW 방치 금지).</summary>
    private static (object? Value, string Type) SerializeReturn(object? raw)
    {
        if (raw == null || ReferenceEquals(raw, System.Reflection.Missing.Value))
            return (null, "Empty");
        switch (raw)
        {
            case bool or byte or sbyte or short or ushort or int or uint or long or ulong
                 or float or double or decimal or string:
                return (raw, raw.GetType().Name);
            case DateTime dt:
                return (dt.ToString("yyyy-MM-dd HH:mm:ss"), "Date");
        }
        if (Marshal.IsComObject(raw))
        {
            Marshal.ReleaseComObject(raw);
            return (null, "Object(COM) — only primitive types can be returned");
        }
        if (raw is Array arr)
            return ($"Array(length {arr.Length})", "Array");
        return (raw.ToString(), raw.GetType().Name);
    }

    private static object ToEventPayload(ModalWatchdog.ModalEvent e, string? staleFor) => new
    {
        ts = e.Ts,
        test = e.Label,
        kind = e.Kind,
        caption = e.Caption,
        body = Truncate(e.Body, 500),
        dismissal = e.Dismissal,
        // 현재 테스트 시작 전에 남아 있던 이전 구간의 잔여 모달임을 표기 (귀속 혼동 방지)
        stale = staleFor != null ? $"leftover modal before {staleFor} started" : null,
    };

    private static string Truncate(string s, int max)
        => s.Length <= max ? s : s[..max] + "...";

    /// <summary>테스트 전후 마커 레코드를 같은 jsonl에 append해 구간을 구분(설계 §4 test).</summary>
    private static void AppendMarker(string jsonlPath, string @event, string test, string runId, string? result)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(jsonlPath)!);
        var rec = new Dictionary<string, object?>
        {
            ["ts"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
            ["kind"] = "X",
            ["event"] = @event,
            ["test"] = test,
            ["runId"] = runId,
        };
        if (result != null) rec["result"] = result;
        string line = JsonSerializer.Serialize(rec, MarkerJson);
        using var fs = new FileStream(jsonlPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
        var bytes = CanonicalText.Utf8NoBom.GetBytes(line + "\n");
        fs.Write(bytes, 0, bytes.Length);
    }

    /// <summary>offset 이후에 append된 레코드 중 kind:"A" ok:false를 수집(판독 계약).</summary>
    private static List<string> ReadFailedAsserts(string jsonlPath, long offset)
        => ReadSpan(jsonlPath, offset).FailedAsserts;

    /// <summary>offset 이후 구간의 전체 레코드(상한 cap)와 실패 어서션을 함께 수집(run payload용).</summary>
    private static (List<JsonElement> Records, List<string> FailedAsserts) ReadSpan(
        string jsonlPath, long offset, int cap = 200)
    {
        var records = new List<JsonElement>();
        var failed = new List<string>();
        if (!File.Exists(jsonlPath)) return (records, failed);

        using var fs = new FileStream(jsonlPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        if (offset >= fs.Length) return (records, failed);
        fs.Seek(offset, SeekOrigin.Begin);
        using var reader = new StreamReader(fs, CanonicalText.Utf8NoBom);
        string? line;
        while ((line = reader.ReadLine()) != null)
        {
            line = line.Trim();
            if (line.Length == 0 || !line.StartsWith('{')) continue;
            try
            {
                using var doc = JsonDocument.Parse(line);
                var root = doc.RootElement;
                if (records.Count < cap) records.Add(root.Clone());
                if (root.TryGetProperty("kind", out var kind) && kind.GetString() == "A" &&
                    root.TryGetProperty("ok", out var ok) && ok.ValueKind == JsonValueKind.False)
                {
                    string name = root.TryGetProperty("name", out var n) ? n.GetString() ?? "?" : "?";
                    string exp = root.TryGetProperty("expected", out var e) ? e.ToString() : "?";
                    string act = root.TryGetProperty("actual", out var a) ? a.ToString() : "?";
                    failed.Add($"{name}: expected={exp} actual={act}");
                }
            }
            catch (JsonException)
            {
                // 깨진 라인은 무시 (판정은 A 레코드만 사용)
            }
        }
        return (records, failed);
    }

    private static void PrintResult(object result)
    {
        try
        {
            Console.WriteLine(JsonSerializer.Serialize(result, Output.JsonOptions));
        }
        catch (IOException)
        {
            // 부모가 사라져 파이프가 닫혔다 - 결과를 읽을 쪽이 없다
        }
    }
}
