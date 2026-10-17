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
/// 출력: stdout에 JSON 한 덩어리. exit 0=실행 완료(실패 포함), 3=poison, 4=셋업 실패.
/// </summary>
public static class RunnerMain
{
    private static readonly JsonSerializerOptions MarkerJson = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

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
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        string runId = opts.RunId ?? Guid.NewGuid().ToString("N");
        var (app, meta) = SessionService.AcquireOrThrow(cfg, opts.Macro != null ? "run" : "test", null, opts.SessionMeta);
        var filter = OleMessageFilter.Register();
        try
        {
            // 모달 워치독: 컴파일/런타임 오류 모달을 kill 없이 해제해 해당 테스트만 fail 처리.
            // 워치독 이벤트는 stdout/jsonl에 직접 쓰지 않고 큐로만 전달(프로토콜 보호).
            using var watchdog = new ModalWatchdog(meta.Pid);
            var watchdogEvents = new List<object>();

            dynamic wb = SessionService.FindWorkbook(app, meta.WorkbookFullName);
            dynamic proj = SessionService.GetVbProject(wb);
            string wbName = Convert.ToString(wb.Name) ?? "";

            // 판독 계약(설계 §4): LogShim이 기록하는 <워크북.Path>\_dbg\<베이스명>.jsonl
            string wbDir = Path.GetDirectoryName(meta.WorkbookFullName)!;
            string jsonlPath = Path.Combine(wbDir, "_dbg",
                Path.GetFileNameWithoutExtension(meta.WorkbookFullName) + ".jsonl");

            // --macro: 테스트 열거 대신 단건 매크로 실행 (run 명령의 자식 모드)
            if (opts.Macro != null)
                return RunMacro(opts.Macro, runId, cfg, app, meta, watchdog, watchdogEvents, wbName, jsonlPath);

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
            int next = 0; // 다음 미처리 테스트 인덱스 (사망 중단 시 잔여 skipped 채우기용)

            for (; next < testList.Count && !sessionDied; next++)
            {
                var (module, sub) = testList[next];
                string full = $"{module}.{sub}";
                AppendMarker(jsonlPath, "test_start", full, runId, null);
                long offset = File.Exists(jsonlPath) ? new FileInfo(jsonlPath).Length : 0;

                string? error = null;
                watchdog.SetLabel(full);
                try
                {
                    // 항상 모듈 한정 호출 — 비한정 호출은 동명 프로시저 존재 시 1004 모호
                    ComRetry.Invoke(() => { app.Run($"'{wbName}'!{full}"); });
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
                            detail = $"세션 종료로 미실행(에러는 오탐: {ex.Message})",
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
                        error = $"test {full}: 치명적 COM 오류: {ex.Message}",
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
                bool deadNow = false;
                if (vbeCloseInDrain)
                    deadNow = ProbeDeathStabilized(app, meta.WorkbookFullName);
                else if (error != null)
                    deadNow = !SessionService.IsWorkbookAlive(app, meta.WorkbookFullName);

                var failedAsserts = ReadFailedAsserts(jsonlPath, offset);
                string? modalDetail = myModals.Count > 0
                    ? $"모달 자동 해제({myModals[0].Dismissal}): {Truncate(myModals[0].Body, 500)}"
                    : null;
                var (result, detail) = RunnerVerdict.Decide(error, failedAsserts, modalDetail, deadNow);
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
            }

            // 사망 중단 시 잔여 테스트는 skipped로 채움 — 죽은 세션에 던져 오탐을 쌓지 않는다
            for (; sessionDied && next < testList.Count; next++)
            {
                var (module, sub) = testList[next];
                skipped++;
                tests.Add(new
                {
                    name = sub, module, result = RunnerVerdict.Skipped, detail = "세션 종료로 미실행",
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
            filter.Revoke();
        }
    }

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
        ModalWatchdog watchdog, List<object> watchdogEvents, string wbName, string jsonlPath)
    {
        if (!MacroName.TryParse(macroSpec, out var module, out var proc, out var nameError))
            throw new CliError(nameError!, "형식: vbaloop run Module.Proc");
        string full = $"{module}.{proc}";

        AppendMarker(jsonlPath, "run_start", full, runId, null);
        long offset = File.Exists(jsonlPath) ? new FileInfo(jsonlPath).Length : 0;

        string? error = null;
        object? rawReturn = null;
        var sw = System.Diagnostics.Stopwatch.StartNew();
        watchdog.SetLabel(full);
        try
        {
            // 항상 모듈 한정 호출
            rawReturn = ComRetry.Invoke<object?>(() => app.Run($"'{wbName}'!{full}"));
        }
        catch (COMException ex) when (ComRetry.IsPoison(ex.HResult))
        {
            meta.MarkPoisoned($"run {full}: {ex.Message}", ex.HResult);
            PrintResult(new
            {
                ok = false,
                poison = true,
                sessionDied = false,
                error = $"run {full}: 치명적 COM 오류: {ex.Message}",
                hint = "session start",
                watchdog = watchdogEvents,
            });
            return 3;
        }
        catch (Exception ex)
        {
            error = ex.Message;
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

        var (records, failedAsserts) = ReadSpan(jsonlPath, offset);
        bool modalHit = myModals.Count > 0;
        string result = error == null && !modalHit ? "ok" : "fail";
        AppendMarker(jsonlPath, "run_end", full, runId, result);

        string? detail = modalHit
            ? $"모달 자동 해제({myModals[0].Dismissal}): {Truncate(myModals[0].Body, 500)}"
              + (error != null ? $" | COM: {error}" : "")
            : error != null ? $"런타임 에러: {error}" : null;

        var (returnValue, returnType) = SerializeReturn(rawReturn);

        watchdog.LingerForGuardWindow();
        foreach (var e in watchdog.Drain())
        {
            if (e.Kind == "vbe_modal" && ModalHunter.RequiresVbeClose(e.Dismissal))
                vbeCloseHappened = true;
            watchdogEvents.Add(ToEventPayload(e, staleFor: null));
        }

        // test 모드와 대칭: VBE 닫기 해제가 있었으면 안정화 probe로 사망을 self-report
        bool sessionDied = vbeCloseHappened && ProbeDeathStabilized(app, meta.WorkbookFullName);

        PrintResult(new
        {
            ok = result == "ok",
            poison = false,
            sessionDied,
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
            return (null, "Object(COM) — 원시 타입만 반환 가능");
        }
        if (raw is Array arr)
            return ($"Array(길이 {arr.Length})", "Array");
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
        stale = staleFor != null ? $"{staleFor} 시작 전 잔여 모달" : null,
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
        Console.WriteLine(JsonSerializer.Serialize(result, Output.JsonOptions));
    }
}
