using System.Diagnostics;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// VBE 컴파일 검증(설계 §9.8, push 컴파일 게이트 §4): 프로젝트를 ActiveVBProject로 잡고
/// "VBAProject 컴파일"(Id 578)을 Execute. 컴파일 오류는 VBE 모달로만 드러나므로 ModalWatchdog(test
/// 러너와 같은 해제 시퀀스)이 잡아 본문을 수집하고 닫는다. 모달 해제가 실패해 Execute가 영영 안
/// 돌아오는 경우는 타임아웃 kill로 끊는다 — 그래서 호출자는 도구가 띄운 인스턴스(dist 임시 Excel,
/// push의 일회용·유지 세션)에서만 부르고, 사용자 Excel(attach)에는 부르지 않는다. 판정은 순수 DistCompileVerdict가 한다.
/// </summary>
public static class VbeCompile
{
    public const int TimeoutSec = 120;

    /// <param name="logPrefix">로그 이벤트 접두어("dist", "push").</param>
    public static DistCompileVerdict.Verdict Run(dynamic app, dynamic proj, int pid, ToolLog log, string logPrefix)
    {
        dynamic vbe = app.VBE;
        try { vbe.ActiveVBProject = proj; }
        catch (Exception ex) { log.Write("warn", $"{logPrefix}_compile_activate_failed", new { error = ex.Message }); }

        dynamic? ctl = null;
        try { ctl = vbe.CommandBars.FindControl(Type.Missing, DistCompileVerdict.CompileControlId); }
        catch (Exception ex) { log.Write("warn", $"{logPrefix}_compile_findcontrol_failed", new { error = ex.Message }); }
        if (ctl == null)
            return DistCompileVerdict.Decide(false, false, Array.Empty<string>(), null, null, timedOut: false);

        bool enabledBefore = ComRetry.Invoke<bool>(() => (bool)Convert.ToBoolean(ctl.Enabled));
        if (!enabledBefore)
        {
            log.Write("info", $"{logPrefix}_compile_already", null);
            return DistCompileVerdict.Decide(true, false, Array.Empty<string>(), null, null, timedOut: false);
        }

        using var watchdog = new ModalWatchdog(pid, DeferVbeCloseMs);
        watchdog.SetLabel($"{logPrefix}:compile");
        CompileErrorLocation? location = null;
        using var done = new ManualResetEventSlim(false);
        bool killed = false;
        var killer = new Thread(() =>
        {
            if (!done.Wait(TimeoutSec * 1000))
            {
                killed = true;
                try { Process.GetProcessById(pid).Kill(entireProcessTree: true); } catch { /* 이미 죽음 */ }
            }
        }) { IsBackground = true };
        killer.Start();

        string? executeError = null;
        bool? enabledAfter = null;
        try
        {
            try { ComRetry.Invoke(() => { ctl.Execute(); }); }
            catch (Exception ex) { executeError = ex.Message; }
            if (!killed)
            {
                // Execute가 명령을 비동기로 게시하는 경우 대비: 모달이 뜰 여지를 잠깐 주고 상태를 다시 읽는다.
                // 모달이 떠 있으면 이 호출은 거부→재시도되는 동안 워치독이 모달을 닫는다.
                Thread.Sleep(500);
                try { enabledAfter = ComRetry.Invoke<bool>(() => (bool)Convert.ToBoolean(ctl.Enabled)); }
                catch (Exception ex) { log.Write("warn", $"{logPrefix}_compile_state_unreadable", new { error = ex.Message }); }

                // 오류 위치: 워치독이 모달을 닫고 VBE 닫기를 미룬 사이에 선택 영역을 읽는다. 위 Enabled 읽기는
                // 모달이 떠 있는 동안 거부-재시도되므로 여기서는 해제 직후다 — 미룬 닫기 설정까지만 잠깐 기다린다.
                if (watchdog.SawVbaModal && watchdog.WaitForDeferredVbeClose(3000))
                    location = ReadErrorLocation(vbe, log, logPrefix);
            }
        }
        finally
        {
            done.Set();
            watchdog.CloseDeferredVbe(); // 읽었든 못 읽었든 이제 닫는다 — 아래 링거가 후속 저장 프롬프트를 드레인
        }

        var events = watchdog.Drain();
        if (events.Count > 0 && !killed) watchdog.LingerForGuardWindow(); // VBE 닫기 후속 저장 프롬프트 드레인
        events.AddRange(watchdog.Drain());
        var modalTexts = events.Where(e => e.Kind == "vbe_modal").Select(e => e.Body).ToList();

        var verdict = DistCompileVerdict.Decide(true, true, modalTexts, executeError, enabledAfter, killed);
        if (verdict.Kind == DistCompileVerdict.Kind.CompileError) verdict = verdict with { Location = location };
        log.Write(verdict.Publishable ? "info" : "error", $"{logPrefix}_compile", new
        {
            result = verdict.Label, detail = verdict.Detail, location = verdict.Location,
            modals = events.Select(e => new { e.Kind, e.Body, e.Dismissal }),
        });
        return verdict;
    }

    /// <summary>미룬 VBE 닫기 기한 — 호출 스레드가 이 안에 위치를 못 읽으면 워치독이 닫는다(실측 판독 ~30ms).</summary>
    public const int DeferVbeCloseMs = 5000;

    /// <summary>
    /// 컴파일 오류 직후 VBE 선택 영역 = 오류 위치(모달 본문에는 없음). 메인 스레드의 기존 vbe 프록시로 읽는다 —
    /// Execute는 비동기 게시라 즉시 반환되므로 이 스레드는 막혀 있지 않다(2026-10-31 실측). 실패는 null(보고만 빠짐).
    /// </summary>
    private static CompileErrorLocation? ReadErrorLocation(dynamic vbe, ToolLog log, string logPrefix)
    {
        try
        {
            dynamic? pane = ComRetry.Invoke<object?>(() => vbe.ActiveCodePane);
            if (pane == null) return null;
            int sl = 0, sc = 0, el = 0, ec = 0;
            ComRetry.Invoke(() => { pane.GetSelection(ref sl, ref sc, ref el, ref ec); });
            if (sl <= 0) return null;
            dynamic cm = pane.CodeModule;
            string module = Convert.ToString(cm.Parent.Name) ?? "";
            string text = Convert.ToString(cm.Lines(sl, 1)) ?? "";
            return new CompileErrorLocation(module, sl, sc, el, ec, text);
        }
        catch (Exception ex)
        {
            log.Write("warn", $"{logPrefix}_compile_location_unreadable", new { error = ex.Message });
            return null;
        }
    }
}
