using System.Runtime.InteropServices;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// 도구 인스턴스의 Workbooks.Open 단일 창구. Excel 복원력 기록("마지막으로 'X'을(를) 열었을 때 오류가
/// 발생했습니다. 그래도 이 문서를 여시겠습니까?")이 경로에 묶이면 DisplayAlerts=False가 기본 버튼
/// "아니요"로 자동 응답해 Open이 0x800A03EC로만 실패한다(실측 2026-11-02 — 파일은 정상, 같은 파일의
/// 다른 이름 사본은 열림, DisabledItems 값을 지워도 계속 뜸). 그 경우 한 번만 알림을 켜고 재시도해 이
/// 프롬프트에만 "예"로 답한다 — "예" 한 번이 기록을 지워 이후 열기는 알림 없이 성공한다(실측).
/// 이벤트·매크로가 꺼진 숨은 도구 인스턴스라 다시 크래시해도 잃을 것이 없다.
/// </summary>
public static class WorkbookOpener
{
    /// <summary>Excel 1004("Workbooks 클래스 중 Open 속성을 구할 수 없습니다") — 알림이 거절된 경우 포함.</summary>
    public const int XlOpenFailed = unchecked((int)0x800A03EC);

    /// <summary>재열기 확인 프롬프트의 숨은 알림 ID 정적 텍스트(한국어 Excel 16 실측). 언어 중립 신호 —
    /// 같은 예/아니요 구조의 "손상된 파일 복구?" 프롬프트와 구분하는 유일한 근거다(복구본을 push가
    /// 저장하면 정본이 바뀐다).</summary>
    public const string CrashReopenAlertId = "700164";

    public const int IDYES = 6;

    /// <summary>Open 결과. AcceptedPrompt는 "예"로 답한 프롬프트 본문(없으면 null).</summary>
    public sealed record Result(dynamic Workbook, string? AcceptedPrompt);

    public static Result Open(dynamic app, int pid, string path, bool readOnly, ToolLog log)
    {
        COMException first;
        try
        {
            // 권장은 쓰기 보호가 아니다. 알림을 끈 Excel의 기본 응답으로 쓰기 요청이
            // 읽기 전용으로 바뀌지 않게 하되, 호출자가 명시한 ReadOnly 값은 유지한다.
            return new Result(app.Workbooks.Open(path, 0, readOnly, IgnoreReadOnlyRecommended: true), null); // UpdateLinks:=0
        }
        catch (COMException ex) when (ex.HResult == XlOpenFailed)
        {
            first = ex;
        }
        catch (COMException ex) when (ComRetry.IsPoison(ex.HResult))
        {
            throw DiedWhileOpening(path, pid, ex, log);
        }

        // 거절된 알림이 무엇인지 보려면 알림을 켜야 한다. 감시 스레드가 재열기 프롬프트에만 "예",
        // 그 밖의 다이얼로그는 ModalHunter 안전 버튼(확인/아니요/취소)으로 닫아 기존 거절과 같게 둔다.
        string fileName = Path.GetFileName(path);
        string? accepted = null;
        var others = new List<string>();
        using var stop = new CancellationTokenSource();
        var watcher = new Thread(() =>
        {
            while (!stop.IsCancellationRequested)
            {
                foreach (var dlg in ModalHunter.FindDialogs(pid))
                {
                    if (dlg.ClassName != "#32770") continue;
                    var buttons = ModalHunter.SnapshotButtons(dlg.Hwnd);
                    if (buttons.Count == 0) continue; // 생성 중 — 다음 폴
                    string info = ModalHunter.CollectInfo(dlg.Hwnd);
                    if (accepted == null && IsCrashReopenPrompt(buttons.Select(b => b.Info).ToList(), info, fileName))
                    {
                        var yes = buttons.First(b => b.Info.CtrlId == IDYES);
                        _ = NativeMethods.SendMessageTimeoutW(yes.Hwnd, NativeMethods.BM_CLICK, IntPtr.Zero,
                            IntPtr.Zero, NativeMethods.SMTO_ABORTIFHUNG | NativeMethods.SMTO_NORMAL, 2000, out _);
                        accepted = info;
                    }
                    else
                    {
                        var r = ModalHunter.Dismiss(dlg);
                        others.Add($"{r?.Dismissal}: {info}");
                    }
                    ModalHunter.WaitUntilGone(dlg.Hwnd, 1500);
                }
                stop.Token.WaitHandle.WaitOne(100);
            }
        }) { IsBackground = true, Name = "open-prompt-watcher" };

        dynamic wb;
        app.DisplayAlerts = true;
        watcher.Start();
        try
        {
            wb = app.Workbooks.Open(path, 0, readOnly, IgnoreReadOnlyRecommended: true);
        }
        catch (COMException ex) when (ComRetry.IsPoison(ex.HResult))
        {
            throw DiedWhileOpening(path, pid, ex, log);
        }
        catch (COMException ex)
        {
            log.Write("error", "workbook_open_failed",
                new { workbook = path, hResult = $"0x{ex.HResult:X8}", firstHResult = $"0x{first.HResult:X8}", dialogs = others });
            throw new CliError(
                $"Excel refused to open the workbook (0x{ex.HResult:X8}): {path}"
                + (others.Count > 0 ? $" — dialog: {string.Join(" / ", others)}" : ""),
                "Ask the user to open the workbook in Excel by hand: a corrupt file fails there too, while a prompt about"
                + " an error the last time it was opened means Excel blocked this path (answer Yes, then retry)",
                environment: true);
        }
        finally
        {
            stop.Cancel();
            watcher.Join(3000);
            try { app.DisplayAlerts = false; } catch { /* 실패 경로에서 인스턴스가 죽었을 수 있다 */ }
        }

        if (accepted != null)
            log.Write("warn", "workbook_open_prompt_accepted", new { workbook = path, prompt = accepted });
        if (others.Count > 0)
            log.Write("warn", "workbook_open_dialogs_dismissed", new { workbook = path, dialogs = others });
        return new Result(wb, accepted);
    }

    /// <summary>
    /// 재열기 확인 프롬프트 판정(순수 로직). 예(IDYES)·아니요(IDNO)만 있고, 숨은 알림 ID
    /// <see cref="CrashReopenAlertId"/> 텍스트와 대상 파일명을 함께 담은 다이얼로그만 해당 —
    /// 본문 문구는 지역화되므로 쓰지 않는다. info는 <see cref="ModalHunter.CollectInfo"/> 형식.
    /// </summary>
    public static bool IsCrashReopenPrompt(IReadOnlyList<ModalHunter.ButtonInfo> buttons, string info, string fileName)
    {
        var usable = buttons.Where(b => b.Enabled && b.Visible && b.CtrlId != ModalHunter.IDHELP).ToList();
        if (usable.Count != 2
            || !usable.Any(b => b.CtrlId == IDYES)
            || !usable.Any(b => b.CtrlId == ModalHunter.IDNO))
            return false;
        bool hasAlertId = info.Split(" | ").Any(t => t.Trim() == CrashReopenAlertId);
        return hasAlertId && info.Contains(fileName, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Open 도중 Excel 프로세스가 사라짐(RPC 단절) — 크래시. 내부 오류가 아니라 환경 오류다.</summary>
    private static CliError DiedWhileOpening(string path, int pid, COMException ex, ToolLog log)
    {
        log.Write("error", "workbook_open_excel_died", new { workbook = path, pid, hResult = $"0x{ex.HResult:X8}" });
        return new CliError(
            $"Excel (PID {pid}) exited while opening the workbook (0x{ex.HResult:X8}) — Excel crashed: {path}",
            "Retry once; if it repeats, check whether the workbook opens in Excel by hand", environment: true);
    }
}
