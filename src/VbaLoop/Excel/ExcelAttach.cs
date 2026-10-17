using System.Runtime.InteropServices;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// 재attach 절차(설계 §4 공통 계약): 메타파일 HWND(XLMAIN)의 EXCEL7 자식 창에
/// AccessibleObjectFromWindow(OBJID_NATIVEOM) → .Application 획득 → PID+StartTime 검증.
/// GetActiveObject는 사용 금지(§6 금지 목록) — ROT 임의 인스턴스 반환으로 특정 불가.
/// xlFlow ExcelBridgeSupport.TryGetExcelByProcessId의 최상위 견고 티어만 이식.
/// </summary>
public static class ExcelAttach
{
    /// <summary>
    /// Busy=true는 Excel이 살아있고 우리 프로세스가 맞지만 호출을 거절 중(RPC_E_CALL_REJECTED 계열 —
    /// 셀 편집 중, 모달 대화상자)이라는 뜻이다. stale이 아니므로 정리/격리 대상이 아니라 재시도 대상(설계 §11).
    /// </summary>
    public sealed record AttachResult(dynamic? App, string? FailReason, bool Busy = false);

    /// <summary>
    /// busy 재시도 창. 워치독 직후의 짧은 경합(실측 1초 미만)은 흡수하고, 그 이상(사람이 셀을
    /// 편집 중)은 Busy로 보고해 호출자가 exit 3 재시도로 돌려준다.
    /// </summary>
    private const int BusyRetries = 10;
    private const int BusyDelayMs = 300;

    public static AttachResult TryAttach(SessionMeta meta)
    {
        if (!ProcessIdentity.IsAlive(meta.Pid, meta.ExcelStartTime))
            return new AttachResult(null, $"PID {meta.Pid} 프로세스가 없거나 StartTime 불일치(PID 재사용)");
        if (meta.IsExternal)
            return TryAttachExternal(meta);

        var xlMain = new IntPtr(meta.Hwnd);
        if (!NativeMethods.IsWindow(xlMain))
            return new AttachResult(null, $"HWND 0x{meta.Hwnd:X}이 유효한 창이 아님");

        NativeMethods.GetWindowThreadProcessId(xlMain, out uint windowPid);
        if (windowPid != (uint)meta.Pid)
            return new AttachResult(null, $"HWND 소유 PID({windowPid})가 메타파일 PID({meta.Pid})와 불일치");

        var excel7 = NativeMethods.FindExcel7Child(xlMain);
        if (excel7 == IntPtr.Zero)
            return new AttachResult(null, "EXCEL7 자식 창을 찾지 못함(워크북 창 없음)");

        var att = ApplicationFromExcel7(excel7, meta.Pid, BusyRetries);
        if (att.App == null) return att;

        long appHwnd;
        try
        {
            appHwnd = Convert.ToInt64(att.App.Hwnd);
        }
        catch (COMException ex) when (ComRetry.IsBusy(ex.HResult))
        {
            ComCleanup.SkipOnExit = true; // busy Excel의 프록시 Release도 블록할 수 있다
            return new AttachResult(null, $"Application.Hwnd busy: {ex.Message}", Busy: true);
        }
        if (appHwnd != meta.Hwnd)
            return new AttachResult(null,
                $"획득한 Application.Hwnd(0x{appHwnd:X})가 메타파일 HWND(0x{meta.Hwnd:X})와 불일치");
        return att;
    }

    /// <summary>
    /// attach(external) 세션의 재attach. 사용자는 여러 워크북 창을 오가므로 Application.Hwnd(활성
    /// 창의 XLMAIN)가 바뀔 수 있어 메타 HWND 일치를 요구하지 않는다 — 대신 PID 소유 창에서 얻은
    /// Application의 소유 PID와, 정본 워크북이 아직 열려 있는지로 검증한다.
    /// </summary>
    private static AttachResult TryAttachExternal(SessionMeta meta)
    {
        var windows = NativeMethods.TopLevelWindowsOf(meta.Pid)
            .Where(w => w.ClassName == "XLMAIN")
            .OrderByDescending(w => w.Hwnd.ToInt64() == meta.Hwnd) // 기록된 창 우선
            .ToList();
        string last = "사용자 Excel에서 워크북 창(EXCEL7)을 찾지 못함";
        foreach (var w in windows)
        {
            var excel7 = NativeMethods.FindExcel7Child(w.Hwnd);
            if (excel7 == IntPtr.Zero) continue;
            var att = ApplicationFromExcel7(excel7, meta.Pid, BusyRetries);
            if (att.App == null)
            {
                if (att.Busy) return att; // 같은 프로세스 — 다른 창도 똑같이 거절한다
                last = att.FailReason ?? last;
                continue;
            }
            if (!SessionService.IsWorkbookAlive(att.App, meta.WorkbookFullName))
                return new AttachResult(null, $"사용자 Excel(PID {meta.Pid})에서 정본 워크북이 닫힘");
            return att;
        }
        return new AttachResult(null, last);
    }

    /// <summary>session attach 탐색 결과. Apps는 PID별 Application(대상 선택 후 사용).</summary>
    public sealed record OpenWorkbookScan(
        List<AttachCandidate> Candidates, Dictionary<int, dynamic> Apps, List<int> BusyPids, List<string> Errors);

    /// <summary>
    /// 현재 데스크톱의 모든 Excel 인스턴스에서 열린 워크북을 수집한다(session attach, 설계 §11 L5).
    /// ROT/GetActiveObject는 쓰지 않는다(§6 금지 — 임의 인스턴스 반환) — XLMAIN 창을 PID별로 모아
    /// 프로세스마다 EXCEL7 하나로 NATIVEOM attach해 그 인스턴스의 Workbooks를 읽는다.
    /// </summary>
    public static OpenWorkbookScan ScanOpenWorkbooks()
    {
        var byPid = new Dictionary<int, List<IntPtr>>();
        NativeMethods.EnumWindows((h, _) =>
        {
            if (NativeMethods.GetWindowClassName(h) == "XLMAIN")
            {
                NativeMethods.GetWindowThreadProcessId(h, out uint pid);
                if (!byPid.TryGetValue((int)pid, out var list)) byPid[(int)pid] = list = new List<IntPtr>();
                list.Add(h);
            }
            return true;
        }, IntPtr.Zero);

        var scan = new OpenWorkbookScan(new(), new(), new(), new());
        foreach (var (pid, xlMains) in byPid)
        {
            bool done = false;
            foreach (var xl in xlMains)
            {
                var excel7 = NativeMethods.FindExcel7Child(xl);
                if (excel7 == IntPtr.Zero) continue;
                var att = ApplicationFromExcel7(excel7, pid, BusyRetries);
                if (att.App == null)
                {
                    if (att.Busy) { scan.BusyPids.Add(pid); done = true; break; }
                    scan.Errors.Add($"PID {pid}: {att.FailReason}");
                    continue;
                }
                try
                {
                    dynamic app = att.App;
                    int count = Convert.ToInt32(app.Workbooks.Count);
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic wb = app.Workbooks.Item(i);
                        scan.Candidates.Add(new AttachCandidate(pid,
                            Convert.ToString(wb.FullName) ?? "", Convert.ToBoolean(wb.ReadOnly)));
                    }
                    scan.Apps[pid] = app;
                }
                catch (COMException ex) when (ComRetry.IsBusy(ex.HResult))
                {
                    scan.BusyPids.Add(pid);
                }
                catch (Exception ex)
                {
                    scan.Errors.Add($"PID {pid}: {ex.Message}");
                }
                done = true;
                break;
            }
            if (!done && xlMains.Count > 0)
                scan.Errors.Add($"PID {pid}: 워크북 창(EXCEL7) 없음");
        }
        return scan;
    }

    /// <summary>
    /// EXCEL7 창 → NATIVEOM → Application. 획득한 Application.Hwnd의 소유 PID가 expectedPid와
    /// 같을 때만 성공. busy 거절은 retries번 재시도 후 Busy=true로 보고한다.
    /// </summary>
    public static AttachResult ApplicationFromExcel7(IntPtr excel7, int expectedPid, int retries)
    {
        string lastError = "";
        bool busy = false;
        for (int attempt = 0; attempt <= retries; attempt++)
        {
            if (attempt > 0) Thread.Sleep(BusyDelayMs);
            var iid = NativeMethods.IID_IDispatch;
            int hr = NativeMethods.AccessibleObjectFromWindow(
                excel7, NativeMethods.OBJID_NATIVEOM, ref iid, out object? nativeOm);
            if (hr != 0 || nativeOm == null)
            {
                lastError = $"AccessibleObjectFromWindow 실패 (hr=0x{hr:X8})";
                busy = false;
                continue;
            }
            try
            {
                dynamic window = nativeOm;
                dynamic app = window.Application;
                long appHwnd = Convert.ToInt64(app.Hwnd);
                NativeMethods.GetWindowThreadProcessId(new IntPtr(appHwnd), out uint appPid);
                if (appPid != (uint)expectedPid)
                    return new AttachResult(null, $"Application 소유 PID({appPid})가 기대 PID({expectedPid})와 불일치");
                return new AttachResult(app, null);
            }
            catch (COMException ex) when (ComRetry.IsBusy(ex.HResult))
            {
                // Excel이 방금 모달 처리/VBE 종료로 일시 busy이거나 사람이 셀을 편집 중
                lastError = $"NATIVEOM → Application busy: {ex.Message}";
                busy = true;
            }
            catch (Exception ex)
            {
                return new AttachResult(null, $"NATIVEOM → Application 획득 실패: {ex.Message}");
            }
        }
        return new AttachResult(null, lastError, busy);
    }
}
