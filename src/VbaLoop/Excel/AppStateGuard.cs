using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// 실행 전 Excel 상태 스냅숏과 중간에 끊긴 매크로의 뒷정리(설계 §14.8). 러너가 매크로·테스트마다 호출 직전에
/// 찍고, 취소·런타임 오류로 끝났을 때만 되돌린다 — 정상 완료한 매크로가 바꾼 설정은 의도이므로 건드리지 않는다.
/// 모든 COM 호출은 예외를 삼킨다: 뒷정리는 판정을 바꾸지 않고, 세션 사망·poison은 러너의 기존 판정이 맡는다.
/// </summary>
public sealed class AppStateGuard
{
    private readonly AppSettings _settings;
    private readonly List<string> _workbooks;

    private AppStateGuard(AppSettings settings, List<string> workbooks)
    {
        _settings = settings;
        _workbooks = workbooks;
    }

    /// <summary>스냅숏. 읽지 못하면 null(그 실행은 뒷정리 없이 진행).</summary>
    public static AppStateGuard? Take(dynamic app)
    {
        try
        {
            AppSettings settings = ReadSettings(app);
            List<string> workbooks = WorkbookNames(app);
            return new AppStateGuard(settings, workbooks);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// 되돌린다. closeNewWorkbooks(일회용 세션)면 매크로가 연 채 남긴 통합 문서를 저장 없이 닫는다 — 남겨 두면 세션 종료가
    /// 모르는 문서로 보고 Excel을 사용자에게 넘긴다. 닫는 것은 경로가 없는 새 문서와 저장할 변경이 없는 문서뿐이고, 세션
    /// 워크북(SaveAs로 이름이 바뀌었어도)은 닫지 않는다. 이벤트는 끈 채로 되돌리고 마지막에 원래 값으로 둔다 — 계산 모드
    /// 복원의 재계산이나 문서 닫기가 이벤트 코드(MsgBox 등)를 불러 이 호출을 막지 않게.
    /// </summary>
    public RunAftermath Restore(dynamic app, bool closeNewWorkbooks, object? sessionWorkbook)
    {
        var restored = new List<string>();
        var closed = new List<string>();
        var leftOpen = new List<string>();
        var failed = new List<string>();

        // 설정은 이벤트를 끄기 전에 읽는다 - 매크로가 바꾼 것만 보고한다
        AppSettings? now = null;
        try { now = ReadSettings(app); }
        catch { failed.Add("Application settings"); }
        IReadOnlyList<string> changed = now != null ? AppSettings.Changed(_settings, now) : Array.Empty<string>();

        IReadOnlyList<string> opened;
        try
        {
            List<string> after = WorkbookNames(app);
            opened = RunAftermath.NewWorkbooks(_workbooks, after);
        }
        catch
        {
            opened = Array.Empty<string>();
            failed.Add("workbook list");
        }

        bool quiet = opened.Count > 0 || changed.Any(n => n != "EnableEvents");
        if (quiet) Try(() => { app.EnableEvents = false; });
        foreach (string name in opened)
        {
            dynamic? wb = FindOpen(app, name);
            if (wb == null) continue; // 그새 닫혔다
            if (sessionWorkbook != null && SameComObject(wb, sessionWorkbook)) continue; // 세션 워크북의 SaveAs
            bool disposable;
            try { disposable = IsDisposable(wb); }
            catch { disposable = false; } // 읽지 못하면 닫지 않는다
            if (!closeNewWorkbooks || !disposable)
            {
                leftOpen.Add(name);
                continue;
            }
            if (Try(() => { ComRetry.Invoke(() => { wb.Close(false); }); })) closed.Add(name);
            else
            {
                leftOpen.Add(name);
                failed.Add($"close {name}");
            }
        }
        foreach (string name in changed.Where(n => n != "EnableEvents"))
        {
            if (Try(() => Apply(app, name, _settings))) restored.Add(name);
            else failed.Add($"Application.{name}");
        }
        // 이벤트는 마지막에: 매크로가 바꿨으면 복원으로 보고하고, 뒷정리가 끈 것이면 조용히 원래 값으로 돌린다
        if (changed.Contains("EnableEvents"))
        {
            if (Try(() => Apply(app, "EnableEvents", _settings))) restored.Add("EnableEvents");
            else failed.Add("Application.EnableEvents");
        }
        else if (quiet)
        {
            Try(() => { app.EnableEvents = _settings.EnableEvents; });
        }

        return new RunAftermath(restored, closed, leftOpen, failed);
    }

    private static AppSettings ReadSettings(dynamic app)
    {
        object statusBar = ComRetry.Invoke<object>(() => app.StatusBar);
        return new AppSettings(
            EnableEvents: Convert.ToBoolean(ComRetry.Invoke<object>(() => app.EnableEvents)),
            Calculation: Convert.ToInt32(ComRetry.Invoke<object>(() => app.Calculation)),
            StatusBar: AppSettings.StatusBarText(statusBar),
            Interactive: Convert.ToBoolean(ComRetry.Invoke<object>(() => app.Interactive)),
            Cursor: Convert.ToInt32(ComRetry.Invoke<object>(() => app.Cursor)));
    }

    private static void Apply(dynamic app, string name, AppSettings s)
    {
        switch (name)
        {
            case "EnableEvents": ComRetry.Invoke(() => { app.EnableEvents = s.EnableEvents; }); break;
            case "Calculation": ComRetry.Invoke(() => { app.Calculation = s.Calculation; }); break;
            case "StatusBar":
                if (s.StatusBar is string text) ComRetry.Invoke(() => { app.StatusBar = text; });
                else ComRetry.Invoke(() => { app.StatusBar = false; });
                break;
            case "Interactive": ComRetry.Invoke(() => { app.Interactive = s.Interactive; }); break;
            case "Cursor": ComRetry.Invoke(() => { app.Cursor = s.Cursor; }); break;
        }
    }

    private static List<string> WorkbookNames(dynamic app)
    {
        var names = new List<string>();
        int count = Convert.ToInt32(ComRetry.Invoke<object>(() => app.Workbooks.Count));
        for (int i = 1; i <= count; i++)
            names.Add(Convert.ToString(ComRetry.Invoke<object>(() => app.Workbooks.Item(i).FullName)) ?? "");
        return names;
    }

    private static dynamic? FindOpen(dynamic app, string fullName)
    {
        try
        {
            int count = Convert.ToInt32(app.Workbooks.Count);
            for (int i = 1; i <= count; i++)
            {
                dynamic wb = app.Workbooks.Item(i);
                if (string.Equals(Convert.ToString(wb.FullName), fullName, StringComparison.OrdinalIgnoreCase))
                    return wb;
            }
        }
        catch { /* 목록을 읽지 못했다 - 닫지 않는다 */ }
        return null;
    }

    /// <summary>저장 없이 닫아도 잃을 것이 없는 문서: 경로가 없는 새 문서, 또는 저장할 변경이 없는 문서. 경로가 있고 변경이
    /// 남은 문서(매크로가 열어 고친 파일, 실행 중 사용자가 연 파일)는 닫지 않고 보고한다.</summary>
    private static bool IsDisposable(dynamic wb)
        => string.IsNullOrEmpty(Convert.ToString(wb.Path)) || Convert.ToBoolean(wb.Saved);

    private static bool SameComObject(object a, object b)
    {
        if (ReferenceEquals(a, b)) return true;
        IntPtr ua = IntPtr.Zero, ub = IntPtr.Zero;
        try
        {
            ua = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(a);
            ub = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(b);
            return ua == ub;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (ua != IntPtr.Zero) System.Runtime.InteropServices.Marshal.Release(ua);
            if (ub != IntPtr.Zero) System.Runtime.InteropServices.Marshal.Release(ub);
        }
    }

    private static bool Try(Action action)
    {
        try
        {
            action();
            return true;
        }
        catch
        {
            return false;
        }
    }
}
