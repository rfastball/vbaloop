using System.Runtime.InteropServices;
using System.Text;

namespace VbaLoop.Excel;

internal static class NativeMethods
{
    public const uint OBJID_NATIVEOM = 0xFFFFFFF0;
    public static readonly Guid IID_IDispatch = new("00020400-0000-0000-C000-000000000046");

    // 모달 워치독용 메시지/플래그 (Watch-Modal.ps1 이식)
    public const uint WM_GETTEXT = 0x000D;
    public const uint WM_CLOSE = 0x0010;
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_COMMAND = 0x0111;
    public const uint SMTO_NORMAL = 0x0000;
    public const uint SMTO_ABORTIFHUNG = 0x0002;

    // 언어 중립 버튼 판별용: 스타일 비트, 소유자 체인
    public const int GWL_STYLE = -16;
    public const uint GW_OWNER = 4;
    public const int WS_VISIBLE = 0x10000000;
    public const int BS_TYPEMASK = 0x0F;
    public const int BS_PUSHBUTTON = 0x00;
    public const int BS_DEFPUSHBUTTON = 0x01;

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    /// <summary>타 프로세스 창에도 행 없이 안전 — 캐시된 타이틀을 읽는다(WM_GETTEXT 미전송).</summary>
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SendMessageTimeoutW(
        IntPtr hWnd, uint msg, IntPtr wParam, StringBuilder? lParam,
        uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageTimeoutW(
        IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam,
        uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    [DllImport("user32.dll")]
    public static extern bool PostMessageW(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);

    /// <summary>다이얼로그 템플릿의 컨트롤 ID — 리소스 지역화와 무관하게 언어 공통.</summary>
    [DllImport("user32.dll")]
    public static extern int GetDlgCtrlID(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowEnabled(IntPtr hWnd);

    /// <summary>조상까지 보는 가시성 — 숨은 다이얼로그 판별용.</summary>
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(
        IntPtr hwnd, uint dwObjectID, ref Guid riid,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppvObject);

    [DllImport("kernel32.dll")]
    public static extern int GetACP();

    /// <summary>이 프로세스의 콘솔에 붙어 있는 프로세스 수. 1이면 콘솔을 단독 소유 —
    /// 탐색기에서 더블클릭해 새 콘솔이 할당된 경우다(셸에서 부르면 셸까지 2 이상).
    /// 0이면 콘솔 없음. setup 마법사 진입 판별에 쓴다.</summary>
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint GetConsoleProcessList(uint[] lpdwProcessList, uint dwProcessCount);

    [DllImport("ole32.dll")]
    public static extern int CoRegisterMessageFilter(
        IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    public static string GetWindowClassName(IntPtr hWnd)
    {
        var sb = new StringBuilder(256);
        _ = GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    /// <summary>
    /// pid가 소유한 최상위 창 목록(보이지 않는 창 포함). 창 핸들/클래스/가시성 조회만 하고
    /// 대상 프로세스에 메시지를 보내지 않으므로 행 걸린 Excel에도 안전하다.
    /// </summary>
    public static List<(IntPtr Hwnd, string ClassName, bool Visible)> TopLevelWindowsOf(int pid)
    {
        var list = new List<(IntPtr, string, bool)>();
        EnumWindows((h, _) =>
        {
            GetWindowThreadProcessId(h, out uint owner);
            if (owner == (uint)pid)
                list.Add((h, GetWindowClassName(h), IsWindowVisible(h)));
            return true;
        }, IntPtr.Zero);
        return list;
    }

    /// <summary>XLMAIN 아래에서 EXCEL7(워크북) 자식 창을 찾는다 (설계 §4 재attach 절차).</summary>
    public static IntPtr FindExcel7Child(IntPtr xlMain)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(xlMain, (h, _) =>
        {
            if (GetWindowClassName(h) == "EXCEL7")
            {
                found = h;
                return false; // 중단
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
