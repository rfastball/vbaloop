using System.Runtime.InteropServices;

namespace VbaLoop.Excel;

/// <summary>
/// 워크북 파일 선택 대화상자. comdlg32 <c>GetOpenFileNameW</c>를 직접 부른다 — FolderPicker와
/// 같은 이유로 WinForms(OpenFileDialog) 프레임워크 참조를 붙이지 않는다. 호출 스레드가 STA여야
/// 한다(Program.Main이 [STAThread]).
/// </summary>
public static class FilePicker
{
    private const int OFN_HIDEREADONLY = 0x00000004;
    private const int OFN_NOCHANGEDIR = 0x00000008;
    private const int OFN_PATHMUSTEXIST = 0x00000800;
    private const int OFN_FILEMUSTEXIST = 0x00001000;
    private const int OFN_EXPLORER = 0x00080000;
    private const int MaxPath = 32768;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct OPENFILENAMEW
    {
        public int lStructSize;
        public IntPtr hwndOwner;
        public IntPtr hInstance;
        public string lpstrFilter;
        public IntPtr lpstrCustomFilter;
        public int nMaxCustFilter;
        public int nFilterIndex;
        public IntPtr lpstrFile;
        public int nMaxFile;
        public IntPtr lpstrFileTitle;
        public int nMaxFileTitle;
        public string? lpstrInitialDir;
        public string lpstrTitle;
        public int Flags;
        public short nFileOffset;
        public short nFileExtension;
        public string? lpstrDefExt;
        public IntPtr lCustData;
        public IntPtr lpfnHook;
        public string? lpTemplateName;
        public IntPtr pvReserved;
        public int dwReserved;
        public int FlagsEx;
    }

    [DllImport("comdlg32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool GetOpenFileNameW(ref OPENFILENAMEW ofn);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetConsoleWindow();

    /// <summary>선택된 .xlsm의 전체 경로. 취소하거나 대화상자를 띄울 수 없으면 null.</summary>
    public static string? PickWorkbook(string title)
    {
        IntPtr buffer = Marshal.AllocHGlobal(MaxPath * sizeof(char));
        try
        {
            Marshal.WriteInt16(buffer, 0);
            var ofn = new OPENFILENAMEW
            {
                lStructSize = Marshal.SizeOf<OPENFILENAMEW>(),
                // 콘솔 창을 소유자로 — 대화상자가 마법사 창 뒤에 숨지 않게 한다
                hwndOwner = GetConsoleWindow(),
                // 필터는 "표시명\0패턴\0...\0\0" — 문자열 마샬링은 내장 NUL을 보존한다
                lpstrFilter = "Excel Macro-Enabled Workbook (*.xlsm)\0*.xlsm\0",
                nFilterIndex = 1,
                lpstrFile = buffer,
                nMaxFile = MaxPath,
                lpstrTitle = title,
                Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR,
            };
            if (!GetOpenFileNameW(ref ofn)) return null; // 취소 또는 오류(CommDlgExtendedError) — 호출부가 폴백
            string? path = Marshal.PtrToStringUni(buffer);
            return string.IsNullOrWhiteSpace(path) ? null : path;
        }
        catch
        {
            return null; // 대화상자 불가(세션 0 등) — 호출부가 콘솔 입력으로 폴백
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }
}
