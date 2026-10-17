using System.Runtime.InteropServices;

namespace VbaLoop.Core;

/// <summary>
/// 다른 프로세스의 명령줄 판독 — NtQueryInformationProcess(ProcessCommandLineInformation, Win 8.1+).
/// PROCESS_QUERY_LIMITED_INFORMATION만 요구하므로 같은 사용자의 프로세스는 읽히고, 다른
/// 사용자/상승된 프로세스는 null(판독 불가)이다 — 호출자는 null을 "정리 대상 아님"으로 다룬다.
/// WMI(System.Management) 의존을 피하기 위한 직접 호출.
/// </summary>
public static class ProcessCommandLine
{
    private const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
    private const int ProcessCommandLineInformation = 60;

    [StructLayout(LayoutKind.Sequential)]
    private struct UNICODE_STRING
    {
        public ushort Length;
        public ushort MaximumLength;
        public IntPtr Buffer;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, int dwProcessId);

    [DllImport("kernel32.dll")]
    private static extern bool CloseHandle(IntPtr hObject);

    [DllImport("ntdll.dll")]
    private static extern int NtQueryInformationProcess(
        IntPtr processHandle, int processInformationClass, IntPtr processInformation,
        int processInformationLength, out int returnLength);

    public static string? Get(int pid)
    {
        IntPtr h = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid);
        if (h == IntPtr.Zero) return null;
        try
        {
            // 첫 호출은 필요한 길이만 받는다(STATUS_INFO_LENGTH_MISMATCH)
            NtQueryInformationProcess(h, ProcessCommandLineInformation, IntPtr.Zero, 0, out int len);
            if (len <= 0 || len > 1 << 20) len = 64 * 1024;
            IntPtr buf = Marshal.AllocHGlobal(len);
            try
            {
                int status = NtQueryInformationProcess(h, ProcessCommandLineInformation, buf, len, out _);
                if (status != 0) return null;
                var us = Marshal.PtrToStructure<UNICODE_STRING>(buf);
                return us.Buffer == IntPtr.Zero ? "" : Marshal.PtrToStringUni(us.Buffer, us.Length / 2);
            }
            finally
            {
                Marshal.FreeHGlobal(buf);
            }
        }
        catch
        {
            return null;
        }
        finally
        {
            CloseHandle(h);
        }
    }
}
