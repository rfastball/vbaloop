using System.Runtime.InteropServices;

namespace VbaLoop.Core;

/// <summary>
/// 부모를 기다리게 하지 않는 분리된 자식 프로세스(업데이트 확인, 설계 §12.4).
/// Process.Start(UseShellExecute=false)는 항상 bInheritHandles=TRUE로 CreateProcess를 부른다 — 그러면
/// 부모의 상속 가능한 stdout/stderr 파이프(에이전트가 출력을 읽는 파이프)가 자식에게 복제되어, 부모가
/// 끝나도 호출자의 읽기가 자식 종료까지 막힌다. 그래서 CreateProcessW를 직접 불러 핸들을 상속하지 않고,
/// 콘솔 창 없이(CREATE_NO_WINDOW) 띄운다. 가능하면 호출자의 job에서도 떼어 낸다(job 종료에 같이 죽지 않게).
/// </summary>
public static class DetachedProcess
{
    private const uint CREATE_NO_WINDOW = 0x08000000;
    private const uint CREATE_NEW_PROCESS_GROUP = 0x00000200;
    private const uint CREATE_BREAKAWAY_FROM_JOB = 0x01000000;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct STARTUPINFO
    {
        public int cb;
        public string? lpReserved, lpDesktop, lpTitle;
        public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
        public short wShowWindow, cbReserved2;
        public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROCESS_INFORMATION
    {
        public IntPtr hProcess, hThread;
        public int dwProcessId, dwThreadId;
    }

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CreateProcessW(string? lpApplicationName, System.Text.StringBuilder lpCommandLine,
        IntPtr lpProcessAttributes, IntPtr lpThreadAttributes, bool bInheritHandles, uint dwCreationFlags,
        IntPtr lpEnvironment, string? lpCurrentDirectory, ref STARTUPINFO lpStartupInfo,
        out PROCESS_INFORMATION lpProcessInformation);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr h);

    /// <summary>exe를 인자 하나와 함께 분리 실행(인자는 공백·따옴표 없는 내부 verb). 실패해도 예외 없이 false.</summary>
    public static bool TryStart(string exePath, string argument)
    {
        try
        {
            var cmd = new System.Text.StringBuilder($"\"{exePath}\" {argument}");
            var dir = Path.GetDirectoryName(exePath);
            foreach (var flags in new[]
                     {
                         CREATE_NO_WINDOW | CREATE_NEW_PROCESS_GROUP | CREATE_BREAKAWAY_FROM_JOB,
                         CREATE_NO_WINDOW | CREATE_NEW_PROCESS_GROUP, // job이 breakaway를 허용하지 않으면
                     })
            {
                var si = new STARTUPINFO { cb = Marshal.SizeOf<STARTUPINFO>() };
                if (!CreateProcessW(null, cmd, IntPtr.Zero, IntPtr.Zero, false, flags,
                        IntPtr.Zero, dir, ref si, out var pi))
                    continue;
                CloseHandle(pi.hThread);
                CloseHandle(pi.hProcess);
                return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }
}
