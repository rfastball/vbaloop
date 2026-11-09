using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace VbaLoop.Core;

/// <summary>파일 사용자를 조회만 한다. 삭제/이름 변경/Restart Manager 종료 API는 호출하지 않는다.</summary>
public static class FileUseDiagnostics
{
    public sealed record User(int Pid, long StartFileTime, string? ProcessName, string Application, string Service, bool ToolExcel);
    public sealed record Snapshot(int WriteAccessError, int DeleteAccessError, int QueryError, User[] Users);

    public static Snapshot Read(string path, int? excelPid = null)
    {
        // DELETE 접근권 요청은 삭제 작업이 아니다. OPEN_EXISTING, FILE_SHARE_ALL, 플래그 0.
        int write = Probe(path, 0xC0000000), delete = Probe(path, 0x00010000);
        uint session;
        int error = RmStartSession(out session, 0, new StringBuilder(33));
        if (error != 0) return new(write, delete, error, []);
        try
        {
            error = RmRegisterResources(session, 1, [Path.GetFullPath(path)], 0, IntPtr.Zero, 0, IntPtr.Zero);
            if (error != 0) return new(write, delete, error, []);
            uint count = 0;
            error = RmGetList(session, out uint needed, ref count, null, out _);
            for (int attempt = 0; error == 234 && attempt < 3; attempt++)
            {
                if (needed > 1024) return new(write, delete, 234, []);
                var users = new ProcessInfo[needed];
                count = needed;
                error = RmGetList(session, out needed, ref count, users, out _);
                if (error == 0)
                    return new(write, delete, 0, users.Take((int)count).Select(u => new User(
                        u.Process.Pid, ((long)u.Process.StartHigh << 32) | u.Process.StartLow,
                        ProcessName(u.Process), u.Application, u.Service, u.Process.Pid == excelPid)).ToArray());
            }
            return new(write, delete, error, []);
        }
        finally { RmEndSession(session); }
    }

    public static void Capture(UpdaterDiagnostics? log, string path, string reason, int? excelPid = null)
    {
        if (log == null) return;
        log.Write("file_use_query_started", new { reason });
        try { log.Write("file_use", new { reason, snapshot = Read(path, excelPid) }); }
        catch (Exception ex) { log.Error("file_use_unavailable", ex); }
    }

    private static int Probe(string path, uint access)
    {
        using var handle = CreateFileW(path, access, 7, IntPtr.Zero, 3, 0, IntPtr.Zero);
        return handle.IsInvalid ? Marshal.GetLastWin32Error() : 0;
    }

    private static string? ProcessName(UniqueProcess identity)
    {
        try
        {
            using var process = Process.GetProcessById(identity.Pid);
            long started = ((long)identity.StartHigh << 32) | identity.StartLow;
            return process.StartTime.ToUniversalTime().ToFileTimeUtc() == started ? process.ProcessName : null;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or Win32Exception) { return null; }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct UniqueProcess { public int Pid; public uint StartLow, StartHigh; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ProcessInfo
    {
        public UniqueProcess Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string Application;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string Service;
        public uint ApplicationType, Status, SessionId;
        [MarshalAs(UnmanagedType.Bool)] public bool Restartable;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
    private static extern SafeFileHandle CreateFileW(string path, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmStartSession(out uint session, uint flags, StringBuilder key);
    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmRegisterResources(uint session, uint files, string[] names, uint applications, IntPtr apps, uint services, IntPtr serviceNames);
    [DllImport("rstrtmgr.dll")]
    private static extern int RmGetList(uint session, out uint needed, ref uint count, [In, Out] ProcessInfo[]? users, out uint rebootReasons);
    [DllImport("rstrtmgr.dll")]
    private static extern int RmEndSession(uint session);
}
