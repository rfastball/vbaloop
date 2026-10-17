using System.Runtime.InteropServices;

namespace VbaLoop.Core;

/// <summary>
/// 프로세스 (pid, ppid, exe명) 스냅샷 — CreateToolhelp32Snapshot 1회로 전 프로세스를 얻는다
/// (hop당 OpenProcess/WMI가 불필요, 문서화된 API — 검토 확정).
/// </summary>
public static class ProcessTree
{
    public sealed record ProcEntry(int Pid, int ParentPid, string Name);

    private const uint TH32CS_SNAPPROCESS = 0x00000002;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct PROCESSENTRY32W
    {
        public uint dwSize;
        public uint cntUsage;
        public uint th32ProcessID;
        public IntPtr th32DefaultHeapID;
        public uint th32ModuleID;
        public uint cntThreads;
        public uint th32ParentProcessID;
        public int pcPriClassBase;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szExeFile;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessID);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern bool Process32FirstW(IntPtr hSnapshot, ref PROCESSENTRY32W lppe);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern bool Process32NextW(IntPtr hSnapshot, ref PROCESSENTRY32W lppe);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    private static readonly IntPtr InvalidHandle = new(-1);

    /// <summary>이름은 확장자 제거 + 소문자 (denylist 매칭용).</summary>
    public static Dictionary<int, ProcEntry> Snapshot()
    {
        var map = new Dictionary<int, ProcEntry>();
        IntPtr snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snap == InvalidHandle) return map;
        try
        {
            var entry = new PROCESSENTRY32W { dwSize = (uint)Marshal.SizeOf<PROCESSENTRY32W>() };
            if (!Process32FirstW(snap, ref entry)) return map;
            do
            {
                string name = Path.GetFileNameWithoutExtension(entry.szExeFile).ToLowerInvariant();
                map[(int)entry.th32ProcessID] = new ProcEntry(
                    (int)entry.th32ProcessID, (int)entry.th32ParentProcessID, name);
            } while (Process32NextW(snap, ref entry));
            return map;
        }
        finally
        {
            CloseHandle(snap);
        }
    }
}
