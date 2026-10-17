using System.Runtime.InteropServices;

namespace VbaLoop.Excel;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

/// <summary>
/// COM 계층 RPC_E_CALL_REJECTED 자동 재시도 (xlFlow ExcelBridgeSupport.OleMessageFilter 이식).
/// SERVERCALL_RETRYLATER를 10초까지 150ms 간격으로 재시도. STA 스레드에서 등록해야 한다.
/// </summary>
public sealed class OleMessageFilter : IOleMessageFilter
{
    private const int ServerCallIsHandled = 0;
    private const int ServerCallRetryLater = 2;
    private const int PendingMsgWaitDefProcess = 2;
    private const int RetryTimeoutMs = 10_000;

    private IOleMessageFilter? _old;

    public static OleMessageFilter Register()
    {
        var filter = new OleMessageFilter();
        int hr = NativeMethods.CoRegisterMessageFilter(filter, out filter._old);
        if (hr != 0)
            throw new InvalidOperationException(
                $"CoRegisterMessageFilter 실패 (hr=0x{hr:X8}) — STA 스레드가 아닐 수 있음");
        return filter;
    }

    public void Revoke()
    {
        NativeMethods.CoRegisterMessageFilter(_old, out _);
        _old = null;
    }

    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        => ServerCallIsHandled;

    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        if (dwRejectType == ServerCallRetryLater && dwTickCount < RetryTimeoutMs)
            return 150; // 150ms 후 재시도
        return -1; // 취소 → 호출자에 RPC_E_CALL_REJECTED 전파 (ComRetry가 2차 방어)
    }

    int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        => PendingMsgWaitDefProcess;
}
