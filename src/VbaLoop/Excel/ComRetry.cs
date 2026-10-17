using System.Runtime.InteropServices;

namespace VbaLoop.Excel;

/// <summary>
/// COM busy 명시 재시도 + poison 판정 (설계 §4 push 2단계, §6).
/// 재시도는 busy HRESULT 목록에 한정한다 — 그 외 예외를 재시도하면
/// 모달 시나리오에서 행이 5배가 되므로 busy만 재시도한다.
/// </summary>
public static class ComRetry
{
    /// <summary>RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER / VBA_E_IGNORE</summary>
    public static readonly int[] BusyHResults =
    {
        unchecked((int)0x80010001), // RPC_E_CALL_REJECTED (-2147418111)
        unchecked((int)0x8001010A), // RPC_E_SERVERCALL_RETRYLATER (-2147417846)
        unchecked((int)0x800AC472), // VBA_E_IGNORE (-2146777998)
    };

    /// <summary>이 HRESULT가 나오면 세션은 오염(poisoned) — 재사용 금지, session start 강제.</summary>
    public static readonly int[] PoisonHResults =
    {
        unchecked((int)0x80010108), // RPC_E_DISCONNECTED
        unchecked((int)0x80010007), // RPC_E_SERVER_DIED
        unchecked((int)0x80010012), // RPC_E_SERVER_DIED_DNE
        unchecked((int)0x800706BA), // RPC 서버를 사용할 수 없음
        unchecked((int)0x800706BE), // 원격 프로시저 호출 실패
    };

    public static bool IsBusy(int hr) => BusyHResults.Contains(hr);
    public static bool IsPoison(int hr) => PoisonHResults.Contains(hr);

    public static T Invoke<T>(Func<T> action, int retries = 5, int delayMs = 400)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                return action();
            }
            catch (COMException ex) when (IsBusy(ex.HResult) && attempt < retries - 1)
            {
                Thread.Sleep(delayMs);
            }
        }
    }

    public static void Invoke(Action action, int retries = 5, int delayMs = 400)
        => Invoke<object?>(() => { action(); return null; }, retries, delayMs);
}
