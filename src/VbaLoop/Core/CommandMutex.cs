using System.Security.Cryptography;
using System.Text;

namespace VbaLoop.Core;

/// <summary>
/// 프로젝트당 vbaloop 명령 직렬화 — OS named mutex.
/// project lease는 "어느 에이전트가 이 프로젝트를 소유하나"를 다루지만, 같은 소유자
/// (동일 호스트 아래 서브에이전트들)의 동시 명령은 걸러내지 못한다 — 둘 다 renew로
/// 게이트를 통과해 push+test가 같은 Excel에 인터리브될 수 있음(검토 확정).
/// 커널 뮤텍스는 보유자 사망 시 자동 해제되므로 staleness
/// 개념 자체가 없다 — AbandonedMutexException은 획득 취급.
/// </summary>
public sealed class CommandMutex : IDisposable
{
    private readonly Mutex _mutex;

    private CommandMutex(Mutex mutex) => _mutex = mutex;

    public static CommandMutex Acquire(ProjectConfig cfg)
    {
        string name = @"Local\vbaloop-" + KeyFor(cfg.RootDir);
        var mutex = new Mutex(initiallyOwned: false, name);
        bool held;
        try
        {
            held = mutex.WaitOne(0);
        }
        catch (AbandonedMutexException)
        {
            held = true; // 전 보유자 프로세스 사망 — 커널이 해제한 것을 승계
        }
        if (!held)
        {
            mutex.Dispose();
            throw new CliError("Another vbaloop command is running in this project",
                "Retry after that command finishes (concurrent runs would interleave on the same Excel session)");
        }
        return new CommandMutex(mutex);
    }

    public void Dispose()
    {
        try { _mutex.ReleaseMutex(); } catch (ApplicationException) { /* 소유 아님 — 무해 */ }
        _mutex.Dispose();
    }

    private static string KeyFor(string rootDir)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(rootDir.ToLowerInvariant()));
        return Convert.ToHexString(bytes)[..16];
    }
}
