using System.Diagnostics;
using System.Runtime.CompilerServices;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ComCleanupTests
{
    private static readonly ManualResetEventSlim Release = new(false);

    private sealed class BlockingFinalizer
    {
        ~BlockingFinalizer() => Release.Wait(TimeSpan.FromSeconds(10));
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void Allocate() => _ = new BlockingFinalizer();

    [Fact]
    public void Run_CompletesQuickly_WhenNothingBlocks()
    {
        Assert.True(ComCleanup.Run(TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public void Run_GivesUp_WhenAFinalizerBlocks()
    {
        // 격리 Excel의 프록시 Release가 멈춘 상황을 흉내 낸다: 파이널라이저가 블록돼도 제한 시간에 반환
        Release.Reset();
        Allocate();
        var sw = Stopwatch.StartNew();
        bool done = ComCleanup.Run(TimeSpan.FromMilliseconds(300));
        sw.Stop();
        Release.Set(); // 파이널라이저 스레드를 풀어 다른 테스트에 영향이 없게 한다
        Assert.False(done);
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(3), $"elapsed {sw.Elapsed}");
    }
}
