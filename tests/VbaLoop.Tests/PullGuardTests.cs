using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class PullGuardTests
{
    [Fact]
    public void NoHistory_Passes()
    {
        // 최초 pull(일회성 정규화): 이력이 전혀 없으면 통과
        Assert.Null(PullGuard.FindBlockReason("AAA", null, null));
    }

    [Fact]
    public void MatchesLastPush_Passes()
    {
        // src 편집이 전부 push된 상태 — 워크북 진실만 반영하므로 안전
        Assert.Null(PullGuard.FindBlockReason("AAA", "AAA", null));
        Assert.Null(PullGuard.FindBlockReason("AAA", "AAA", "BBB"));
    }

    [Fact]
    public void MatchesLastPull_Passes()
    {
        // 마지막 pull 이후 src 무편집 — 덮어써도 잃을 것이 없음
        Assert.Null(PullGuard.FindBlockReason("BBB", null, "BBB"));
        Assert.Null(PullGuard.FindBlockReason("BBB", "AAA", "BBB"));
    }

    [Fact]
    public void MatchesNeither_Blocks()
    {
        // 동기화된 적 없는 로컬 편집 — 차단 사유 반환
        Assert.NotNull(PullGuard.FindBlockReason("CCC", "AAA", "BBB"));
        Assert.NotNull(PullGuard.FindBlockReason("CCC", "AAA", null));
        Assert.NotNull(PullGuard.FindBlockReason("CCC", null, "BBB"));
    }
}
