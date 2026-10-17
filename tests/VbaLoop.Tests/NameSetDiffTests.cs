using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class NameSetDiffTests
{
    [Fact]
    public void DuplicateInjection_Caught()
    {
        // 검증 계획 §7-4(a): 의도적 중복 주입 시나리오(<name>1 패턴)를 잡아낸다
        var r = NameSetDiff.Check(
            before: new[] { "modCache", "modLogger" },
            expectedAdded: Array.Empty<string>(),
            expectedRemoved: Array.Empty<string>(),
            after: new[] { "modCache", "modCache1", "modLogger" });
        Assert.Equal(new[] { "modCache1" }, r.Unexpected);
        Assert.Equal(new[] { "modCache1" }, r.DuplicateSignatures);
        Assert.Empty(r.Missing);
    }

    [Fact]
    public void LegitimateAdd_NoFalsePositive()
    {
        // 검증 계획 §7-4(b): 정상 신규 모듈 Add를 오탐하지 않음
        var r = NameSetDiff.Check(
            before: new[] { "modCache" },
            expectedAdded: new[] { "modNew" },
            expectedRemoved: Array.Empty<string>(),
            after: new[] { "modCache", "modNew" });
        Assert.Empty(r.Unexpected);
        Assert.Empty(r.Missing);
        Assert.Empty(r.DuplicateSignatures);
    }

    [Fact]
    public void RemovedAndReAdded_FormRoundtrip_Expected()
    {
        // 폼 Export/Import 왕복: Remove 후 같은 이름 Import → 기대 집합에 유지
        var r = NameSetDiff.Check(
            before: new[] { "frmFilter01" },
            expectedAdded: new[] { "frmFilter01" },
            expectedRemoved: new[] { "frmFilter01" },
            after: new[] { "frmFilter01" });
        Assert.Empty(r.Unexpected);
        Assert.Empty(r.Missing);
    }

    [Fact]
    public void MissingModule_Caught()
    {
        var r = NameSetDiff.Check(
            before: new[] { "modA", "modB" },
            expectedAdded: Array.Empty<string>(),
            expectedRemoved: Array.Empty<string>(),
            after: new[] { "modA" });
        Assert.Equal(new[] { "modB" }, r.Missing);
    }

    [Fact]
    public void CaseInsensitive_NoFalsePositive()
    {
        var r = NameSetDiff.Check(
            before: new[] { "ModCache" },
            expectedAdded: Array.Empty<string>(),
            expectedRemoved: Array.Empty<string>(),
            after: new[] { "modcache" });
        Assert.Empty(r.Unexpected);
        Assert.Empty(r.Missing);
    }
}
