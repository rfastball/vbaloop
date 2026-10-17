using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>
/// 격리 모델의 판정 상태 머신(설계 §11). kill은 "도구 기동 + 신원 일치 + 세션 워크북 외 워크북
/// 없음의 적극적 확인"일 때만, 그 밖의 불확실성은 전부 격리여야 한다.
/// </summary>
public class RecoveryPolicyTests
{
    private const string Wb = @"C:\Work\Book.xlsm";

    private static ProbeResult Responsive(params string[] wbs) => new(ProbeKind.Responsive, wbs);

    [Fact]
    public void Auto_ProcessGone_Clears_EvenForExternal()
    {
        Assert.Equal(AutoAction.Clear, RecoveryPolicy.DecideAuto(false, IdentityState.Gone, null, Wb).Action);
        Assert.Equal(AutoAction.Clear, RecoveryPolicy.DecideAuto(true, IdentityState.Gone, null, Wb).Action);
    }

    [Fact]
    public void Auto_External_NeverKills_EvenWhenProvablyEmpty()
    {
        var d = RecoveryPolicy.DecideAuto(true, IdentityState.Match, Responsive(), Wb);
        Assert.Equal(AutoAction.Quarantine, d.Action);
        d = RecoveryPolicy.DecideAuto(true, IdentityState.Match,
            new ProbeResult(ProbeKind.Windowless, Array.Empty<string>()), Wb);
        Assert.Equal(AutoAction.Quarantine, d.Action);
    }

    [Fact]
    public void Auto_UnknownIdentity_Quarantines()
    {
        Assert.Equal(AutoAction.Quarantine,
            RecoveryPolicy.DecideAuto(false, IdentityState.Unknown, Responsive(Wb), Wb).Action);
    }

    [Fact]
    public void Auto_NoProbe_Quarantines()
    {
        Assert.Equal(AutoAction.Quarantine, RecoveryPolicy.DecideAuto(false, IdentityState.Match, null, Wb).Action);
    }

    [Theory]
    [InlineData(ProbeKind.Busy)]
    [InlineData(ProbeKind.Unresponsive)]
    public void Auto_UncertainProbe_Quarantines(ProbeKind kind)
    {
        var d = RecoveryPolicy.DecideAuto(false, IdentityState.Match,
            new ProbeResult(kind, Array.Empty<string>(), "x"), Wb);
        Assert.Equal(AutoAction.Quarantine, d.Action);
    }

    [Fact]
    public void Auto_OnlySessionWorkbook_Kills_CaseInsensitive()
    {
        var d = RecoveryPolicy.DecideAuto(false, IdentityState.Match, Responsive(@"c:\work\BOOK.XLSM"), Wb);
        Assert.Equal(AutoAction.Kill, d.Action);
    }

    [Fact]
    public void Auto_NoWorkbooks_OrWindowless_Kills()
    {
        Assert.Equal(AutoAction.Kill, RecoveryPolicy.DecideAuto(false, IdentityState.Match, Responsive(), Wb).Action);
        Assert.Equal(AutoAction.Kill, RecoveryPolicy.DecideAuto(false, IdentityState.Match,
            new ProbeResult(ProbeKind.Windowless, Array.Empty<string>()), Wb).Action);
    }

    [Fact]
    public void Auto_ForeignWorkbook_Quarantines_AndNamesIt()
    {
        var d = RecoveryPolicy.DecideAuto(false, IdentityState.Match,
            Responsive(Wb, @"C:\Users\u\Documents\Budget.xlsx"), Wb);
        Assert.Equal(AutoAction.Quarantine, d.Action);
        Assert.Contains("Budget.xlsx", d.Why);
    }

    [Fact]
    public void Auto_SameNameDifferentFolder_IsForeign()
    {
        // SaveAs 오탈출 사본도 "세션 워크북"이 아니다 — 사람 작업일 수 있어 보호
        var d = RecoveryPolicy.DecideAuto(false, IdentityState.Match, Responsive(@"C:\Other\Book.xlsm"), Wb);
        Assert.Equal(AutoAction.Quarantine, d.Action);
    }

    [Fact]
    public void Clear_ProcessGone_Clears()
    {
        Assert.Equal(ClearAction.Clear, RecoveryPolicy.DecideClear(false, IdentityState.Gone, force: false).Action);
    }

    [Theory]
    [InlineData(IdentityState.Match)]
    [InlineData(IdentityState.Unknown)]
    public void Clear_ProcessAlive_RefusesWithoutForce(IdentityState id)
    {
        Assert.Equal(ClearAction.Refuse, RecoveryPolicy.DecideClear(false, id, force: false).Action);
        Assert.Equal(ClearAction.ForceClear, RecoveryPolicy.DecideClear(false, id, force: true).Action);
    }

    [Fact]
    public void Clear_CorruptMarker_RequiresForce()
    {
        // 손상 마커는 프로세스가 없어 보여도(pid 0 = Gone) 해제하지 않는다 — fail closed
        Assert.Equal(ClearAction.Refuse, RecoveryPolicy.DecideClear(true, IdentityState.Gone, force: false).Action);
        Assert.Equal(ClearAction.ForceClear, RecoveryPolicy.DecideClear(true, IdentityState.Gone, force: true).Action);
    }

    [Fact]
    public void Discard_External_DetachesOnly()
    {
        var d = RecoveryPolicy.DecideDiscard(true, IdentityState.Match, Responsive(Wb), Wb);
        Assert.Equal(DiscardAction.Detach, d.Action);
    }

    [Fact]
    public void Discard_Gone_CleansWithoutKill()
    {
        Assert.Equal(DiscardAction.CleanOnly, RecoveryPolicy.DecideDiscard(false, IdentityState.Gone, null, Wb).Action);
    }

    [Fact]
    public void Discard_UnknownIdentity_Refuses()
    {
        Assert.Equal(DiscardAction.Refuse, RecoveryPolicy.DecideDiscard(false, IdentityState.Unknown, null, Wb).Action);
    }

    [Fact]
    public void Discard_ForeignWorkbookVisible_Refuses()
    {
        var d = RecoveryPolicy.DecideDiscard(false, IdentityState.Match, Responsive(Wb, @"C:\x\Other.xlsx"), Wb);
        Assert.Equal(DiscardAction.Refuse, d.Action);
    }

    [Theory]
    [InlineData(ProbeKind.Busy)]
    [InlineData(ProbeKind.Unresponsive)]
    [InlineData(ProbeKind.Windowless)]
    public void Discard_ToolSession_ExplicitCommand_KillsWhenNotProvablyForeign(ProbeKind kind)
    {
        // 명시 명령: 확인 불가는 거부 사유가 아니다(그것이 discard의 존재 이유)
        var d = RecoveryPolicy.DecideDiscard(false, IdentityState.Match,
            new ProbeResult(kind, Array.Empty<string>()), Wb);
        Assert.Equal(DiscardAction.Kill, d.Action);
    }
}
