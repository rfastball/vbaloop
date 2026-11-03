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

    private static ProbeResult Uncertain(ProbeKind kind, params string[] titles)
        => new(kind, Array.Empty<string>(), "x") { WindowTitles = titles };

    [Theory]
    [InlineData(ProbeKind.Busy)]
    [InlineData(ProbeKind.Unresponsive)]
    public void Auto_BusyEphemeral_OnlySessionWindows_Kills(ProbeKind kind)
    {
        // 2026-11-03 시간 초과: 매크로 실행 중인 읽기 전용 일회용 Excel — COM은 실패해도 창 제목이 근거(설계 §14.3)
        var d = RecoveryPolicy.DecideAuto(false, IdentityState.Match,
            Uncertain(kind, "Book.xlsm  [읽기 전용]"), Wb, ephemeral: true);
        Assert.Equal(AutoAction.Kill, d.Action);
    }

    [Fact]
    public void Auto_BusyPersistent_StaysQuarantined_EvenWithOnlySessionWindows()
    {
        // 유지 세션의 세션 워크북은 미저장 사람 작업을 담을 수 있다 — 창 제목 근거는 일회용에만
        var d = RecoveryPolicy.DecideAuto(false, IdentityState.Match,
            Uncertain(ProbeKind.Unresponsive, "Book.xlsm"), Wb, ephemeral: false);
        Assert.Equal(AutoAction.Quarantine, d.Action);
    }

    [Fact]
    public void Auto_BusyEphemeral_ForeignOrUnknownWindow_Quarantines_AndNamesWindows()
    {
        // 같은 인스턴스에 열린 두 번째 워크북(숨은 창 포함 실측), Caption을 바꾼 창, 창 제목 없음
        var d = RecoveryPolicy.DecideAuto(false, IdentityState.Match,
            Uncertain(ProbeKind.Unresponsive, "Book.xlsm  [읽기 전용]", "통합 문서1"), Wb, ephemeral: true);
        Assert.Equal(AutoAction.Quarantine, d.Action);
        Assert.Contains("통합 문서1", d.Why);
        Assert.Equal(AutoAction.Quarantine, RecoveryPolicy.DecideAuto(false, IdentityState.Match,
            Uncertain(ProbeKind.Busy, "Renamed"), Wb, ephemeral: true).Action);
        Assert.Equal(AutoAction.Quarantine, RecoveryPolicy.DecideAuto(false, IdentityState.Match,
            Uncertain(ProbeKind.Unresponsive), Wb, ephemeral: true).Action);
    }

    [Fact]
    public void Auto_ExternalEphemeralFlag_NeverKills()
    {
        Assert.Equal(AutoAction.Quarantine, RecoveryPolicy.DecideAuto(true, IdentityState.Match,
            Uncertain(ProbeKind.Busy, "Book.xlsm"), Wb, ephemeral: true).Action);
    }

    [Theory]
    [InlineData("Book.xlsm", true)]
    [InlineData("BOOK.XLSM", true)]
    [InlineData("Book.xlsm  [읽기 전용]", true)]
    [InlineData("Book.xlsm:2  [읽기 전용]", true)]
    [InlineData("Book.xlsm:1", true)]
    [InlineData("Book.xlsm  [읽기 전용]  [호환 모드]", true)]
    [InlineData("Book.xlsm.bak", false)]
    [InlineData("Book.xlsm - 복사본", false)]
    [InlineData("Book.xlsm:x", false)]
    [InlineData("Renamed", false)]
    [InlineData("", false)]
    public void SessionWindowTitle_Matching(string title, bool expected)
    {
        Assert.Equal(expected, RecoveryPolicy.IsSessionWindowTitle(title, Wb));
    }

    [Fact]
    public void SessionWindowTitle_BracketInsideName_NotStripped()
    {
        // 이름 자체의 대괄호는 꼬리로 취급하지 않는다 — 이름 전체가 접두로 일치해야 한다
        const string bracketed = @"C:\Work\Report [v2].xlsm";
        Assert.True(RecoveryPolicy.IsSessionWindowTitle("Report [v2].xlsm  [읽기 전용]", bracketed));
        Assert.False(RecoveryPolicy.IsSessionWindowTitle("Report.xlsm", bracketed));
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
