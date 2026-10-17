using VbaLoop.Core;
using Xunit;
using static VbaLoop.Core.PushRecovery;

namespace VbaLoop.Tests;

public class PushRecoveryTests
{
    private static Touch T(string m, TouchKind k) => new(m, k);

    [Fact]
    public void NoTouches_NothingToDo_Untouched()
    {
        var plan = Decide(Array.Empty<Touch>(), Phase.Inject);
        Assert.False(plan.Poison);
        Assert.Empty(plan.RestoreOrder);
        Assert.Equal(WorkbookState.Untouched, Resolve(plan, 0, 0));
    }

    [Fact]
    public void RestorableTouches_RestoreAllInReverse()
    {
        // modA 교체 성공 → modB 신규 성공 → modC 교체 중 실패: 세 모듈 전부 역순 원복 (실패 모듈 포함)
        var touches = new[]
        {
            T("modA", TouchKind.CodeReplaced),
            T("modB", TouchKind.ComponentAdded),
            T("modC", TouchKind.CodeReplaced),
        };
        var plan = Decide(touches, Phase.Inject);
        Assert.False(plan.Poison);
        Assert.Equal(new[] { "modC", "modB", "modA" }, plan.RestoreOrder.Select(t => t.Module));
        Assert.Equal(WorkbookState.Restored, Resolve(plan, touches.Length, restoreFailures: 0));
    }

    [Fact]
    public void AnyRestoreFailure_Poisons()
    {
        var touches = new[] { T("modA", TouchKind.CodeReplaced), T("modB", TouchKind.CodeReplaced) };
        var plan = Decide(touches, Phase.Inject);
        Assert.Equal(WorkbookState.Poisoned, Resolve(plan, touches.Length, restoreFailures: 1));
    }

    [Fact]
    public void FatalCom_Poisons_EvenIfPlanWasRestorable()
    {
        var touches = new[] { T("modA", TouchKind.CodeReplaced) };
        var plan = Decide(touches, Phase.Inject);
        Assert.Equal(WorkbookState.Poisoned, Resolve(plan, 1, 0, poisonedByFatalCom: true));
    }

    [Theory]
    [InlineData(TouchKind.FormReimported, "frmMain(폼 재Import)")]
    [InlineData(TouchKind.ComponentPruned, "frmMain(prune 삭제)")]
    public void UnrestorableTouch_PoisonsWithoutAttemptingRestore(TouchKind kind, string label)
    {
        var touches = new[] { T("modA", TouchKind.CodeReplaced), T("frmMain", kind) };
        var plan = Decide(touches, Phase.Inject);
        Assert.True(plan.Poison);
        Assert.Empty(plan.RestoreOrder); // 일부만 되돌리는 것은 무의미 — 어차피 Excel을 버린다
        Assert.Contains(label, plan.PoisonReason);
        Assert.Equal(WorkbookState.Poisoned, Resolve(plan, touches.Length, 0));
    }

    [Fact]
    public void PrunePhaseFailure_Poisons()
    {
        var touches = new[] { T("modA", TouchKind.CodeReplaced), T("modOld", TouchKind.ComponentPruned) };
        Assert.True(Decide(touches, Phase.Prune).Poison);
    }

    [Fact]
    public void IntegrityFailure_AlwaysPoisons()
    {
        Assert.True(Decide(new[] { T("modA", TouchKind.CodeReplaced) }, Phase.Integrity).Poison);
        Assert.True(Decide(Array.Empty<Touch>(), Phase.Integrity).Poison);
    }

    [Fact]
    public void SameModuleTouchedTwice_CollapsesToWorst_KeepsFirstOrder()
    {
        var touches = new[]
        {
            T("modA", TouchKind.CodeReplaced),
            T("modB", TouchKind.ComponentAdded),
            T("MODB", TouchKind.CodeReplaced), // 신규 생성 후 본문 주입 — 제거가 둘 다 되돌린다
        };
        var plan = Decide(touches, Phase.Inject);
        Assert.Equal(2, plan.RestoreOrder.Count);
        Assert.Equal(T("modB", TouchKind.ComponentAdded), plan.RestoreOrder[0]);
        Assert.Equal(T("modA", TouchKind.CodeReplaced), plan.RestoreOrder[1]);
    }

    [Fact]
    public void ReimportAfterReplace_SameModule_IsUnrestorable()
    {
        var plan = Decide(new[] { T("frmA", TouchKind.CodeReplaced), T("frmA", TouchKind.FormReimported) }, Phase.Inject);
        Assert.True(plan.Poison);
    }

    // ---- 사용자 문구: 실제 상태만 말한다 ----

    [Fact]
    public void Describe_Restored_MentionsCountAndRollbackWithBackupId()
    {
        var (summary, hint) = Describe(WorkbookState.Restored, 3, "20261017-031500-123-push", null);
        Assert.Contains("3개", summary);
        Assert.Contains("저장 안 함", summary);
        Assert.Contains("rollback --latest", hint);
        Assert.Contains("20261017-031500-123-push", hint);
        Assert.Contains("--no-save", hint);
    }

    [Fact]
    public void Describe_Poisoned_SaysDiskIsPrePushAndNeverClaimsRestore()
    {
        var (summary, hint) = Describe(WorkbookState.Poisoned, 0, "20261017-031500-123-push", "원복 실패: modA");
        Assert.Contains("poison", summary);
        Assert.Contains("이번 push 이전 상태", summary);
        Assert.Contains("원복 실패: modA", summary);
        Assert.DoesNotContain("원복·검증됨", summary);
        Assert.Contains("rollback --latest", hint);
    }

    [Fact]
    public void Describe_NoBackup_SaysRollbackUnavailable()
    {
        var (_, hint) = Describe(WorkbookState.Poisoned, 0, backupId: null, poisonReason: null);
        Assert.Contains("keep = 0", hint);
        Assert.DoesNotContain("rollback --latest", hint);
    }

    [Fact]
    public void Describe_Untouched_NoRollbackNoise()
    {
        var (summary, hint) = Describe(WorkbookState.Untouched, 0, "id", null);
        Assert.Contains("변경 없음", summary);
        Assert.DoesNotContain("rollback", hint);
    }

    // ---- 스냅샷 검증 ----

    [Theory]
    [InlineData("Sub A()\r\nEnd Sub", "Sub A()\r\nEnd Sub")]
    [InlineData("Sub A()\r\nEnd Sub", "Sub A()\r\nEnd Sub\r\n\r\n")]  // AddFromString 말미 빈 줄
    [InlineData("Dim myVar\r\nSub A()", "Dim MyVar\r\nSub A()")]      // VBE 재케이싱
    [InlineData("", "")]
    [InlineData("", "\r\n")]
    public void SnapshotMatches_ToleratesBenignDifferences(string expected, string actual)
        => Assert.True(SnapshotMatches(expected, actual));

    [Theory]
    [InlineData("Sub A()\r\nEnd Sub", "Sub B()\r\nEnd Sub")]
    [InlineData("Sub A()\r\nEnd Sub", "Sub A()\r\nEnd Sub\r\n()")]   // 팬텀 잔존 = 실패
    [InlineData("Sub A()\r\nEnd Sub", "")]                           // 빈 채로 남음
    [InlineData("Sub A()\r\n\r\nEnd Sub", "Sub A()\r\nEnd Sub")]     // 중간 빈 줄은 의미 있음
    public void SnapshotMatches_RejectsRealDifferences(string expected, string actual)
        => Assert.False(SnapshotMatches(expected, actual));
}
