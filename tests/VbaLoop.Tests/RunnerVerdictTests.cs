using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class RunnerVerdictTests
{
    private static readonly string[] NoAsserts = Array.Empty<string>();

    [Fact]
    public void CleanRun_Passes()
    {
        var (result, detail) = RunnerVerdict.Decide(null, NoAsserts, null, sessionDead: false);
        Assert.Equal(RunnerVerdict.Pass, result);
        Assert.Null(detail);
    }

    [Fact]
    public void RuntimeError_Fails()
    {
        var (result, detail) = RunnerVerdict.Decide("오버플로", NoAsserts, null, sessionDead: false);
        Assert.Equal(RunnerVerdict.Fail, result);
        Assert.Equal("런타임 에러: 오버플로", detail);
    }

    [Fact]
    public void FailedAsserts_Fail_WithTruncationTail()
    {
        var five = new[] { "a1: expected=1 actual=2", "a2", "a3", "a4", "a5" };
        var (r1, d1) = RunnerVerdict.Decide(null, five, null, sessionDead: false);
        Assert.Equal(RunnerVerdict.Fail, r1);
        Assert.DoesNotContain("외", d1);

        var seven = five.Concat(new[] { "a6", "a7" }).ToArray();
        var (r2, d2) = RunnerVerdict.Decide(null, seven, null, sessionDead: false);
        Assert.Equal(RunnerVerdict.Fail, r2);
        Assert.EndsWith("외 2건", d2);
    }

    [Fact]
    public void ModalDetail_TakesPrecedence_AndAppendsComError()
    {
        var (result, detail) = RunnerVerdict.Decide("1004", NoAsserts, "모달 자동 해제(확인): 컴파일 오류", false);
        Assert.Equal(RunnerVerdict.Fail, result);
        Assert.Equal("모달 자동 해제(확인): 컴파일 오류 | COM: 1004", detail);

        var (r2, d2) = RunnerVerdict.Decide(null, NoAsserts, "모달 자동 해제(확인): 컴파일 오류", false);
        Assert.Equal(RunnerVerdict.Fail, r2);
        Assert.Equal("모달 자동 해제(확인): 컴파일 오류", d2);
    }

    [Fact]
    public void DeadSession_ErrorWithoutOwnModal_IsSkippedMisfire()
    {
        // 2026-07-08 실측 버그: 죽은 세션에 이름만으로 Run → 1004 "찾을 수 없습니다" 오탐
        var (result, detail) = RunnerVerdict.Decide(
            "'C:\\Users\\x\\Documents\\A.xlsm'을(를) 찾을 수 없습니다.", NoAsserts, null, sessionDead: true);
        Assert.Equal(RunnerVerdict.Skipped, result);
        Assert.Contains("오탐", detail);
    }

    [Fact]
    public void DeadSession_WithOwnModal_KeepsFail_AsTrueSignal()
    {
        // 컴파일 에러 모달이 이 테스트에 귀속 — 세션이 죽었어도 이게 진짜 신호
        var (result, detail) = RunnerVerdict.Decide(null, NoAsserts, "모달 자동 해제(확인): 컴파일 오류", true);
        Assert.Equal(RunnerVerdict.Fail, result);
        Assert.StartsWith("모달 자동 해제", detail);
    }

    [Fact]
    public void DeadSession_NoErrorNoModal_StillPasses()
    {
        // 테스트 자체는 정상 완료 후 세션이 죽은 경우(이전 테스트의 잔여 모달 등) — 판정 유지
        var (result, _) = RunnerVerdict.Decide(null, NoAsserts, null, sessionDead: true);
        Assert.Equal(RunnerVerdict.Pass, result);
    }
}
