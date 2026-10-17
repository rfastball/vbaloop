namespace VbaLoop.Core;

/// <summary>
/// 러너 테스트 1건의 판정 매트릭스 (pure — 단위 테스트 대상, RunnerMain에서 추출).
/// 세션 사망(sessionDead) 시 자기 귀속 모달이 없는 에러는 죽은 세션에 던진 호출의
/// 오탐(1004 "찾을 수 없습니다" 등)이므로 fail이 아니라 skipped — 진짜 신호(모달 본문)가
/// 오탐에 묻히는 것 방지(2026-07-08 실측 버그).
/// </summary>
public static class RunnerVerdict
{
    public const string Pass = "pass";
    public const string Fail = "fail";
    public const string Skipped = "skipped";

    public static (string Result, string? Detail) Decide(
        string? runError, IReadOnlyList<string> failedAsserts, string? modalDetail, bool sessionDead)
    {
        // 모달 본문이 실제 오류 메시지 — 세션이 죽었어도 이 테스트가 진짜 신호이므로 fail 유지
        if (modalDetail != null)
            return (Fail, modalDetail + (runError != null ? $" | COM: {runError}" : ""));
        if (sessionDead && runError != null)
            return (Skipped, $"not run because the session ended (error was a false positive: {runError})");
        if (runError != null)
            return (Fail, $"runtime error: {runError}");
        if (failedAsserts.Count > 0)
            return (Fail, string.Join(" | ", failedAsserts.Take(5))
                          + (failedAsserts.Count > 5 ? $" and {failedAsserts.Count - 5} more" : ""));
        return (Pass, null);
    }
}
