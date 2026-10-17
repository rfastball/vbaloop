namespace VbaLoop.Core;

/// <summary>
/// pull 덮어쓰기 가드(설계 §4 pull): 현재 src 지문이 마지막 push/pull 지문 어느 쪽과도
/// 다르면 동기화된 적 없는 로컬 편집이 있다는 뜻 — 무경고 덮어쓰기를 차단한다.
/// 이력이 전혀 없으면(최초 pull, 일회성 정규화 경로) 통과.
/// </summary>
public static class PullGuard
{
    /// <summary>차단 사유를 반환, 통과면 null.</summary>
    public static string? FindBlockReason(string current, string? lastPush, string? lastPull)
    {
        if (lastPush == null && lastPull == null) return null;
        if (current == lastPush || current == lastPull) return null;
        return "src에 동기화되지 않은 편집이 있음(지문이 마지막 push/pull 어느 쪽과도 불일치)";
    }
}
