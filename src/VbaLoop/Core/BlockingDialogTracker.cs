namespace VbaLoop.Core;

/// <summary>
/// 실행을 막는 대화상자의 지속 판정(설계 §14.5) — 순수부. 실행 시작 전부터 있던 창(attach한 사용자 Excel의
/// 기존 대화상자)은 제외하고, 같은 창이 <see cref="PersistMs"/> 넘게 남아 있을 때만 막힘으로 본다.
/// 워치독이 해제하는 VBE 닫기 후속 프롬프트처럼 잠깐 뜨고 사라지는 창은 걸리지 않는다.
/// </summary>
public sealed class BlockingDialogTracker
{
    public const int DefaultPersistMs = 5000;

    private readonly HashSet<IntPtr> _preexisting;
    private readonly Dictionary<IntPtr, long> _firstSeen = new();

    public int PersistMs { get; }

    public BlockingDialogTracker(IEnumerable<IntPtr> preexisting, int persistMs = DefaultPersistMs)
    {
        _preexisting = new HashSet<IntPtr>(preexisting);
        PersistMs = persistMs;
    }

    /// <summary>지금 보이는 막는 창 목록을 관측한다. 지속 기준을 넘긴 창이 있으면 그 핸들, 없으면 null.</summary>
    public IntPtr? Observe(IEnumerable<IntPtr> current, long nowMs)
    {
        var now = new HashSet<IntPtr>(current.Where(h => !_preexisting.Contains(h)));
        foreach (var gone in _firstSeen.Keys.Where(h => !now.Contains(h)).ToList())
            _firstSeen.Remove(gone);
        foreach (var h in now)
        {
            if (!_firstSeen.TryGetValue(h, out long seen))
                _firstSeen[h] = nowMs;
            else if (nowMs - seen >= PersistMs)
                return h;
        }
        return null;
    }
}
