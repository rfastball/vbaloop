namespace VbaLoop.Core;

/// <summary>실행 중인 Excel 인스턴스에서 발견한 열린 워크북 한 건.</summary>
public sealed record AttachCandidate(int Pid, string FullName, bool ReadOnly);

public enum AttachPickKind { Found, NotOpen, ReadOnlyOnly, Ambiguous }

public sealed record AttachPick(AttachPickKind Kind, AttachCandidate? Target, IReadOnlyList<string> Suspects);

/// <summary>
/// session attach의 대상 선택 순수부(설계 §11 L5). 정본(cfg.WorkbookPath)과 FullName이 경로로 같은
/// (대소문자·구분자·상대 세그먼트 무시) 쓰기 가능 워크북이 정확히 한 프로세스에 있을 때만 채택한다.
/// read-only로만 열려 있으면 채택하지 않는다 — push의 Save가 SaveAs 오탈출하는 split-brain을
/// 막기 위해서다. 정본이 없으면 동명 타경로/URL(OneDrive FullName) 의심 목록을 돌려준다.
/// </summary>
public static class AttachMatch
{
    public static bool IsCanonical(string fullName, string configuredPath)
        => PathDrift.SamePath(fullName, configuredPath);

    public static AttachPick Pick(IEnumerable<AttachCandidate> candidates, string configuredPath)
    {
        var all = candidates.ToList();
        var matches = all.Where(c => IsCanonical(c.FullName, configuredPath)).ToList();
        var writable = matches.Where(c => !c.ReadOnly).ToList();

        if (writable.Select(c => c.Pid).Distinct().Count() > 1)
            return new(AttachPickKind.Ambiguous, null,
                writable.Select(c => $"PID {c.Pid}: {c.FullName}").ToList());
        if (writable.Count > 0)
            return new(AttachPickKind.Found, writable[0], Array.Empty<string>());
        if (matches.Count > 0)
            return new(AttachPickKind.ReadOnlyOnly, matches[0], Array.Empty<string>());
        return new(AttachPickKind.NotOpen, null,
            PathDrift.FindDriftSuspects(all.Select(c => c.FullName), configuredPath));
    }
}
