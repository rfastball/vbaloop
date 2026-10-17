namespace VbaLoop.Core;

/// <summary>
/// 셸에서 인자 없이 setup을 부른 경우(설치 스크립트가 이어서 부르는 경우 포함)의 현재 폴더 판독(설계 §10.2).
/// "워크북이 있는 폴더에서 한 줄 실행" 또는 "빈 폴더에서 한 줄 실행"만으로 시작점이 정해지게 한다 —
/// 결정은 여기서 하지 않고 마법사가 이 결과로 사용자에게 먼저 묻는다.
/// </summary>
public static class CwdStart
{
    /// <param name="IsProject">이미 vbaloop 프로젝트(vbaloop.toml)다 — 다시 세팅을 제안.</param>
    /// <param name="Workbooks">최상위 *.xlsm(잠금 스텁 ~$ 제외), 이름순.</param>
    /// <param name="IsEmpty">점(.)으로 시작하는 항목(.git, .gitignore 등) 말고는 비어 있다 — 새 프로젝트 후보.</param>
    public sealed record State(bool IsProject, IReadOnlyList<string> Workbooks, bool IsEmpty);

    public static State Inspect(string dir)
    {
        if (!Directory.Exists(dir)) return new State(false, Array.Empty<string>(), true);

        bool isProject = File.Exists(Path.Combine(dir, "vbaloop.toml"));
        var workbooks = Directory.EnumerateFiles(dir, "*.xlsm", SearchOption.TopDirectoryOnly)
            .Select(p => Path.GetFileName(p))
            .Where(n => !n.StartsWith("~$"))
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToList();
        bool isEmpty = !Directory.EnumerateFileSystemEntries(dir)
            .Any(p => !Path.GetFileName(p).StartsWith('.'));
        return new State(isProject, workbooks, isEmpty);
    }
}
