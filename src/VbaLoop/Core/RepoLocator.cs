namespace VbaLoop.Core;

/// <summary>
/// exe 위치에서 vbaloop 레포와 온보딩 프로젝트들을 찾는다 (init --all, bugreport 공용).
/// 마커는 레포 루트의 VbaLoop.sln — publish 산출물(src\...\publish)과 dev 빌드(bin\Debug)
/// 어느 쪽에서든 상향 탐색으로 도달한다. PublishSingleFile에서는 Assembly.Location이
/// 빈 문자열이므로 호출부는 Environment.ProcessPath 기준 디렉토리를 넘겨야 한다.
/// </summary>
public static class RepoLocator
{
    private const string RepoMarker = "VbaLoop.sln";

    /// <summary>startDir에서 상위로 VbaLoop.sln 보유 디렉토리 탐색. 없으면 null
    /// (ProjectConfig.FindTomlPath와 동일 패턴).</summary>
    public static string? FindRepoRoot(string startDir)
    {
        var dir = new DirectoryInfo(Path.GetFullPath(startDir));
        while (dir != null)
        {
            if (File.Exists(Path.Combine(dir.FullName, RepoMarker))) return dir.FullName;
            dir = dir.Parent;
        }
        return null;
    }

    /// <summary>
    /// 온보딩 프로젝트 열거: projectsRoot(레포 부모)의 1단계 하위 중 vbaloop.toml 보유
    /// 디렉토리(레포 자신 제외) + repoRoot\_pilot_form(레포 내 샘플, toml 보유 시).
    /// 별도 레지스트리 없음 — toml 존재가 곧 온보딩이다.
    /// </summary>
    public static List<(string Name, string Root)> EnumerateOnboardedProjects(
        string projectsRoot, string repoRoot)
    {
        var repoFull = Path.GetFullPath(repoRoot);
        var result = new List<(string, string)>();
        foreach (var dir in Directory.EnumerateDirectories(Path.GetFullPath(projectsRoot))
                     .OrderBy(d => d, StringComparer.OrdinalIgnoreCase))
        {
            if (string.Equals(Path.GetFullPath(dir), repoFull, StringComparison.OrdinalIgnoreCase))
                continue;
            if (HasToml(dir))
                result.Add((Path.GetFileName(dir), dir));
        }
        var pilot = Path.Combine(repoFull, "_pilot_form");
        if (HasToml(pilot))
            result.Add((Path.GetFileName(pilot), pilot));
        return result;
    }

    private static bool HasToml(string dir) => File.Exists(Path.Combine(dir, "vbaloop.toml"));
}
