namespace VbaLoop.Core;

/// <summary>
/// setup이 고른 프로젝트 폴더와 기존 프로젝트의 관계(설계 §10.2). init·push·test는 cwd에서 위로 올라가며
/// 처음 만나는 vbaloop.toml을 프로젝트 루트로 삼는다 — 그래서 다른 프로젝트 안쪽 폴더에 새 프로젝트를 세우면
/// 바깥 프로젝트의 설정·워크북이 대신 잡힌다. 그런 폴더는 변경 전에 거부하고, 이미 프로젝트인 폴더는
/// 설정된 워크북을 그대로 채택한다.
/// </summary>
public static class ProjectRoot
{
    /// <summary>folder 자신이 아니라 그 상위에 있는 프로젝트 루트. folder에 자기 vbaloop.toml이 있거나 상위에 없으면 null.</summary>
    public static string? Enclosing(string folder)
    {
        string full = Path.GetFullPath(folder);
        var toml = ProjectConfig.FindTomlPath(null, full);
        if (toml == null) return null;
        string root = Path.GetDirectoryName(toml)!;
        return PathDrift.SamePath(root, full) ? null : root;
    }

    /// <summary>다른 프로젝트 안쪽이면 CliError — 어떤 파일도 만들기 전에 부른다.</summary>
    public static void EnsureNotNested(string folder)
    {
        string? outer = Enclosing(folder);
        if (outer != null)
            throw new CliError($"다른 vbaloop 프로젝트 안쪽에는 새 프로젝트를 만들 수 없음: {Path.GetFullPath(folder)} (바깥 프로젝트: {outer})",
                "바깥 프로젝트 폴더 밖에서 다시 실행하거나 --folder로 다른 위치를 지정");
    }

    /// <summary>
    /// folder가 이미 프로젝트(자기 vbaloop.toml 보유)면 설정된 워크북의 루트 기준 상대 경로, 아니면 null.
    /// 설정이 가리키는 워크북이 없거나 프로젝트 밖이면 CliError — 새 워크북을 지어내지 않는다.
    /// </summary>
    public static string? ConfiguredWorkbook(string folder)
    {
        string full = Path.GetFullPath(folder);
        var toml = ProjectConfig.FindTomlPath(null, full);
        if (toml == null || !PathDrift.SamePath(Path.GetDirectoryName(toml)!, full)) return null;

        var cfg = ProjectConfig.Load(toml, full);
        string rel = Path.GetRelativePath(full, cfg.WorkbookPath);
        if (rel.StartsWith("..") || Path.IsPathRooted(rel))
            throw new CliError($"vbaloop.toml의 워크북이 프로젝트 폴더 밖을 가리킴: {cfg.WorkbookPath}",
                "[excel].workbook을 프로젝트 폴더 기준 상대 경로로 고친 뒤 다시 실행");
        if (!File.Exists(cfg.WorkbookPath))
            throw new CliError($"vbaloop.toml이 가리키는 워크북이 없음: {cfg.WorkbookPath}",
                "[excel].workbook 경로를 확인하거나 워크북을 그 위치에 둔 뒤 다시 실행");
        return rel;
    }
}
