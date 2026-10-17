namespace VbaLoop.Core;

/// <summary>
/// 개명(vbaflow → vbaloop, 2026-10-17) 이전 산출물 이름. 기존 온보딩 프로젝트를 깨지 않고 읽으며,
/// <c>vbaloop init</c>이 새 이름으로 이전한다. 새 산출물에 이 이름을 쓰지 않는다.
/// </summary>
public static class LegacyNames
{
    public const string TomlFile = "vbaflow.toml";
    public const string ManagedStart = "<!-- vbaflow:managed:start -->";
    public const string ManagedEnd = "<!-- vbaflow:managed:end -->";
    public const string GitignoreStart = "# vbaflow:managed:start";
    public const string GitignoreEnd = "# vbaflow:managed:end";
    /// <summary>마커 도입 전 수동 작성 시절의 관리 섹션 헤딩.</summary>
    public const string DocsHeading = "## 개발 루프 (_VbaFlow)";
    public const string SkillDirName = "vbaflow";
    public const string SessionEnv = "VBAFLOW_SESSION_ID";

    /// <summary>vbaflow가 벤더링하던 스킬 파일(스킬 디렉터리 기준 상대). 이전 시 이것만 지운다 —
    /// 사용자가 같은 폴더에 둔 다른 파일은 보존.</summary>
    public static readonly string[] SkillFiles =
    {
        "SKILL.md",
        Path.Combine("references", "vba-pitfalls.md"),
        Path.Combine("references", "dist.md"),
    };
}
