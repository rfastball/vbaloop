using System.Reflection;
using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Commands;

/// <summary>
/// init: 온보딩 부트스트랩 + 드리프트 수렴 (v1.1). Excel/세션 불필요(순수 파일시스템).
/// 기본은 정본으로 즉시 수렴(멱등), --check는 무변경 검사(비정상 시 exit 2).
/// </summary>
public static class InitCommand
{
    public static Outcome Run(CliArgs opts)
    {
        if (opts.All) return RunAll(opts);

        var o = new Outcome { Command = "init" };
        var cwd = Environment.CurrentDirectory;

        // 루트 결정: 상위 탐색으로 toml을 찾으면 그 위치가 루트(기존 명령과 동일 의미론,
        // src/에서 실행해도 중첩 스캐폴드 방지). 없으면 cwd(또는 --config의 디렉토리)에 스캐폴드.
        var tomlPath = ProjectConfig.FindTomlPath(opts.Config, cwd);
        string root = tomlPath != null
            ? Path.GetDirectoryName(tomlPath)!
            : opts.Config != null
                ? Path.GetDirectoryName(Path.GetFullPath(opts.Config))!
                : Path.GetFullPath(cwd);

        var shimLines = LoadAssetLines("VbaLoop.assets.LogShim.bas");
        var docsBody = LoadAssetLines("VbaLoop.assets.AgentDocsBlock.md");
        var skillLines = LoadAssetLines("VbaLoop.assets.SkillMd");
        var pitfallLines = LoadAssetLines("VbaLoop.assets.PitfallsMd");
        var distLines = LoadAssetLines("VbaLoop.assets.DistMd");

        var items = Onboarding.Analyze(root, tomlPath, shimLines, docsBody, skillLines, pitfallLines, distLines);

        if (opts.Check)
        {
            foreach (var item in items)
                o.HumanLines.Add($"[{StatusLabel(item.Status)}] {item.Name}");
            var bad = items.Count(i => i.Status != ArtifactStatus.Ok);
            o.Payload = new
            {
                mode = "check",
                root,
                artifacts = items.Select(i => new
                {
                    name = i.Name,
                    path = i.Path,
                    status = i.Status.ToString().ToLowerInvariant(),
                }).ToArray(),
            };
            if (bad > 0)
                o.Fail($"온보딩 아티팩트 {bad}건 비정상", "vbaloop init");
            else
                o.HumanLines.Add("온보딩 상태 정상");
            return o;
        }

        var results = Onboarding.Apply(items);
        foreach (var (item, action) in results)
            o.HumanLines.Add($"[{ActionLabel(action)}] {item.Name}");
        var shim = results.FirstOrDefault(r => r.Item.Path.EndsWith("LogShim.bas", StringComparison.OrdinalIgnoreCase));
        if (shim.Item != null && shim.Action != ArtifactAction.Unchanged)
            o.HumanLines.Add("LogShim.bas 갱신됨 - 다음 push에서 워크북에 반영 필요");
        var skill = results.FirstOrDefault(r => r.Item.Path.EndsWith("SKILL.md", StringComparison.OrdinalIgnoreCase));
        if (skill.Item != null && skill.Action != ArtifactAction.Unchanged)
            o.HumanLines.Add("SKILL.md 갱신됨 - 관리 사본(직접 수정 금지, 갱신은 vbaloop init)");
        o.Payload = new
        {
            mode = "apply",
            root,
            artifacts = results.Select(r => new
            {
                name = r.Item.Name,
                path = r.Item.Path,
                action = r.Action.ToString().ToLowerInvariant(),
            }).ToArray(),
        };
        return o;
    }

    /// <summary>
    /// --all: 온보딩된 전 프로젝트 일괄 수렴 (rollout.ps1의 3단계). exe 위치에서 레포를 찾아
    /// 형제 toml 프로젝트 + _pilot_form을 순회한다. 한 프로젝트 실패(깨진 toml 등)는 기록 후
    /// 계속 — 하나가 나머지 수렴을 인질 잡지 않는다. 실패/드리프트 존재 시 exit 2.
    /// toml 없는 폴더는 열거 자체가 제외하므로 여기서 스캐폴드는 일어나지 않는다.
    /// </summary>
    private static Outcome RunAll(CliArgs opts)
    {
        var o = new Outcome { Command = "init" };
        var exeDir = Environment.ProcessPath != null
            ? Path.GetDirectoryName(Environment.ProcessPath)! : Environment.CurrentDirectory;
        var repoRoot = RepoLocator.FindRepoRoot(exeDir)
            ?? throw new CliError("--all은 vbaloop 레포 아래의 exe에서만 동작 (VbaLoop.sln 상향 탐색 실패)",
                @"vbaloop에서 .\rollout.ps1 실행 (test -> publish -> init --all)");
        var projectsRoot = Path.GetDirectoryName(repoRoot)!;
        var projects = RepoLocator.EnumerateOnboardedProjects(projectsRoot, repoRoot);

        var shimLines = LoadAssetLines("VbaLoop.assets.LogShim.bas");
        var docsBody = LoadAssetLines("VbaLoop.assets.AgentDocsBlock.md");
        var skillLines = LoadAssetLines("VbaLoop.assets.SkillMd");
        var pitfallLines = LoadAssetLines("VbaLoop.assets.PitfallsMd");
        var distLines = LoadAssetLines("VbaLoop.assets.DistMd");

        var rows = new List<object>();
        int failed = 0, drifted = 0, shimUpdated = 0;
        foreach (var (name, root) in projects)
        {
            var tomlPath = ProjectConfig.FindTomlPath(Path.Combine(root, "vbaloop.toml"), root)
                ?? ProjectConfig.FindTomlPath(Path.Combine(root, LegacyNames.TomlFile), root); // 이전 전 프로젝트
            try
            {
                var items = Onboarding.Analyze(root, tomlPath, shimLines, docsBody, skillLines, pitfallLines, distLines);
                if (opts.Check)
                {
                    int bad = items.Count(i => i.Status != ArtifactStatus.Ok);
                    if (bad > 0) drifted++;
                    o.HumanLines.Add(bad == 0
                        ? $"[ok]    {name}"
                        : $"[DRIFT] {name} ({string.Join(", ", items.Where(i => i.Status != ArtifactStatus.Ok).Select(i => i.Name))})");
                    rows.Add(new { name, root, ok = bad == 0, drifted = bad, error = (string?)null });
                }
                else
                {
                    var results = Onboarding.Apply(items);
                    int created = results.Count(r => r.Action == ArtifactAction.Created);
                    int overwritten = results.Count(r => r.Action == ArtifactAction.Overwritten);
                    int unchanged = results.Count(r => r.Action == ArtifactAction.Unchanged);
                    if (results.Any(r => r.Action != ArtifactAction.Unchanged
                            && r.Item.Path.EndsWith("LogShim.bas", StringComparison.OrdinalIgnoreCase)))
                        shimUpdated++;
                    o.HumanLines.Add(created + overwritten == 0
                        ? $"[ok]      {name} (unchanged {unchanged})"
                        : $"[UPDATED] {name} (created {created}, overwritten {overwritten}, unchanged {unchanged})");
                    rows.Add(new { name, root, ok = true, created, overwritten, unchanged, error = (string?)null });
                }
            }
            catch (CliError e)
            {
                failed++;
                o.HumanLines.Add($"[FAIL]    {name}: {e.Message}");
                rows.Add(new { name, root, ok = false, error = (string?)e.Message });
            }
        }
        if (shimUpdated > 0)
            o.HumanLines.Add($"LogShim.bas 갱신 {shimUpdated}건 - 각 프로젝트의 다음 push에서 워크북 반영 필요");

        o.Payload = new { mode = opts.Check ? "all-check" : "all", projectsRoot, projects = rows };
        if (failed > 0)
            o.Fail($"프로젝트 {failed}건 수렴 실패", "위 [FAIL] 라인의 원인 해결 후 재실행");
        else if (opts.Check && drifted > 0)
            o.Fail($"프로젝트 {drifted}건 드리프트", "vbaloop init --all");
        else
            o.HumanLines.Add($"전 프로젝트 {(opts.Check ? "온보딩 상태 정상" : "수렴 완료")} ({projects.Count}개)");
        return o;
    }

    private static string StatusLabel(ArtifactStatus s) => s switch
    {
        ArtifactStatus.Ok => "ok",
        ArtifactStatus.Missing => "MISSING",
        _ => "DRIFT",
    };

    private static string ActionLabel(ArtifactAction a) => a switch
    {
        ArtifactAction.Created => "created",
        ArtifactAction.Overwritten => "overwritten",
        ArtifactAction.Migrated => "migrated",
        ArtifactAction.Removed => "removed",
        _ => "unchanged",
    };

    private static string[] LoadAssetLines(string logicalName)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(logicalName)
            ?? throw new CliError($"임베디드 자산 누락: {logicalName}", "vbaloop 재빌드 필요 (csproj EmbeddedResource)");
        using var reader = new StreamReader(stream);
        var text = reader.ReadToEnd();
        if (text.Length > 0 && text[0] == '\uFEFF') text = text[1..];
        return CanonicalText.SplitLines(text);
    }
}
