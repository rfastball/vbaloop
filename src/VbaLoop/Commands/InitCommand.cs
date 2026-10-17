using System.Reflection;
using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Commands;

/// <summary>
/// init: 온보딩 부트스트랩 + 드리프트 수렴 (v1.1). Excel/세션 불필요(순수 파일시스템).
/// 기본은 정본으로 즉시 수렴(멱등), --check는 무변경 검사(비정상 시 exit 2).
/// 관리 스탬프(설계 §12): exe가 프로젝트 스탬프보다 오래되면 쓰지 않는다(exit 3, hint vbaloop update) -
/// 오래된 exe의 정본으로 "드리프트"를 수렴하면 프로젝트가 다운그레이드되기 때문이다.
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

        var exe = ExeVersion.Current;
        var stamp = ManagedStamp.ReadFile(tomlPath);
        var stampState = ManagedStamp.Decide(exe, stamp);
        if (!ManagedStamp.AllowsWrite(stampState))
            return Refuse(o, opts.Check ? "check" : "apply", root, stamp!, stampState);

        var shimLines = LoadAssetLines("VbaLoop.assets.LogShim.bas");
        var docsBody = LoadAssetLines("VbaLoop.assets.AgentDocsBlock.md");
        var skillLines = LoadAssetLines("VbaLoop.assets.SkillMd");
        var pitfallLines = LoadAssetLines("VbaLoop.assets.PitfallsMd");
        var distLines = LoadAssetLines("VbaLoop.assets.DistMd");

        var items = Onboarding.Analyze(root, tomlPath, shimLines, docsBody, skillLines, pitfallLines, distLines,
            ManagedStamp.StampToWrite(exe, stampState));

        if (opts.Check)
        {
            foreach (var item in items)
                o.HumanLines.Add($"[{StatusLabel(item.Status)}] {item.Name}" + (item.Detail != null ? $" ({item.Detail})" : ""));
            var bad = items.Count(i => i.Status != ArtifactStatus.Ok);
            o.Payload = new
            {
                mode = "check",
                root,
                state = bad > 0 ? "drift" : "ok",
                managedBy = stamp,
                exeVersion = ExeVersion.Text,
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
            o.HumanLines.Add($"[{ActionLabel(action)}] {item.Name}"
                + (action != ArtifactAction.Unchanged && item.Detail != null ? $" ({item.Detail})" : ""));
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
            managedBy = ManagedStamp.ReadFile(Path.Combine(root, "vbaloop.toml")) ?? stamp,
            exeVersion = ExeVersion.Text,
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
    /// 관리 스탬프가 exe보다 새로운 프로젝트는 건너뛰고 [REFUSED]로 보고한다(exit 3, hint vbaloop update).
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

        var exe = ExeVersion.Current;
        var rows = new List<object>();
        int failed = 0, drifted = 0, shimUpdated = 0, refused = 0;
        foreach (var (name, root) in projects)
        {
            var tomlPath = ProjectConfig.FindTomlPath(Path.Combine(root, "vbaloop.toml"), root);
            var stamp = ManagedStamp.ReadFile(tomlPath);
            var stampState = ManagedStamp.Decide(exe, stamp);
            if (!ManagedStamp.AllowsWrite(stampState))
            {
                refused++;
                o.HumanLines.Add(stampState == StampState.ExeOlder
                    ? $"[REFUSED] {name}: managed_by {stamp} > exe {ExeVersion.Text}"
                    : $"[REFUSED] {name}: managed_by 값이 semver가 아님 ({stamp})");
                rows.Add(new { name, root, ok = false, state = ManagedStamp.Label(stampState), managedBy = stamp,
                    error = (string?)(stampState == StampState.ExeOlder
                        ? "관리 스탬프가 exe보다 새로움 - 쓰지 않음" : "관리 스탬프가 semver가 아님 - 쓰지 않음") });
                continue;
            }
            try
            {
                var items = Onboarding.Analyze(root, tomlPath, shimLines, docsBody, skillLines, pitfallLines, distLines,
                    ManagedStamp.StampToWrite(exe, stampState));
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

        o.Payload = new { mode = opts.Check ? "all-check" : "all", projectsRoot, exeVersion = ExeVersion.Text, projects = rows };
        if (failed > 0)
            o.Fail($"프로젝트 {failed}건 수렴 실패", "위 [FAIL] 라인의 원인 해결 후 재실행");
        if (refused > 0)
            o.Fail($"프로젝트 {refused}건 거부 - [REFUSED] 라인의 프로젝트는 실행 exe({ExeVersion.Text})보다 새 vbaloop가 관리 중",
                failed > 0 ? o.Hint : ManagedStamp.UpdateHint, environment: true);
        if (failed > 0 || refused > 0)
            return o;
        if (opts.Check && drifted > 0)
            o.Fail($"프로젝트 {drifted}건 드리프트", "vbaloop init --all");
        else
            o.HumanLines.Add($"전 프로젝트 {(opts.Check ? "온보딩 상태 정상" : "수렴 완료")} ({projects.Count}개)");
        return o;
    }

    /// <summary>
    /// 스탬프가 쓰기를 막는 상태(exe-older/invalid)의 보고. 아티팩트 비교는 하지 않는다 - 더 새 정본과
    /// 이 exe의 정본을 바이트 비교한 "드리프트"는 의미가 없고, 그걸 고치라는 신호가 다운그레이드를 부른다.
    /// exe-older는 환경 오류(exit 3, hint vbaloop update), invalid는 toml 확인이 필요한 일반 오류(exit 2).
    /// </summary>
    private static Outcome Refuse(Outcome o, string mode, string root, string stamp, StampState state)
    {
        bool older = state == StampState.ExeOlder;
        o.HumanLines.Add(older
            ? $"[EXE-OLDER] vbaloop.toml managed_by {stamp} > 실행 exe {ExeVersion.Text}"
            : $"[INVALID] vbaloop.toml managed_by = \"{stamp}\" (semver 아님)");
        if (older)
            o.HumanLines.Add("관리 파일을 쓰지 않았음 - vbaloop를 갱신해야 함 (init으로 고치지 말 것)");
        o.Payload = new
        {
            mode,
            root,
            state = ManagedStamp.Label(state),
            managedBy = stamp,
            exeVersion = ExeVersion.Text,
            artifacts = Array.Empty<object>(),
        };
        if (older)
            o.Fail(ManagedStamp.DescribeExeOlder(stamp, ExeVersion.Text), ManagedStamp.UpdateHint, environment: true);
        else
            o.Fail($"vbaloop.toml [vbaloop] managed_by 값이 semver가 아님: {stamp}",
                "vbaloop.toml의 managed_by 줄을 확인(지우면 다음 init이 다시 기록)");
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
