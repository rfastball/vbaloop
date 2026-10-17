namespace VbaLoop.Core;

/// <summary>
/// <c>setup --plan</c>: 아무것도 바꾸지 않고 setup이 할 일을 미리 계산한다(설계 §10.8). 에이전트가
/// 이 결과를 사용자에게 보여 동의를 받은 뒤 <see cref="Result.Command"/>를 실행하는 것이 계약이다.
/// 레지스트리·PATH 같은 환경 프로브는 호출부(SetupCommand)가 채워 넘기고, 여기서는 판단만 한다 —
/// 파일시스템은 읽기만 한다(WorkbookIntake.Decide와 같은 규칙).
/// </summary>
public static class SetupPlan
{
    /// <summary>호출부가 채우는 환경 상태(읽기 전용 프로브 결과).</summary>
    public sealed record Probes(
        bool ExcelOk, string? ExcelVersion,
        bool VbomOk, string VbomDetail,
        bool Installed, bool RunningInstalled, bool OnUserPath,
        bool OleaccOk = true);

    /// <param name="Workbook">검증된 원본 워크북 절대 경로(--workbook).</param>
    /// <param name="NewFolder">NewProjectFolder.Resolve가 확정한 새 프로젝트 폴더(--new).</param>
    /// <param name="Folder">--folder 원문(--workbook과 함께면 사본 위치).</param>
    public sealed record Request(
        string? Workbook, string? NewFolder, string? Folder,
        bool InstallOnly, bool NoInstall, bool SkipVerify);

    /// <param name="Start">workbook | new | folder</param>
    /// <param name="Action">copy | reuse | inplace | create | adopt</param>
    /// <param name="Verify">push+test | pull | skipped</param>
    public sealed record Project(string Start, string Folder, string Workbook, string Action, string? Source, string Verify);

    /// <summary>
    /// 사용자 동의가 필요한 변경. <paramref name="ApproveFlag"/>는 동의했을 때 명령에 덧붙일 플래그
    /// (null이면 명령에 이미 포함된 기본 동작), <paramref name="DeclineFlag"/>는 거절했을 때 덧붙일 플래그
    /// (null이면 거절 = setup을 진행할 수 없음).
    /// </summary>
    public sealed record Consent(string Id, string Reason, string? ApproveFlag, string? DeclineFlag);

    public sealed record Result(
        bool Ready, List<string> Blockers, string? Hint, bool EnvironmentBlocker,
        Project? Project, List<string> Changes, List<Consent> Consents, string Command, List<string> Args);

    public static Result Build(Request r, Probes p)
    {
        var blockers = new List<string>();
        string? hint = null;
        bool environmentBlocker = false;

        if (!p.ExcelOk)
        {
            blockers.Add("Desktop Excel is not installed (Excel.Application ProgID missing)");
            hint = "Install desktop Excel and rerun (Microsoft 365 web/mobile do not support COM automation)";
            environmentBlocker = true;
        }
        if (!p.OleaccOk)
        {
            // setup 1단계가 같은 이유로 거부한다 — 미리보기도 같은 판정을 내야 동의를 헛되이 받지 않는다
            blockers.Add("Cannot load oleacc.dll — Windows installation damaged");
            hint ??= "Run vbaloop doctor for details";
            environmentBlocker = true;
        }

        Project? project = null;
        if (!r.InstallOnly)
        {
            try
            {
                project = DescribeProject(r);
                if (project == null)
                {
                    blockers.Add("No starting point specified");
                    hint ??= "Existing workbook: --workbook <path.xlsm> / new project: --new <name> / prepared folder: --folder <path>";
                }
            }
            catch (CliError e)
            {
                blockers.Add(e.Message);
                hint ??= e.Hint;
            }
        }

        var consents = new List<Consent>();
        var changes = new List<string>();

        if (!r.NoInstall && (!p.RunningInstalled || !p.OnUserPath))
        {
            consents.Add(new Consent("install",
                $"Copy vbaloop.exe to {SelfInstall.InstallDir} and add it to the user PATH (vbaloop then runs from any shell)",
                null, "--no-install"));
            if (!p.RunningInstalled) changes.Add($"Copy exe: {SelfInstall.InstalledExePath}");
            if (!p.OnUserPath) changes.Add($"Add to user PATH: {SelfInstall.InstallDir}");
        }
        if (!p.VbomOk)
        {
            consents.Add(new Consent("accessVbom",
                "Turn on the Excel security setting 'Trust access to the VBA project object model' (HKCU Excel\\Security\\AccessVBOM=1). "
                + "vbaloop does not work while it is off. To revert: Excel Options > Trust Center > Macro Settings",
                "--enable-vbom", null));
            changes.Add("Turn on AccessVBOM (only when --enable-vbom is given)");
        }

        if (project != null)
        {
            string wbPath = Path.Combine(project.Folder, project.Workbook);
            changes.Add(project.Action switch
            {
                "copy" => $"Create a copy of the original: {wbPath} (original left untouched)",
                "reuse" => $"Reuse previously created dev folder: {project.Folder}",
                "inplace" => $"Original is already inside a project; use in place: {project.Folder}",
                "create" => $"Create folder and empty macro workbook: {wbPath} (starts Excel briefly)",
                _ => $"Adopt existing workbook in folder: {wbPath}",
            });
            changes.Add("Install managed files: vbaloop.toml, CLAUDE.md, AGENTS.md, .claude\\skills\\vbaloop\\, src\\LogShim.bas, .gitignore block"
                        + (project.Action == "create" ? ", sample test module" : ""));
            changes.Add(project.Verify switch
            {
                "pull" => "Verify: start a tool-owned Excel and pull workbook code into src\\ (the user's Excel is not touched)",
                "push+test" => "Verify: start a tool-owned Excel, push the sample test and run test to confirm green",
                _ => "Verification skipped (--skip-verify)",
            });
        }

        return new Result(blockers.Count == 0, blockers, hint, environmentBlocker,
            project, changes, consents, BuildCommand(r), BuildArgs(r));
    }

    /// <summary>시작점별 프로젝트 결정(파일시스템 읽기만). 시작점이 없으면 null.</summary>
    public static Project? DescribeProject(Request r)
    {
        if (r.Workbook != null)
        {
            var plan = WorkbookIntake.Decide(r.Workbook, r.Folder);
            ProjectRoot.EnsureNotNested(plan.Folder);
            return new Project("workbook", plan.Folder, plan.WorkbookName,
                plan.Kind.ToString().ToLowerInvariant(), plan.Source, r.SkipVerify ? "skipped" : "pull");
        }
        if (r.NewFolder != null)
        {
            ProjectRoot.EnsureNotNested(r.NewFolder);
            return new Project("new", r.NewFolder, new DirectoryInfo(r.NewFolder).Name + ".xlsm",
                "create", null, r.SkipVerify ? "skipped" : "push+test");
        }
        if (r.Folder != null)
        {
            string folder = Path.GetFullPath(r.Folder);
            ProjectRoot.EnsureNotNested(folder);
            // 이미 프로젝트면 toml이 가리키는 워크북을 채택한다(setup이 같은 규칙을 따른다)
            string? configured = ProjectRoot.ConfiguredWorkbook(folder);
            if (configured != null)
                return new Project("folder", folder, configured, "adopt", null, r.SkipVerify ? "skipped" : "pull");
            var existing = Directory.Exists(folder)
                ? Directory.EnumerateFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName).Where(n => n != null && !n.StartsWith("~$")).Select(n => n!)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList()
                : new List<string>();
            if (existing.Count > 1)
                throw new CliError($"Multiple workbook candidates: {string.Join(", ", existing)}",
                    "Specify the workbook to import: vbaloop setup --workbook <path.xlsm> (a copy is made in a new folder next to the original)");
            return existing.Count == 0
                ? new Project("folder", folder, new DirectoryInfo(folder).Name + ".xlsm", "create", null,
                    r.SkipVerify ? "skipped" : "push+test")
                : new Project("folder", folder, existing[0], "adopt", null, r.SkipVerify ? "skipped" : "pull");
        }
        return null;
    }

    /// <summary>
    /// 동의를 받은 뒤 실행할 비대화형 명령의 인자(exe 뒤, "setup"부터). --new는 cwd 기준이라 실행 위치가
    /// 바뀌면 다른 곳에 생기므로 확정된 절대 경로를 --folder로 넘긴다(빈 폴더 흐름과 동일한 결과). 동의
    /// 플래그(--enable-vbom)는 넣지 않는다 — 사용자가 동의했을 때만 에이전트가 덧붙인다.
    /// </summary>
    public static List<string> BuildArgs(Request r)
    {
        var args = new List<string> { "setup" };
        if (r.InstallOnly)
        {
            args.Add("--install-only");
        }
        else if (r.Workbook != null)
        {
            args.Add("--workbook");
            args.Add(r.Workbook);
            if (r.Folder != null)
            {
                args.Add("--folder");
                args.Add(Path.GetFullPath(r.Folder));
            }
        }
        else if (r.NewFolder != null || r.Folder != null)
        {
            args.Add("--folder");
            args.Add(Path.GetFullPath(r.NewFolder ?? r.Folder!));
        }
        if (r.NoInstall) args.Add("--no-install");
        if (r.SkipVerify) args.Add("--skip-verify");
        args.Add("--yes");
        args.Add("--json");
        return args;
    }

    /// <summary>BuildArgs를 PowerShell에서 그대로 실행할 수 있게 인용한 한 줄(경로는 작은따옴표 — 식 해석 없음).</summary>
    public static string BuildCommand(Request r) => ShellQuote.PowerShellCommand("vbaloop", BuildArgs(r));
}
