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
            blockers.Add("Excel 데스크톱이 설치돼 있지 않음 (Excel.Application ProgID 없음)");
            hint = "Excel 데스크톱 버전을 설치한 뒤 다시 실행 (Microsoft 365 웹/모바일은 COM 자동화 불가)";
            environmentBlocker = true;
        }
        if (!p.OleaccOk)
        {
            // setup 1단계가 같은 이유로 거부한다 — 미리보기도 같은 판정을 내야 동의를 헛되이 받지 않는다
            blockers.Add("oleacc.dll을 로드할 수 없음 — Windows 구성 손상");
            hint ??= "vbaloop doctor로 상세 점검";
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
                    blockers.Add("시작점이 지정되지 않음");
                    hint ??= "기존 워크북: --workbook <경로.xlsm> / 새 프로젝트: --new <이름> / 준비된 폴더: --folder <경로>";
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
                $"vbaloop.exe를 {SelfInstall.InstallDir}에 복사하고 사용자 PATH에 등록 (이후 어느 셸에서나 vbaloop 실행)",
                null, "--no-install"));
            if (!p.RunningInstalled) changes.Add($"exe 복사: {SelfInstall.InstalledExePath}");
            if (!p.OnUserPath) changes.Add($"사용자 PATH에 추가: {SelfInstall.InstallDir}");
        }
        if (!p.VbomOk)
        {
            consents.Add(new Consent("accessVbom",
                "Excel 보안 설정 'VBA 프로젝트 개체 모델에 대한 신뢰할 수 있는 액세스'를 켬 (HKCU Excel\\Security\\AccessVBOM=1). "
                + "꺼져 있으면 vbaloop가 동작하지 않음. 되돌리기: Excel 옵션 > 보안 센터 > 매크로 설정",
                "--enable-vbom", null));
            changes.Add("AccessVBOM 켜기 (--enable-vbom을 붙였을 때만)");
        }

        if (project != null)
        {
            string wbPath = Path.Combine(project.Folder, project.Workbook);
            changes.Add(project.Action switch
            {
                "copy" => $"원본 사본 생성: {wbPath} (원본은 그대로)",
                "reuse" => $"이전에 만든 개발 폴더 재사용: {project.Folder}",
                "inplace" => $"원본이 이미 프로젝트 안에 있어 그 자리에서 사용: {project.Folder}",
                "create" => $"폴더와 빈 매크로 워크북 생성: {wbPath} (Excel을 잠시 띄움)",
                _ => $"폴더의 기존 워크북 채택: {wbPath}",
            });
            changes.Add("관리 파일 설치: vbaloop.toml, CLAUDE.md, AGENTS.md, .claude\\skills\\vbaloop\\, src\\LogShim.bas, .gitignore 블록"
                        + (project.Action == "create" ? ", 샘플 테스트 모듈" : ""));
            changes.Add(project.Verify switch
            {
                "pull" => "검증: 도구 소유 Excel을 띄워 워크북 코드를 src\\로 꺼냄 (사용자가 연 Excel은 건드리지 않음)",
                "push+test" => "검증: 도구 소유 Excel을 띄워 샘플 테스트 push + test로 초록불 확인",
                _ => "검증 생략 (--skip-verify)",
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
                throw new CliError($"워크북 후보가 복수임: {string.Join(", ", existing)}",
                    "가져올 워크북을 지정: vbaloop setup --workbook <경로.xlsm> (원본 옆 새 폴더에 사본을 만든다)");
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
