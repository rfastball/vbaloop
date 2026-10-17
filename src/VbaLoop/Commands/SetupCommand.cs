using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop setup — 사람이 생애 1회 쓰는 개발환경 마법사. 배포본을 받아 실행하면 환경 점검 →
/// AccessVBOM → 자기설치(PATH) → 워크북/폴더 선택 → 워크북 → 하네스 → 검증까지 한 번에 세운다.
/// 기존 워크북은 exe에 끌어다 놓거나 --workbook으로 넘기면 원본 옆 사본 폴더에서 시작하고,
/// 대화형이면 끝에서 그 폴더에 Claude Code를 띄워 준다.
///
/// 이 명령만 대화형이다. 에이전트가 부르는 나머지 CLI 표면은 전부 비대화형 --json이며,
/// 여기서도 --folder/--yes를 주면 프롬프트 없이 끝까지 진행한다(스크립트 설치 경로).
///
/// 마지막 단계가 push/test(신규) 또는 pull(기존)인 것이 설계의 핵심 —
/// 마법사가 끝날 때 실제 초록불을 보여줘야 "환경이 섰다"가 증명되고, 이후 에이전트가 만나는
/// 실패는 최소한 환경 문제가 아니라는 보장이 선다.
/// </summary>
public static class SetupCommand
{
    private const string RepoUrl = "https://github.com/rfastball/vbaloop";

    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "setup" };
        void Say(string line = "") { if (!opts.Json) Console.WriteLine(line); }

        // --install-only는 프로젝트를 건드리지 않고 설치까지만 — 단계 표시도 거기서 끝난다
        int total = opts.InstallOnly ? 3 : 6;

        // 시작점은 하나만: 기존 워크북(--workbook) / 새 프로젝트(--new) / 준비된 폴더(--folder).
        // --workbook + --folder만 예외로 허용한다(사본을 둘 폴더 지정). 환경 점검 전에 거른다.
        if (opts.NewName != null && (opts.Workbook != null || opts.Folder != null))
        {
            o.Fail("--new는 --workbook·--folder와 함께 쓸 수 없음",
                "새 프로젝트: vbaloop setup --new <이름> (cwd 아래 생성) / 기존 워크북: vbaloop setup --workbook <경로.xlsm>");
            return o;
        }
        // 잘못 끌어다 놓은 파일·쓸 수 없는 이름은 설치·AccessVBOM 변경 전에 거른다(CliError → 봉투)
        string? givenWorkbook = opts.Workbook != null ? WorkbookIntake.Validate(opts.Workbook) : null;
        string? givenNewFolder = opts.NewName != null
            ? NewProjectFolder.Resolve(Environment.CurrentDirectory, opts.NewName)
            : null;
        // 다른 프로젝트 안쪽 폴더는 init·검증이 바깥 프로젝트로 새어 나간다 — 역시 변경 전에 거른다.
        // --plan은 같은 판정을 blocker로 보고하므로 여기서 던지지 않는다.
        if (!opts.Plan)
        {
            if (givenNewFolder != null) ProjectRoot.EnsureNotNested(givenNewFolder);
            if (givenWorkbook != null) ProjectRoot.EnsureNotNested(WorkbookIntake.Decide(givenWorkbook, opts.Folder).Folder);
            else if (opts.Folder != null) ProjectRoot.EnsureNotNested(opts.Folder);
        }

        if (opts.Plan)
            return PlanOnly(opts, givenWorkbook, givenNewFolder, o);

        Say(opts.InstallOnly ? "vbaloop setup --install-only — exe 설치/갱신" : "vbaloop setup — 개발환경 마법사");
        Say();

        // ── [1] 환경 점검 ───────────────────────────────────────────────
        Say($"[1/{total}] 환경 점검");
        string? curVer = null;
        var progType = Type.GetTypeFromProgID("Excel.Application");
        if (progType == null)
        {
            Say("  FAIL Excel: Excel.Application ProgID 없음");
            o.Fail("Excel이 설치돼 있지 않음 (Excel.Application ProgID 없음)",
                "Excel 데스크톱 버전을 설치한 뒤 다시 실행 (Microsoft 365 웹/모바일은 COM 자동화 불가)",
                environment: true);
            return o;
        }
        using (var curVerKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(@"Excel.Application\CurVer"))
            curVer = curVerKey?.GetValue(null) as string;
        Say($"  ok   Excel: {curVer ?? "버전 미상"}");

        bool oleacc = CanLoadOleacc();
        if (!oleacc)
        {
            Say("  FAIL oleacc.dll 로드 실패");
            o.Fail("oleacc.dll을 로드할 수 없음 — Windows 구성 손상", "vbaloop doctor로 상세 점검",
                environment: true);
            return o;
        }
        Say("  ok   oleacc.dll");

        // ── [2/6] AccessVBOM — 꺼져 있으면 아무것도 동작하지 않는 진짜 첫 관문 ──
        Say();
        Say($"[2/{total}] VBA 프로젝트 개체 모델 접근(AccessVBOM)");
        var (vbomOk, vbomDetail, _) = AccessVbom.Detect();
        bool vbomFixed = false;
        if (vbomOk)
        {
            Say($"  ok   {vbomDetail}");
        }
        else
        {
            Say($"  FAIL {vbomDetail}");
            Say("       이 설정이 꺼져 있으면 vbaloop의 push/pull/test가 전부 실패합니다.");
            // 보안 설정이라 동의는 두 가지로만 받는다: 사람이 프롬프트에 답하거나, 사용자 동의를 받은
            // 에이전트가 --enable-vbom을 명시하거나. --yes(프롬프트 생략)는 이 동의를 대신하지 않는다.
            bool canAsk = !opts.Yes && !opts.Json && ConsoleHost.CanPrompt;
            bool agree = opts.EnableVbom
                         || (canAsk && ConsoleHost.Confirm("       지금 켤까요?", defaultYes: true, fallback: false));
            if (!agree)
            {
                o.Fail("AccessVBOM이 비활성 상태 — 설정 없이는 진행할 수 없음",
                    canAsk
                        ? "Excel 옵션 > 보안 센터 > 매크로 설정 > 'VBA 프로젝트 개체 모델에 안전하게 액세스' 체크 후 재실행"
                        : "보안 설정 변경이라 사용자 동의가 필요: 사용자에게 묻고 동의하면 같은 명령에 --enable-vbom을 붙여 재실행 "
                          + "(또는 사용자가 Excel 옵션 > 보안 센터 > 매크로 설정에서 직접 체크)",
                    environment: true);
                o.Payload = new { consentRequired = new[] { "accessVbom" }, approveFlag = "--enable-vbom", detail = vbomDetail };
                return o;
            }
            var (changed, detail) = AccessVbom.Enable(curVer);
            vbomFixed = changed;
            Say($"  ok   {detail}");
        }

        // ── [3/6] 설치: exe를 고정 폴더로 + PATH 등록 ────────────────────
        Say();
        Say($"[3/{total}] 설치");
        bool copied = false, pathAdded = false;
        if (opts.NoInstall)
        {
            Say("  skip --no-install (현재 exe를 그대로 사용)");
        }
        else
        {
            (copied, string installedPath) = SelfInstall.CopySelf();
            Say($"  ok   {installedPath}{(copied ? "" : " (이미 이 경로에서 실행 중)")}");
            pathAdded = SelfInstall.EnsureOnPath();
            Say(pathAdded
                ? $"  ok   PATH 등록: {SelfInstall.InstallDir}"
                : $"  ok   PATH 이미 등록됨: {SelfInstall.InstallDir}");
        }

        if (opts.InstallOnly)
        {
            Say();
            Say("설치 완료. 프로젝트에 하네스를 얹으려면: vbaloop setup");
            if (pathAdded)
                Say("주의: PATH는 새로 여는 셸부터 반영됩니다 — 지금 열려 있는 터미널은 재시작하세요.");
            o.Payload = new
            {
                installOnly = true,
                accessVbomFixed = vbomFixed,
                installed = !opts.NoInstall,
                installDir = opts.NoInstall ? null : SelfInstall.InstallDir,
                exeCopied = copied,
                pathAdded,
                repo = RepoUrl,
            };
            return o;
        }

        // ── [4/6] 프로젝트 폴더 ─────────────────────────────────────────
        // 기존 워크북을 받으면(드래그 앤 드롭/--workbook/파일 선택) 폴더는 도구가 정한다 —
        // 사용자가 빈 폴더를 만들어 워크북을 혼자 넣어 두는 준비 단계를 없앤다(원본은 불변).
        Say();
        Say("[4/6] 프로젝트 폴더");
        var (workbookArg, newFolder, chosenFolder) = givenNewFolder != null || givenWorkbook != null
            ? (givenWorkbook, givenNewFolder, (string?)null)
            : ChooseStart(opts, Say);
        string folder;
        string? intakeWorkbook = null;
        WorkbookIntake.Kind? intakeKind = null;
        if (newFolder != null)
        {
            folder = newFolder;
            ProjectRoot.EnsureNotNested(folder); // 대화형으로 고른 경우(현재 폴더·이름 입력)도 여기서 걸린다
            Say(Directory.Exists(folder)
                ? "  ok   빈 폴더를 새 프로젝트로 씁니다"
                : "  ok   새 프로젝트 폴더를 만들었습니다");
            Directory.CreateDirectory(folder);
        }
        else if (workbookArg != null)
        {
            string source = WorkbookIntake.Validate(workbookArg);
            var plan = WorkbookIntake.Decide(source, opts.Folder);
            ProjectRoot.EnsureNotNested(plan.Folder); // 사본을 만들기 전에
            WorkbookIntake.Execute(plan);
            folder = plan.Folder;
            intakeWorkbook = plan.WorkbookName;
            intakeKind = plan.Kind;
            Say($"  원본: {source}");
            Say(plan.Kind switch
            {
                WorkbookIntake.Kind.Copy => "  ok   사본을 만들었습니다 (원본은 그대로 둡니다)",
                WorkbookIntake.Kind.Reuse => "  ok   이전에 만든 개발 폴더를 재사용합니다 (원본의 이후 변경은 반영되지 않음)",
                _ => "  ok   이미 vbaloop 프로젝트 안의 워크북입니다 — 그 자리에서 사용합니다",
            });
        }
        else
        {
            string? picked = chosenFolder ?? ResolveFolder(opts, Say);
            if (picked == null)
            {
                o.Fail("프로젝트 폴더가 지정되지 않음",
                    "vbaloop setup --workbook <경로.xlsm> 또는 vbaloop setup --folder <경로>");
                return o;
            }
            folder = Path.GetFullPath(picked);
            ProjectRoot.EnsureNotNested(folder);
            Directory.CreateDirectory(folder);
        }
        Say($"  → {folder}");

        // ── [5/6] 워크북 + 하네스 ───────────────────────────────────────
        Say();
        Say("[5/6] 워크북 + 하네스");
        // 이미 프로젝트인 폴더(다시 세팅)는 toml이 정본이다 — 하위 폴더의 워크북이나 최상위 워크북 여러 개도
        // 설정된 것 하나로 확정한다. 스캔하면 엉뚱한 새 워크북을 만들거나 복수 후보로 실패한다.
        string? configured = intakeWorkbook == null ? ProjectRoot.ConfiguredWorkbook(folder) : null;
        var existing = intakeWorkbook != null
            ? new List<string> { intakeWorkbook } // Decide가 폴더 안의 대상을 이미 하나로 확정했다
            : configured != null
            ? new List<string> { configured }
            : Directory.EnumerateFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                .Where(n => n != null && !n.StartsWith("~$"))
                .Select(n => n!)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        if (existing.Count > 1)
        {
            o.Fail($"워크북 후보가 복수임: {string.Join(", ", existing)}",
                "가져올 워크북을 지정: vbaloop setup --workbook <경로.xlsm> (원본 옆 새 폴더에 사본을 만든다)");
            return o;
        }

        bool isNew = existing.Count == 0;
        string workbookName;
        if (isNew)
        {
            workbookName = new DirectoryInfo(folder).Name + ".xlsm";
            string wbPath = Path.Combine(folder, workbookName);
            Say($"  ..   빈 매크로 워크북 생성: {workbookName}");
            WorkbookFactory.CreateMacroWorkbook(wbPath);
            Say($"  ok   {workbookName} 생성");
        }
        else
        {
            workbookName = existing[0];
            Say($"  ok   기존 워크북 채택: {workbookName}");
        }

        // init/push/pull/test는 cwd 기준으로 프로젝트를 찾는다 — 대상 폴더로 옮겨 놓고 부른다
        string prevCwd = Environment.CurrentDirectory;
        try
        {
            Environment.CurrentDirectory = folder;

            var initOut = InitCommand.Run(CliArgs.Parse(Array.Empty<string>()));
            foreach (var line in initOut.HumanLines) Say($"  {line}");
            if (!initOut.Ok)
            {
                o.Fail("온보딩(init) 실패: " + string.Join("; ", initOut.Errors), initOut.Hint,
                    initOut.EnvironmentError); // exe-older 거부(exit 3, hint vbaloop update)를 보존
                return o;
            }

            // 신규 프로젝트에만 샘플을 넣는다 — 기존 워크북은 pull이 진실을 가져온다
            bool sampleWritten = false;
            if (isNew)
            {
                // src 위치는 toml이 권위다("src"는 스캐폴드 기본값일 뿐) — 하드코딩하면
                // [src].root를 바꾼 프로젝트에서 샘플이 엉뚱한 곳에 떨어진다
                string srcRoot = ProjectConfig.Load(null, folder).SrcRoot;
                string samplePath = Path.Combine(srcRoot, SampleProject.FileName);
                if (!File.Exists(samplePath))
                {
                    Directory.CreateDirectory(srcRoot);
                    CanonicalText.WriteCanonical(samplePath, SampleProject.ModuleLines());
                    sampleWritten = true;
                    Say($"  ok   {Path.GetRelativePath(folder, samplePath)} 샘플 생성");
                }
            }

            // ── [6/6] 검증 ──────────────────────────────────────────────
            Say();
            Say("[6/6] 검증");
            string verify;
            if (opts.SkipVerify)
            {
                verify = "skipped";
                Say("  skip --skip-verify");
            }
            else if (isNew)
            {
                var pushOut = PushCommand.Run(CliArgs.Parse(Array.Empty<string>()));
                foreach (var line in pushOut.HumanLines) Say($"  {line}");
                if (!pushOut.Ok)
                {
                    o.Fail("push 실패: " + string.Join("; ", pushOut.Errors), pushOut.Hint,
                        pushOut.EnvironmentError);
                    return o;
                }
                var testOut = TestCommand.Run(CliArgs.Parse(Array.Empty<string>()));
                foreach (var line in testOut.HumanLines) Say($"  {line}");
                if (!testOut.Ok)
                {
                    o.Fail("test 실패: " + string.Join("; ", testOut.Errors), testOut.Hint,
                        testOut.EnvironmentError);
                    return o;
                }
                verify = "push+test";
            }
            else
            {
                var pullOut = PullCommand.Run(CliArgs.Parse(Array.Empty<string>()));
                foreach (var line in pullOut.HumanLines) Say($"  {line}");
                if (!pullOut.Ok)
                {
                    o.Fail("pull 실패: " + string.Join("; ", pullOut.Errors), pullOut.Hint,
                        pullOut.EnvironmentError);
                    return o;
                }
                verify = "pull";
            }

            // ── 마무리 안내 ─────────────────────────────────────────────
            bool claudeCli = HasClaudeCli();
            Say();
            Say("완료. 다음 단계:");
            Say($"  cd {folder}");
            Say("  claude          ← 이 폴더에서 코드 에이전트를 실행하면 vbaloop 루프를 바로 사용합니다");
            if (!isNew)
                Say("                    (기존 워크북 코드는 src\\ 로 내려받았습니다)");
            if (pathAdded)
            {
                Say();
                Say("  주의: PATH는 새로 여는 셸부터 반영됩니다 — 지금 열려 있는 터미널은 재시작하세요.");
            }
            // 일회용 모드(기본)는 검증 명령이 끝나면 Excel도 닫힌다 — 세션이 남는 것은 persistent 모드뿐
            bool sessionOpen = verify != "skipped" && File.Exists(Path.Combine(folder, "_dbg", "session.json"));
            if (sessionOpen)
            {
                Say();
                Say("  참고: 검증을 위해 도구 소유 Excel 세션이 열려 있습니다(정상). 다음 명령이 재사용하며,");
                Say("        지금 닫으려면 vbaloop session stop.");
            }
            if (!claudeCli)
            {
                Say();
                Say("  참고: PATH에서 claude 실행 파일을 찾지 못했습니다. Claude Code가 없으면");
                Say("        .claude\\skills\\vbaloop\\SKILL.md는 사용되지 않습니다(AGENTS.md는 유효).");
            }

            // 마지막 관문(폴더 찾아가서 claude 실행)까지 마법사가 대신한다 — 대화형에서만 묻는다
            bool claudeLaunched = false;
            if (claudeCli && !opts.Yes && !opts.Json && ConsoleHost.CanPrompt)
            {
                Say();
                if (ConsoleHost.Confirm("지금 이 폴더에서 Claude Code를 새 창으로 열까요?", defaultYes: true, fallback: false))
                {
                    string? launchError = LaunchClaude(folder);
                    claudeLaunched = launchError == null;
                    Say(claudeLaunched
                        ? "  ok   Claude Code 창을 열었습니다. 예: \"Test_로 시작하는 테스트를 하나 만들고 돌려줘\""
                        : $"  FAIL Claude Code 실행 실패: {launchError} — 위 안내대로 직접 실행하세요");
                }
            }

            o.Payload = new
            {
                folder,
                workbook = workbookName,
                intake = intakeKind?.ToString().ToLowerInvariant(),
                createdWorkbook = isNew,
                sample = sampleWritten,
                accessVbomFixed = vbomFixed,
                installed = !opts.NoInstall,
                installDir = opts.NoInstall ? null : SelfInstall.InstallDir,
                exeCopied = copied,
                pathAdded,
                verify,
                claudeCli,
                claudeLaunched,
                repo = RepoUrl,
                next = BuildNext(folder, opts.NoInstall, sessionOpen),
            };
            return o;
        }
        finally
        {
            Environment.CurrentDirectory = prevCwd;
        }
    }

    /// <summary>
    /// 대화형 시작점 선택(--workbook·--new·--folder가 모두 없을 때만). 반환은 (가져올 워크북, 새 프로젝트
    /// 폴더, 기존 프로젝트 폴더) 중 하나, 모두 null이면 준비된 폴더 흐름(ResolveFolder)으로 간다.
    /// <list type="bullet">
    /// <item>현재 폴더 우선(셸 실행만): 프로젝트면 다시 세팅, .xlsm이 있으면 그것, 비었으면 여기에 새 프로젝트</item>
    /// <item>기존 워크북: 파일 선택 대화상자, 취소하면 콘솔 경로 입력(창에 파일을 끌어다 놓으면 입력됨)</item>
    /// <item>새 프로젝트: 이름만 받는다. 부모는 더블클릭이면 %USERPROFILE%\vbaloop-projects, 셸이면 cwd</item>
    /// </list>
    /// </summary>
    private static (string? Workbook, string? NewFolder, string? Folder) ChooseStart(CliArgs opts, Action<string> say)
    {
        if (opts.Folder != null || opts.Yes || opts.Json || !ConsoleHost.CanPrompt) return (null, null, null);

        // 셸에서 부른 경우(설치 스크립트가 이어서 부르는 경우 포함) 현재 폴더가 곧 의도다 — "워크북이 있는
        // 폴더에서 한 줄 실행"만으로 시작되게 먼저 제안한다. 더블클릭은 cwd가 exe 위치라 건너뛴다.
        if (!ConsoleHost.IsOwnConsole())
        {
            var fromCwd = OfferCurrentFolder(Environment.CurrentDirectory, say);
            if (fromCwd.Workbook != null || fromCwd.NewFolder != null || fromCwd.Folder != null) return fromCwd;
        }

        if (ConsoleHost.Confirm("  기존 워크북(.xlsm)을 가져올까요? (n = 새 빈 프로젝트)", defaultYes: true, fallback: false))
        {
            say("  파일 선택 대화상자를 엽니다 (원본은 그대로 두고 옆에 사본 폴더를 만듭니다)");
            string? picked = FilePicker.PickWorkbook("vbaloop로 가져올 워크북(.xlsm)을 고르세요")
                ?? ConsoleHost.Ask("  워크북 경로 (이 창에 파일을 끌어다 놓아도 됨, 빈 줄 = 폴더 직접 선택):");
            return (picked, null, null);
        }

        string parent = ConsoleHost.IsOwnConsole() ? NewProjectFolder.ExplorerDefaultParent : Environment.CurrentDirectory;
        say($"  새 프로젝트는 {parent}\\<이름> 에 만듭니다");
        while (true)
        {
            string? name = ConsoleHost.Ask("  프로젝트 이름 (빈 줄 = 준비된 폴더 직접 선택):");
            if (name == null) return (null, null, null);
            try
            {
                return (null, NewProjectFolder.Resolve(parent, name), null);
            }
            catch (CliError e)
            {
                say($"  FAIL {e.Message}" + (e.Hint != null ? $" — {e.Hint}" : ""));
            }
        }
    }

    /// <summary>현재 폴더 상태에 맞는 시작점을 하나 제안한다. 거절하거나 해당 없으면 전부 null.</summary>
    private static (string? Workbook, string? NewFolder, string? Folder) OfferCurrentFolder(string cwd, Action<string> say)
    {
        var here = CwdStart.Inspect(cwd);
        if (here.IsProject)
        {
            if (ConsoleHost.Confirm("  현재 폴더는 이미 vbaloop 프로젝트입니다. 다시 세팅할까요? (관리 파일 갱신 + 검증)",
                    defaultYes: true, fallback: false))
                return (null, null, cwd);
        }
        else if (here.Workbooks.Count == 1)
        {
            string name = here.Workbooks[0];
            if (ConsoleHost.Confirm(
                    $"  현재 폴더의 {name}을(를) 가져올까요? (옆에 {Path.GetFileNameWithoutExtension(name)}\\ 폴더를 만들어 사본으로 작업)",
                    defaultYes: true, fallback: false))
                return (Path.Combine(cwd, name), null, null);
        }
        else if (here.Workbooks.Count > 1)
        {
            say("  현재 폴더의 워크북:");
            for (int i = 0; i < here.Workbooks.Count; i++) say($"    {i + 1}. {here.Workbooks[i]}");
            while (true)
            {
                string? pick = ConsoleHost.Ask("  가져올 번호 (빈 줄 = 다른 방법):");
                if (pick == null) break;
                if (int.TryParse(pick, out int k) && k >= 1 && k <= here.Workbooks.Count)
                    return (Path.Combine(cwd, here.Workbooks[k - 1]), null, null);
                say("  FAIL 목록에 있는 번호를 입력하세요");
            }
        }
        else if (here.IsEmpty)
        {
            // 폴더 이름이 곧 워크북 이름이 된다 — 드라이브 루트처럼 이름으로 못 쓰는 곳은 제안하지 않는다
            string folderName = new DirectoryInfo(cwd).Name;
            bool usable;
            try { NewProjectFolder.ValidateName(folderName); usable = true; }
            catch (CliError) { usable = false; }
            if (usable && ConsoleHost.Confirm($"  현재 폴더({folderName})가 비어 있습니다. 여기에 새 프로젝트를 만들까요?",
                    defaultYes: true, fallback: false))
                return (null, cwd, null);
        }
        return default;
    }

    /// <summary>
    /// 대상 폴더 결정: --folder > 폴더 선택 대화상자 > 콘솔 입력. 전부 불가하면 null.
    /// 대화상자는 Shell COM이라 세션 0 등에서 실패할 수 있어 콘솔 입력 폴백을 둔다.
    /// </summary>
    private static string? ResolveFolder(CliArgs opts, Action<string> say)
    {
        if (opts.Folder != null) return opts.Folder;
        if (opts.Json || !ConsoleHost.CanPrompt) return null; // 비대화형에서 대화상자 금지

        say("  폴더 선택 대화상자를 엽니다 (취소하면 경로를 직접 입력할 수 있습니다)");
        string? picked = FolderPicker.Pick("vbaloop 개발환경을 얹을 프로젝트 폴더를 고르세요");
        return picked ?? ConsoleHost.Ask("  프로젝트 폴더 경로:");
    }

    /// <summary>
    /// setup --plan 본체: 환경을 읽기만 하고 SetupPlan으로 할 일·동의 항목·실행 명령을 보고한다.
    /// 막는 요인(blocker)이 있으면 ok=false지만 payload는 그대로 채운다 — 에이전트가 무엇이 문제인지와
    /// 나머지 계획을 한 번에 사용자에게 보여줄 수 있게.
    /// </summary>
    private static Outcome PlanOnly(CliArgs opts, string? workbook, string? newFolder, Outcome o)
    {
        string? curVer;
        using (var curVerKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(@"Excel.Application\CurVer"))
            curVer = curVerKey?.GetValue(null) as string;
        var (vbomOk, vbomDetail, _) = AccessVbom.Detect();
        var probes = new SetupPlan.Probes(
            ExcelOk: Type.GetTypeFromProgID("Excel.Application") != null,
            ExcelVersion: curVer,
            VbomOk: vbomOk,
            VbomDetail: vbomDetail,
            Installed: File.Exists(SelfInstall.InstalledExePath),
            RunningInstalled: Environment.ProcessPath != null
                              && PathDrift.SamePath(Environment.ProcessPath, SelfInstall.InstalledExePath),
            OnUserPath: SelfInstall.IsOnUserPath(),
            OleaccOk: CanLoadOleacc());
        var request = new SetupPlan.Request(workbook, newFolder, opts.Folder,
            opts.InstallOnly, opts.NoInstall, opts.SkipVerify);
        var plan = SetupPlan.Build(request, probes);

        if (!plan.Ready)
            foreach (var b in plan.Blockers)
                o.Fail(b, plan.Hint, plan.EnvironmentBlocker);

        o.Payload = new
        {
            plan = true,
            ready = plan.Ready,
            blockers = plan.Blockers,
            environment = new
            {
                excel = probes.ExcelOk,
                excelVersion = probes.ExcelVersion,
                accessVbom = probes.VbomOk,
                accessVbomDetail = probes.VbomDetail,
            },
            install = new
            {
                installDir = SelfInstall.InstallDir,
                exe = SelfInstall.InstalledExePath,
                installed = probes.Installed,
                onUserPath = probes.OnUserPath,
            },
            project = plan.Project,
            changes = plan.Changes,
            consent = plan.Consents,
            command = plan.Command,   // PowerShell용 한 줄(경로는 작은따옴표)
            args = plan.Args,         // 셸에 무관한 인자 배열("setup"부터) — exe 절대 경로 뒤에 그대로 붙인다
            agents = DetectAgents(),
        };

        if (!opts.Json)
        {
            o.HumanLines.Add("vbaloop setup --plan (아무것도 바꾸지 않음)");
            o.HumanLines.Add("");
            o.HumanLines.Add("할 일:");
            foreach (var c in plan.Changes) o.HumanLines.Add($"  - {c}");
            if (plan.Consents.Count > 0)
            {
                o.HumanLines.Add("");
                o.HumanLines.Add("동의가 필요한 항목:");
                foreach (var c in plan.Consents)
                    o.HumanLines.Add($"  - [{c.Id}] {c.Reason}"
                        + (c.ApproveFlag != null ? $" (동의 시 {c.ApproveFlag})" : "")
                        + (c.DeclineFlag != null ? $" (거절 시 {c.DeclineFlag})" : ""));
            }
            o.HumanLines.Add("");
            o.HumanLines.Add($"실행할 명령: {plan.Command}");
        }
        return o;
    }

    /// <summary>
    /// 세팅을 마친 에이전트가 이어서 할 일(설계 §10.8). setup을 부른 에이전트 세션은 보통 다른 폴더에서
    /// 돌고 있어 새 프로젝트의 CLAUDE.md/AGENTS.md·스킬이 로드돼 있지 않다 — 그 세션에서 이어갈지,
    /// 프로젝트 폴더에서 새 세션을 열지 사용자가 고르게 둘 다 준다. 에이전트 셸의 PATH는 설치 전
    /// 상태로 굳어 있으므로 exe는 절대 경로로 준다.
    /// </summary>
    private static object BuildNext(string folder, bool noInstall, bool sessionOpen)
    {
        string toml = Path.Combine(folder, "vbaloop.toml");
        return new
        {
            folder,
            exe = noInstall ? Environment.ProcessPath : SelfInstall.InstalledExePath,
            agentDocs = new[] { Path.Combine(folder, "CLAUDE.md"), Path.Combine(folder, "AGENTS.md") },
            continueHere = $"이 세션에서 이어가려면 {Path.Combine(folder, "CLAUDE.md")}(Codex는 AGENTS.md)의 vbaloop 블록과 "
                           + $"{Path.Combine(folder, ".claude", "skills", "vbaloop", "SKILL.md")}를 읽고, "
                           + $"이후 vbaloop 명령은 그 폴더에서 실행(또는 --config {ShellQuote.PowerShell(toml)})",
            // 셸마다 문법이 달라(cmd의 cd /d·&&는 PowerShell 5.1에서 실패) 작업 폴더와 명령을 따로 주고,
            // 바로 실행할 수 있는 PowerShell 한 줄을 곁들인다
            newSession = DetectAgents().Select(a => new
            {
                agent = a,
                workingDirectory = folder,
                command = a,
                powershell = $"Set-Location -LiteralPath {ShellQuote.PowerShell(folder)}; {a}",
            }).ToArray(),
            excelSession = sessionOpen
                ? "검증용 도구 소유 Excel 세션이 열려 있음(정상) - 다음 명령이 재사용, 끝낼 때 vbaloop session stop"
                : null,
        };
    }

    /// <summary>attach가 쓰는 oleacc.dll을 로드할 수 있는가 — setup 1단계와 --plan이 같은 판정을 쓴다.</summary>
    private static bool CanLoadOleacc()
    {
        bool ok = System.Runtime.InteropServices.NativeLibrary.TryLoad("oleacc.dll", out var h);
        if (ok) System.Runtime.InteropServices.NativeLibrary.Free(h);
        return ok;
    }

    /// <summary>PATH에서 찾은 코드 에이전트 CLI — 세팅 후 새 세션 안내(next)에 쓴다.</summary>
    private static string[] DetectAgents()
        => new[] { "claude", "codex" }.Where(OnPath).ToArray();

    /// <summary>PATH에서 claude 실행 파일을 찾는다 — 안내 문구 분기·실행 제안용(실패해도 setup은 성공).</summary>
    private static bool HasClaudeCli() => OnPath("claude");

    private static bool OnPath(string baseName)
    {
        string[] names = { baseName + ".exe", baseName + ".cmd", baseName + ".bat" };
        foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'))
        {
            if (dir.Length == 0) continue;
            try
            {
                foreach (var n in names)
                    if (File.Exists(Path.Combine(dir, n))) return true;
            }
            catch { /* 잘못된 PATH 항목 무시 */ }
        }
        return false;
    }

    /// <summary>
    /// 프로젝트 폴더에서 새 콘솔 창으로 Claude Code를 띄운다. <c>cmd /k</c>를 거치는 이유:
    /// npm 설치본은 claude.cmd라 직접 실행할 수 없고, claude가 끝나도 창이 남아 이어서 쓸 수 있다.
    /// 셸 실행(ShellExecute)은 이 프로세스의 환경을 물려주므로, 방금 PATH에 넣은 설치 폴더도
    /// 새 창에서 바로 보인다(SelfInstall.EnsureOnPath가 현재 프로세스 PATH도 갱신).
    /// 마법사가 claude를 자식으로 붙잡고 기다리지 않는다 — Excel RCW를 쥔 채 세션 내내 살아 있게 된다.
    /// </summary>
    private static string? LaunchClaude(string folder)
    {
        try
        {
            // PATH가 레지스트리엔 있어도 이 프로세스가 그 전에 시작됐으면 빠져 있다 — 새 창에서
            // 에이전트가 vbaloop를 못 찾지 않게 설치본이 있으면 현재 프로세스 PATH에 보장한다
            if (File.Exists(SelfInstall.InstalledExePath))
            {
                var (live, changed) = SelfInstall.AddToPathValue(
                    Environment.GetEnvironmentVariable("Path"), SelfInstall.InstallDir);
                if (changed) Environment.SetEnvironmentVariable("Path", live);
            }
            using var p = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/k claude",
                WorkingDirectory = folder,
                UseShellExecute = true,
            });
            return null;
        }
        catch (Exception e)
        {
            return e.Message;
        }
    }
}
