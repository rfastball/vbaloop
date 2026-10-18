using System.Text;
using VbaLoop.Cli;
using VbaLoop.Commands;
using VbaLoop.Core;

namespace VbaLoop;

internal static class Program
{
    // STA: OleMessageFilter 등록과 Excel COM 마샬링의 전제 (설계 §2 런타임 근거 2)
    [STAThread]
    private static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        try { Console.OutputEncoding = Encoding.UTF8; } catch { /* 리다이렉트 환경 */ }

        if (args.Length == 0)
        {
            // 탐색기에서 더블클릭한 경우(콘솔을 이 프로세스가 단독 소유)는 배포본을 처음 받은
            // 사람이다 — usage 대신 마법사로 맞이한다. 셸에서 인자 없이 부른 경우는 usage.
            if (ConsoleHost.IsOwnConsole())
                return RunSetupFromDoubleClick();
            PrintUsage();
            return 0;
        }

        // 탐색기에서 워크북을 exe 위로 끌어다 놓으면 인자가 그 파일 경로로만 온다 — 명령(verb)은
        // 경로 문자를 포함하지 않으므로 setup --workbook으로 해석한다. 셸에서 `vbaloop Book.xlsm`도 같다.
        if (WorkbookIntake.LooksLikePath(args[0]))
        {
            int paths = CliArgs.Positionals(args).Count(WorkbookIntake.LooksLikePath);
            if (paths > 1)
            {
                var multi = new Outcome { Command = "setup" };
                multi.Fail($"only one workbook can be imported ({paths} files received)",
                    "Drop a single workbook file, or run vbaloop setup --workbook <path.xlsm>");
                Output.Print(multi, json: false);
                ConsoleHost.PauseIfOwnConsole();
                return multi.ExitCode;
            }
            if (ConsoleHost.IsOwnConsole() && args.Length == 1)
                return RunSetupFromDoubleClick(args[0]);
            args = new[] { "setup", "--workbook" }.Concat(args).ToArray(); // 뒤따르는 옵션(--json 등) 유지
        }

        string verb = args[0];
        // 파싱 실패 시 opts.Json을 못 읽으므로 raw argv에서 --json을 미리 감지(xlFlow rawArgsRequestJSON)
        bool rawJson = CliArgs.RawArgsRequestJson(args);
        CliArgs opts;
        try
        {
            opts = CliArgs.Parse(args.Skip(1).ToArray());
        }
        catch (CliError e)
        {
            var parseFail = new Outcome { Command = verb };
            parseFail.Fail(e.Message, e.Hint);
            Output.Print(parseFail, rawJson);
            return parseFail.ExitCode;
        }

        // 자식 러너 verb: 봉투 대신 자체 JSON 프로토콜로 부모와 통신 (설계 §4 test)
        if (verb == "__runner")
            return RunnerMain.Run(opts);
        if (verb == "__probe")
            return Excel.SessionProbe.RunChild(opts);
        if (verb == UpdateCheck.ChildVerb)
            return UpdateCheck.RunChild();

        // vbaloop update가 남긴 이전 exe(vbaloop.exe.old) 정리 — 설치 폴더에서 돌 때만(설계 §12.5)
        string? exeDir = Environment.ProcessPath != null ? Path.GetDirectoryName(Environment.ProcessPath) : null;
        if (exeDir != null && PathDrift.SamePath(exeDir, SelfInstall.InstallDir))
            SelfReplace.CleanupOld(exeDir);

        // 프로젝트를 읽는 명령의 부수 신호: 사용자 프로젝트 레지스트리에 루트 기록(설계 §12.6, init --all 대상),
        // 관리 스탬프가 이 exe보다 새로우면 경고(§12.3). 쓰기(init)는 자체적으로 거부하고, push/test 등은
        // 막지 않는다 - 원인은 exe라 루프를 세울 이유가 없다.
        var warnings = new List<string>();
        ProjectConfig.Loaded = cfg => OnProjectLoaded(cfg, warnings);

        var outcome = new Outcome { Command = verb };
        try
        {
            outcome = verb switch
            {
                "setup" => SetupCommand.Run(opts),
                "init" => InitCommand.Run(opts),
                "doctor" => DoctorCommand.Run(opts),
                "status" => StatusCommand.Run(opts),
                "diff" => DiffCommand.Run(opts),
                "session" => SessionCommand.Run(opts),
                "recovery" => RecoveryCommand.Run(opts),
                "pull" => PullCommand.Run(opts),
                "push" => PushCommand.Run(opts),
                "form" => FormCommand.Run(opts),
                "test" => TestCommand.Run(opts),
                "lease" => LeaseCommand.Run(opts),
                "lint" => LintCommand.Run(opts),
                "run" => RunCommand.Run(opts),
                "dist" => DistCommand.Run(opts),
                "rollback" => RollbackCommand.Run(opts),
                "bugreport" => BugReportCommand.Run(opts),
                "version" or "--version" => VersionCommand.Run(opts),
                "update" => UpdateCommand.Run(opts),
                "help" or "--help" or "-h" => UsageOutcome(),
                _ => UnknownVerb(verb),
            };
        }
        catch (CliError e)
        {
            outcome.Fail(e.Message, e.Hint, e.Environment);
        }
        catch (Exception e)
        {
            outcome.Fail($"internal error: {e.GetType().Name}: {e.Message}");
        }

        // opts.Json이 authoritative지만, --json이 값으로 소비됐거나 verb 위치에 온 경로까지
        // 균일하게 봉투를 유지하도록 rawJson을 폴백으로 OR (xlFlow a.json || rawArgsRequestJSON)
        // 출력이 RCW 정리보다 먼저다: 격리된(VBA 실행 중) Excel의 프록시 Release는 무기한 블록할 수 있다.
        // 새 릴리스 알림(설계 §12.4): 사람이 상태를 보는 명령에서만, 캐시만 읽는다(stale이면 분리된 자식이 갱신)
        if (verb is "status" or "doctor" or "version" or "--version" || (verb == "init" && opts.Check))
            outcome.Update = UpdateCheck.Poll(repoBuild: exeDir != null && RepoLocator.FindRepoRoot(exeDir) != null);

        Output.Print(outcome, opts.Json || rawJson);
        foreach (var w in warnings)
            Console.Error.WriteLine($"vbaloop {verb}: warning: {w}");

        // 격리·busy로 응답 없는 Excel의 RCW를 쥐고 있으면 그 프록시 Release가 무기한 블록한다. 주 STA
        // 스레드는 어떤 대기에서든 COM을 펌프해 그 호출을 받으므로 시간 제한도 무력하다(런타임 실측
        // 2026-10-17: 출력 후 미종료). 정리 없이 즉시 종료 — Excel 처분은 recovery 경로 소관.
        // exit 3 전체로 넓히면 read-only 거부처럼 Quit한 Excel이 RCW 미해제로 남는다(같은 날 실측).
        if (ComCleanup.SkipOnExit)
            ComCleanup.HardExit(outcome.ExitCode);

        // RCW 정리 — 참조 잔류로 인한 좀비 EXCEL 방지 (설계 §2 런타임 근거 1)
        if (!ComCleanup.Run(TimeSpan.FromSeconds(5)))
            ComCleanup.HardExit(outcome.ExitCode);
        return outcome.ExitCode;
    }

    private static void OnProjectLoaded(ProjectConfig cfg, List<string> warnings)
    {
        ProjectRegistry.TryAdd(ProjectRegistry.DefaultPath, cfg.RootDir);
        if (ManagedStamp.Decide(ExeVersion.Current, cfg.ManagedBy) != StampState.ExeOlder) return;
        var w = $"{ManagedStamp.DescribeExeOlder(cfg.ManagedBy!, ExeVersion.Text)} - {ManagedStamp.UpdateHint}; init is refused";
        if (!warnings.Contains(w)) warnings.Add(w);
    }

    /// <summary>
    /// 탐색기 더블클릭 진입: 마법사를 돌리고, 콘솔 창이 즉시 닫혀 결과를 못 읽는 일이 없도록
    /// 끝에서 입력을 기다린다. 예외도 여기서 사람이 읽을 수 있게 출력한다.
    /// <paramref name="droppedWorkbook"/>은 exe 위로 드래그 앤 드롭된 워크북 경로.
    /// </summary>
    private static int RunSetupFromDoubleClick(string? droppedWorkbook = null)
    {
        var outcome = new Outcome { Command = "setup" };
        try
        {
            outcome = SetupCommand.Run(CliArgs.Parse(droppedWorkbook == null
                ? Array.Empty<string>()
                : new[] { "--workbook", droppedWorkbook }));
        }
        catch (CliError e)
        {
            outcome.Fail(e.Message, e.Hint, e.Environment);
        }
        catch (Exception e)
        {
            outcome.Fail($"internal error: {e.GetType().Name}: {e.Message}");
        }
        Output.Print(outcome, json: false);
        ConsoleHost.PauseIfOwnConsole();
        return outcome.ExitCode;
    }

    private static Outcome UnknownVerb(string verb)
    {
        var o = new Outcome { Command = verb };
        o.Fail($"unknown command: {verb}", "See the command list with vbaloop help");
        return o;
    }

    private static Outcome UsageOutcome()
    {
        var o = new Outcome { Command = "help" };
        o.HumanLines.Add(UsageText());
        o.Payload = new { commands = new[] { "setup [--workbook <path.xlsm> | --new <name> | --folder <path>] [--plan] [--yes] [--enable-vbom] [--skip-verify] [--install-only]", "init [--check] [--all]", "doctor [--kill-orphans]", "status", "diff [--module <name>]", "session start|attach|stop [--discard]", "recovery status|clear [--force]", "pull [--force]", "push [--force] [--prune] [--skip-compile]", "form build --spec <design.json> --code <body.vba> [--force]", "test", "run <Module.Proc>", "dist [--out <path>] [--strip-comments] [--bootstrap <Module.Proc>] [--skip-compile]", "rollback --list | --latest | --backup <id>", "lease acquire|release|status|break", "lint", "bugreport <title> | --list", "version", "update [--check] [--yes] [--force]" } };
        return o;
    }

    private static void PrintUsage() => Console.Error.WriteLine(UsageText());

    private static string UsageText() => """
        vbaloop - minimal agent development loop for Excel VBA projects

        First time: vbaloop setup   (environment check -> install -> pick workbook/folder -> harness -> verify)
        Existing workbook: drag the .xlsm onto vbaloop.exe (works on a copy in a <name>\ folder next to the original)
        New project: vbaloop setup --new <name>   (workbook + sample test in .\<name>\)

        Usage:
          vbaloop setup   [--workbook <path.xlsm> | --new <name> | --folder <path>] [--plan] [--yes] [--enable-vbom] [--skip-verify] [--install-only]  dev-environment wizard (interactive; --workbook starts from a copy of an existing workbook, --new starts a new project in .\<name>\, --plan shows the plan without changes, --enable-vbom consents to the AccessVBOM change (not included in --yes), --install-only installs the exe only)
          vbaloop init    [--check] [--all] [--json]  onboarding bootstrap / drift convergence (no Excel needed; --all covers every registered project)
          vbaloop doctor  [--json] [--kill-orphans] [--fix]  environment check + orphan Excel report (kill/fix only when specified)
          vbaloop status  [--json]                 session/project status
          vbaloop diff    [--json] [--module <name>]  read-only comparison of workbook VBProject vs src (ephemeral mode uses a read-only Excel; persistent needs an active session)
          vbaloop session start|stop [--no-save] [--discard]   start/stop a persistent tool-owned Excel session (optional in the default ephemeral mode; --discard: discard without saving and clear quarantine)
          vbaloop session attach                   adopt the canonical workbook the user already opened (stop only detaches; Excel is not closed or saved)
          vbaloop recovery status|clear [--force]  check/clear quarantine (clear only after confirming the quarantined Excel is closed; --force with a warning)
          vbaloop pull    [--json] [--force]       workbook -> src (blocked if src has unsynced edits)
          vbaloop push    [--json] [--force] [--prune] [--force-form-import] [--skip-compile]   src -> workbook (in-place injection -> save only if VBE compile passes; --force bypasses the fingerprint cache and re-injects everything)
          vbaloop test    [--json] [--timeout N]   run/read Test_* procedures
          vbaloop form build --spec <design.json> --code <body.vba> [--force] [--json]   generate src frm/frx in private Excel; then push
          vbaloop run     <Module.Proc> [--timeout N]   run a single arbitrary macro
          vbaloop dist    [--out <path>] [--strip-comments] [--bootstrap <Module.Proc>] [--skip-compile]   clean-copy distribution build (§9, VBE compile check before publishing, .xlam also supported)
          vbaloop rollback --list | --latest | --backup <id> [--json]   restore the workbook from a pre-push backup (only while no persistent session is running)
          vbaloop lease   acquire|release|status|break [--force]   multi-agent lease
          vbaloop lint    [--json]                 full check of conventions/encoding/headers (no Excel needed)
          vbaloop bugreport "<title>" | --list     create/list vbaloop's own defect reports (vbaloop\docs\bugreports)
          vbaloop version                          running exe's version + git revision + build time
          vbaloop update  [--check] [--yes] [--force]   update the installed copy to the latest release (SHA-256 verified), then ask whether to run init --all (--yes: run without asking)

        Common: --config <vbaloop.toml path>  (default: search upward from cwd)
        Session mode: vbaloop.toml [excel] session = "ephemeral" (default, hidden Excel per command) | "persistent"
        Exit codes: success 0, environment error (session/Excel/COM/busy/quarantine - see hint) 3, otherwise 2
        """;
}
