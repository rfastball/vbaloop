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
                multi.Fail($"워크북은 하나만 가져올 수 있음 (받은 파일 {paths}개)",
                    "워크북 파일 하나만 끌어다 놓거나 vbaloop setup --workbook <경로.xlsm>");
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

        // 프로젝트를 읽는 명령의 부수 신호(설계 §12.3): 관리 스탬프가 이 exe보다 새로우면 경고만 한다.
        // 쓰기(init)는 자체적으로 거부하고, push/test 등은 막지 않는다 - 원인은 exe라 루프를 세울 이유가 없다.
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
                "test" => TestCommand.Run(opts),
                "lease" => LeaseCommand.Run(opts),
                "lint" => LintCommand.Run(opts),
                "run" => RunCommand.Run(opts),
                "dist" => DistCommand.Run(opts),
                "rollback" => RollbackCommand.Run(opts),
                "bugreport" => BugReportCommand.Run(opts),
                "version" or "--version" => VersionCommand.Run(opts),
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
            outcome.Fail($"내부 오류: {e.GetType().Name}: {e.Message}");
        }

        // opts.Json이 authoritative지만, --json이 값으로 소비됐거나 verb 위치에 온 경로까지
        // 균일하게 봉투를 유지하도록 rawJson을 폴백으로 OR (xlFlow a.json || rawArgsRequestJSON)
        // 출력이 RCW 정리보다 먼저다: 격리된(VBA 실행 중) Excel의 프록시 Release는 무기한 블록할 수 있다.
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
        if (ManagedStamp.Decide(ExeVersion.Current, cfg.ManagedBy) != StampState.ExeOlder) return;
        var w = $"{ManagedStamp.DescribeExeOlder(cfg.ManagedBy!, ExeVersion.Text)} - {ManagedStamp.UpdateHint} 권장 (init은 거부됨)";
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
            outcome.Fail($"내부 오류: {e.GetType().Name}: {e.Message}");
        }
        Output.Print(outcome, json: false);
        ConsoleHost.PauseIfOwnConsole();
        return outcome.ExitCode;
    }

    private static Outcome UnknownVerb(string verb)
    {
        var o = new Outcome { Command = verb };
        o.Fail($"미지 명령: {verb}", "vbaloop help로 명령 목록 확인");
        return o;
    }

    private static Outcome UsageOutcome()
    {
        var o = new Outcome { Command = "help" };
        o.HumanLines.Add(UsageText());
        o.Payload = new { commands = new[] { "setup [--workbook <경로.xlsm> | --new <이름> | --folder <경로>] [--plan] [--yes] [--enable-vbom] [--skip-verify] [--install-only]", "init [--check] [--all]", "doctor [--kill-orphans]", "status", "diff [--module <이름>]", "session start|attach|stop [--discard]", "recovery status|clear [--force]", "pull [--force]", "push [--force] [--prune]", "test", "run <Module.Proc>", "dist [--out <경로>] [--strip-comments] [--bootstrap <Module.Proc>] [--skip-compile]", "rollback --list | --latest | --backup <id>", "lease acquire|release|status|break", "lint", "bugreport <제목> | --list", "version" } };
        return o;
    }

    private static void PrintUsage() => Console.Error.WriteLine(UsageText());

    private static string UsageText() => """
        vbaloop - Excel VBA 프로젝트를 위한 최소 에이전트 개발 루프

        처음이라면: vbaloop setup   (환경 점검 -> 설치 -> 워크북/폴더 선택 -> 하네스 -> 검증)
        기존 워크북: vbaloop.exe 위로 .xlsm을 끌어다 놓기 (원본 옆 <이름>\ 폴더에 사본을 만들어 작업)
        새 프로젝트: vbaloop setup --new <이름>   (.\<이름>\ 에 워크북 + 샘플 테스트)

        사용법:
          vbaloop setup   [--workbook <경로.xlsm> | --new <이름> | --folder <경로>] [--plan] [--yes] [--enable-vbom] [--skip-verify] [--install-only]  개발환경 마법사(대화형, --workbook은 기존 워크북 사본으로, --new는 .\<이름>\ 새 프로젝트로 시작, --plan은 변경 없이 계획만, --enable-vbom은 AccessVBOM 변경 동의(--yes에 미포함), --install-only는 exe 설치만)
          vbaloop init    [--check] [--all] [--json]  온보딩 부트스트랩/드리프트 수렴 (Excel 불필요, --all은 전 프로젝트 일괄)
          vbaloop doctor  [--json] [--kill-orphans] [--fix]  환경 점검 + 고아 Excel 리포트(kill/fix는 명시 시)
          vbaloop status  [--json]                 세션/프로젝트 상태
          vbaloop diff    [--json] [--module <이름>]  워크북 VBProject vs src 읽기 전용 비교 (활성 세션 필요)
          vbaloop session start|stop [--no-save] [--discard]   도구 소유 Excel 세션 시작/종료 (--discard: 저장 없이 버리고 격리 해제)
          vbaloop session attach                   사용자가 이미 연 정본 워크북을 채택 (stop은 분리만, Excel 종료/저장 안 함)
          vbaloop recovery status|clear [--force]  격리 상태 확인/해제 (clear는 격리된 Excel 종료 확인 후에만, --force는 경고와 함께)
          vbaloop pull    [--json] [--force]       워크북 -> src (동기화 안 된 src 편집 시 차단)
          vbaloop push    [--json] [--force] [--prune] [--force-form-import]   src -> 워크북 (in-place 주입, --force는 지문 캐시 우회 전량 재주입)
          vbaloop test    [--json] [--timeout N]   Test_* 프로시저 실행/판독
          vbaloop run     <Module.Proc> [--timeout N]   임의 매크로 단건 실행
          vbaloop dist    [--out <경로>] [--strip-comments] [--bootstrap <Module.Proc>] [--skip-compile]   클린 사본 배포 빌드(§9, 게시 전 VBE 컴파일 검증, .xlam도 지원)
          vbaloop rollback --list | --latest | --backup <id> [--json]   push 전 백업으로 워크북 복원 (세션 종료 상태에서만)
          vbaloop lease   acquire|release|status|break [--force]   다중 에이전트 lease
          vbaloop lint    [--json]                 규약/인코딩/헤더 전수 검사 (Excel 불필요)
          vbaloop bugreport "<제목>" | --list       vbaloop 자체 결함 리포트 생성/열거 (vbaloop\docs\bugreports)
          vbaloop version                          실행 exe의 git 리비전 + 빌드 시각

        공통: --config <vbaloop.toml 경로>  (기본: cwd에서 상위로 탐색)
        종료 코드: 성공 0, 환경 오류(세션/Excel/COM/busy/격리 - hint 확인) 3, 그 외 2
        """;
}
