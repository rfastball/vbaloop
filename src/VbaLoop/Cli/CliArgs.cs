using VbaLoop.Core;

namespace VbaLoop.Cli;

public sealed class CliArgs
{
    public List<string> Positional { get; } = new();
    public bool Json { get; private set; }
    public bool Prune { get; private set; }
    public bool NoSave { get; private set; }
    public bool Check { get; private set; }
    public bool Force { get; private set; }
    public bool ForceFormImport { get; private set; }
    public bool KillOrphans { get; private set; }
    /// <summary>session stop 전용: 저장 없이 세션을 버리고 격리를 해제(설계 §11).</summary>
    public bool Discard { get; private set; }
    /// <summary>init 전용: 온보딩된 전 프로젝트에 일괄 수렴 (vbaloop 레포 아래 exe에서만).</summary>
    public bool All { get; private set; }
    /// <summary>bugreport 전용: 기존 리포트 열거 (open/resolved).</summary>
    public bool List { get; private set; }
    public int? TimeoutSec { get; private set; }
    public string? Config { get; private set; }
    public string? RunId { get; private set; }
    /// <summary>__runner 전용: 부모가 확보한 세션의 메타 파일(session.json 또는 일회용 인스턴스 메타).</summary>
    public string? SessionMeta { get; private set; }
    /// <summary>__runner 전용: 테스트 열거 대신 단건 매크로 실행 (Module.Proc).</summary>
    public string? Macro { get; private set; }
    /// <summary>diff 전용: 단일 모듈로 좁히고 워크북 쪽 본문을 덤프.</summary>
    public string? Module { get; private set; }
    public string? Spec { get; private set; }
    public string? Code { get; private set; }
    /// <summary>dist 전용: 배포 아티팩트 출력 경로(선택, 생략 시 dist/&lt;워크북명&gt;.xlsm).
    /// 확장자로 형식 선언 — .xlsm(매크로 워크북) / .xlam(애드인).</summary>
    public string? Out { get; private set; }
    /// <summary>dist 전용: 주입 후 실행할 부트스트랩 프로시저(Module.Proc), toml 오버라이드.</summary>
    public string? Bootstrap { get; private set; }
    /// <summary>dist 전용: 주석 제거 opt-in, toml stripComments 오버라이드.</summary>
    public bool StripComments { get; private set; }
    /// <summary>setup 전용: 대상 프로젝트 폴더(생략 시 폴더 선택 대화상자).</summary>
    public string? Folder { get; private set; }
    /// <summary>setup 전용: 가져올 기존 워크북(.xlsm). 원본 옆 &lt;이름&gt;\ 폴더에 사본을 만들어
    /// 프로젝트로 쓴다(--folder를 주면 그 폴더). exe에 워크북을 드래그 앤 드롭한 경우도 이것으로 온다.</summary>
    public string? Workbook { get; private set; }
    /// <summary>setup 전용: 새 빈 프로젝트 이름. cwd 아래 &lt;이름&gt;\ 폴더에 &lt;이름&gt;.xlsm과 샘플 테스트를 만든다.</summary>
    public string? NewName { get; private set; }
    /// <summary>setup 전용: AccessVBOM(보안 설정)을 켜는 데 사용자가 동의했음. --yes는 이것을 포함하지 않는다 —
    /// 에이전트가 프롬프트를 넘기려고 붙인 --yes로 보안 설정이 바뀌면 안 되기 때문(설계 §10.8).</summary>
    public bool EnableVbom { get; private set; }
    /// <summary>setup 전용: 아무것도 바꾸지 않고 할 일·필요한 동의·실행할 명령만 보고한다.</summary>
    public bool Plan { get; private set; }
    /// <summary>setup/doctor 전용: 확인 프롬프트를 묻지 않고 승인으로 진행(비대화형 설치).</summary>
    public bool Yes { get; private set; }
    /// <summary>setup 전용: 마지막 push/test(또는 pull) 검증 단계를 건너뛴다 — Excel 불필요.</summary>
    public bool SkipVerify { get; private set; }
    /// <summary>setup 전용: exe 복사/PATH 등록을 건너뛴다(개발 빌드에서 마법사만 시험할 때).</summary>
    public bool NoInstall { get; private set; }
    /// <summary>setup 전용: 설치(exe 복사 + PATH)까지만 하고 프로젝트는 건드리지 않는다 —
    /// exe만 갱신할 때(rollout, 업데이트) 쓴다.</summary>
    public bool InstallOnly { get; private set; }
    /// <summary>doctor 전용: 고칠 수 있는 환경 항목(AccessVBOM)을 기록한다 — 명시가 곧 승인.</summary>
    public bool Fix { get; private set; }
    /// <summary>rollback 전용: 최신 완결 백업으로 복원.</summary>
    public bool Latest { get; private set; }
    /// <summary>rollback 전용: 지정 id 백업으로 복원(rollback --list의 id).</summary>
    public string? Backup { get; private set; }
    /// <summary>dist 전용: 게시 전 VBE 컴파일 검증을 건너뛴다(탈출구).</summary>
    public bool SkipCompile { get; private set; }

    public static CliArgs Parse(string[] args)
    {
        var r = new CliArgs();
        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--json": r.Json = true; break;
                case "--prune": r.Prune = true; break;
                case "--no-save" or "-NoSave": r.NoSave = true; break;
                case "--check": r.Check = true; break;
                case "--force": r.Force = true; break;
                case "--force-form-import": r.ForceFormImport = true; break;
                case "--kill-orphans": r.KillOrphans = true; break;
                case "--discard": r.Discard = true; break;
                case "--all": r.All = true; break;
                case "--list": r.List = true; break;
                case "--timeout":
                    if (++i >= args.Length || !int.TryParse(args[i], out var t) || t <= 0)
                        throw new CliError("--timeout requires a positive integer (seconds)");
                    r.TimeoutSec = t;
                    break;
                case "--config":
                    if (++i >= args.Length) throw new CliError("--config requires a path");
                    r.Config = args[i];
                    break;
                case "--runid":
                    if (++i >= args.Length) throw new CliError("--runid requires a value");
                    r.RunId = args[i];
                    break;
                case "--session-meta":
                    if (++i >= args.Length) throw new CliError("--session-meta requires a path");
                    r.SessionMeta = args[i];
                    break;
                case "--macro":
                    if (++i >= args.Length) throw new CliError("--macro requires Module.Proc");
                    r.Macro = args[i];
                    break;
                case "--module":
                    if (++i >= args.Length) throw new CliError("--module requires a module name");
                    r.Module = args[i];
                    break;
                case "--spec":
                    if (++i >= args.Length || args[i].StartsWith('-')) throw new CliError("--spec requires a JSON path");
                    r.Spec = args[i];
                    break;
                case "--code":
                    if (++i >= args.Length || args[i].StartsWith('-')) throw new CliError("--code requires a VBA path");
                    r.Code = args[i];
                    break;
                case "--out":
                    if (++i >= args.Length) throw new CliError("--out requires an output path");
                    r.Out = args[i];
                    break;
                case "--bootstrap":
                    if (++i >= args.Length) throw new CliError("--bootstrap requires Module.Proc");
                    r.Bootstrap = args[i];
                    break;
                case "--strip-comments":
                    r.StripComments = true;
                    break;
                case "--yes" or "-y": r.Yes = true; break;
                case "--skip-verify": r.SkipVerify = true; break;
                case "--enable-vbom": r.EnableVbom = true; break;
                case "--plan": r.Plan = true; break;
                case "--no-install": r.NoInstall = true; break;
                case "--install-only": r.InstallOnly = true; break;
                case "--fix": r.Fix = true; break;
                case "--latest": r.Latest = true; break;
                case "--skip-compile": r.SkipCompile = true; break;
                case "--backup":
                    if (++i >= args.Length) throw new CliError("--backup requires a backup id");
                    r.Backup = args[i];
                    break;
                case "--folder":
                    if (++i >= args.Length) throw new CliError("--folder requires a path");
                    r.Folder = args[i];
                    break;
                case "--workbook":
                    if (++i >= args.Length) throw new CliError("--workbook requires an .xlsm path");
                    r.Workbook = args[i];
                    break;
                case "--new":
                    if (++i >= args.Length) throw new CliError("--new requires a project name");
                    r.NewName = args[i];
                    break;
                default:
                    if (args[i].StartsWith('-'))
                        throw new CliError($"unknown option: {args[i]}",
                            "Supported options: --json --prune --no-save --check --force --force-form-import --kill-orphans --discard --all --list --yes --skip-verify --enable-vbom --plan --no-install --install-only --fix --latest --skip-compile --backup <id> --timeout <seconds> --config <path> --module <name> --out <path> --bootstrap <Module.Proc> --folder <path> --workbook <path> --new <name> --spec <design.json> --code <body.vba> --strip-comments");
                    r.Positional.Add(args[i]);
                    break;
            }
        }
        return r;
    }

    /// <summary>
    /// 파싱이 실패해 Parse가 --json을 채우지 못하는 경로(Program의 CliError catch)를 위해
    /// raw argv에서 --json을 직접 감지한다(xlFlow rawArgsRequestJSON 이식). 값-옵션
    /// (--timeout/--config/--runid/--macro/--module) 다음 토큰은 그 옵션의 값이므로 건너뛴다 —
    /// 값이 우연히 "--json"이어도 JSON 요청으로 오인하지 않는다.
    /// </summary>
    public static bool RawArgsRequestJson(string[] argv)
    {
        for (int i = 0; i < argv.Length; i++)
        {
            if (argv[i] == "--json") return true;
            if (ValueOptions.Contains(argv[i])) i++; // 다음 토큰은 이 옵션의 값
        }
        return false;
    }

    /// <summary>값을 하나 받는 옵션 — 다음 토큰은 옵션 값이지 위치 인자가 아니다.</summary>
    public static readonly HashSet<string> ValueOptions = new(StringComparer.Ordinal)
    {
        "--timeout", "--config", "--runid", "--session-meta", "--macro", "--module",
        "--out", "--bootstrap", "--folder", "--workbook", "--new", "--backup", "--spec", "--code",
    };

    /// <summary>
    /// 옵션도 옵션 값도 아닌 위치 인자들. 드래그 앤 드롭 축약형(<c>vbaloop Book.xlsm --folder D:\x</c>)에서
    /// 워크북 경로 수를 셀 때 --folder·--config 같은 값의 경로를 워크북으로 오인하지 않게 한다.
    /// </summary>
    public static List<string> Positionals(string[] argv)
    {
        var result = new List<string>();
        for (int i = 0; i < argv.Length; i++)
        {
            if (ValueOptions.Contains(argv[i])) { i++; continue; }
            if (!argv[i].StartsWith('-')) result.Add(argv[i]);
        }
        return result;
    }
}
