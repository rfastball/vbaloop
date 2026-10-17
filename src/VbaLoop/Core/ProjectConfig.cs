namespace VbaLoop.Core;

/// <summary>
/// vbaloop.toml 로드 결과(설계 §5). 탐색 규칙: cwd에서 상위로 올라가며 최초로 만나는
/// vbaloop.toml의 위치가 프로젝트 루트(git rev-parse 방식). --config로 오버라이드 가능.
/// </summary>
public sealed class ProjectConfig
{
    public required string RootDir { get; init; }
    /// <summary>실제로 읽은 설정 파일(vbaloop.toml, 또는 이전 전이면 vbaflow.toml). 경로를 조립하지 말고 이걸 쓴다.</summary>
    public string TomlPath { get; init; } = "";
    public required string ProjectName { get; init; }
    /// <summary>워크북 절대 경로.</summary>
    public required string WorkbookPath { get; init; }
    public bool Visible { get; init; } = true;
    /// <summary>src 루트 절대 경로(평면 구조).</summary>
    public required string SrcRoot { get; init; }

    // ---- [dist] 배포 파이프라인(설계 §9, 전부 선택) ----
    /// <summary>배포 버전(semver). git rev/빌드시각은 dist가 자동 부가.</summary>
    public string? DistVersion { get; init; }
    /// <summary>배포 제외 모듈 글롭(보강). 1차 소스는 모듈 헤더 ' @dist exclude.</summary>
    public IReadOnlyList<string> DistExclude { get; init; } = Array.Empty<string>();
    /// <summary>주입 후 run할 워크시트 초기설정 프로시저(Module.Proc). --bootstrap로 오버라이드.</summary>
    public string? DistBootstrap { get; init; }
    /// <summary>opt-in 주석 제거(토크나이저 기반). --strip-comments로 오버라이드.</summary>
    public bool DistStripComments { get; init; }

    // ---- [backup] push 전 백업(설계 §4 push 0.5단계, 선택) ----
    /// <summary>보존할 최신 백업 개수. 0이면 push 전 백업 비활성(기존 백업은 건드리지 않음).</summary>
    public int BackupKeep { get; init; } = DefaultBackupKeep;
    public const int DefaultBackupKeep = 10;

    public string DbgDir => Path.Combine(RootDir, "_dbg");
    public string SessionMetaPath => Path.Combine(DbgDir, "session.json");
    public string ToolLogPath => Path.Combine(DbgDir, "vbaloop.jsonl");
    /// <summary>격리 마커(설계 §11) — 존재하는 동안 세션 필요 명령은 exit 3.</summary>
    public string RecoveryMarkerPath => Path.Combine(DbgDir, "recovery.json");
    public string BackupsDir => Path.Combine(DbgDir, "backups");

    private static readonly HashSet<string> KnownSections = new(StringComparer.OrdinalIgnoreCase)
        { "project", "excel", "src", "dist", "backup" };
    private static readonly Dictionary<string, HashSet<string>> KnownKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        ["project"] = new(StringComparer.OrdinalIgnoreCase) { "name" },
        ["excel"] = new(StringComparer.OrdinalIgnoreCase) { "workbook", "visible" },
        ["src"] = new(StringComparer.OrdinalIgnoreCase) { "root" },
        ["dist"] = new(StringComparer.OrdinalIgnoreCase) { "version", "exclude", "bootstrap", "stripcomments" },
        ["backup"] = new(StringComparer.OrdinalIgnoreCase) { "keep" },
    };

    public static string? FindTomlPath(string? configOverride, string cwd)
    {
        if (configOverride != null)
        {
            var full = Path.GetFullPath(configOverride);
            return File.Exists(full) ? full : null;
        }
        var dir = new DirectoryInfo(Path.GetFullPath(cwd));
        while (dir != null)
        {
            var candidate = Path.Combine(dir.FullName, "vbaloop.toml");
            if (File.Exists(candidate)) return candidate;
            // 개명 전 프로젝트(vbaflow.toml)도 그대로 동작 — vbaloop init이 새 이름으로 이전한다
            var legacy = Path.Combine(dir.FullName, LegacyNames.TomlFile);
            if (File.Exists(legacy)) return legacy;
            dir = dir.Parent;
        }
        return null;
    }

    /// <summary>개명 전 설정 파일(vbaflow.toml)인지 — init이 이전 대상으로 다룬다.</summary>
    public static bool IsLegacyToml(string tomlPath)
        => string.Equals(Path.GetFileName(tomlPath), LegacyNames.TomlFile, StringComparison.OrdinalIgnoreCase);

    public static ProjectConfig Load(string? configOverride, string cwd)
    {
        var tomlPath = FindTomlPath(configOverride, cwd)
            ?? throw new CliError(
                configOverride != null
                    ? $"--config 경로에 파일이 없음: {configOverride}"
                    : $"vbaloop.toml을 찾지 못함 (탐색 시작: {cwd})",
                "프로젝트 루트에서 vbaloop init 실행 (온보딩 부트스트랩)");

        var root = Path.GetDirectoryName(tomlPath)!;
        var doc = TomlMini.Parse(File.ReadAllText(tomlPath), tomlPath);

        foreach (var section in doc.Keys)
        {
            if (!KnownSections.Contains(section))
                throw new CliError($"{tomlPath}: 미지 섹션 [{section}]",
                    "지원 섹션: [project], [excel], [src], [dist], [backup]");
            foreach (var key in doc[section].Keys)
                if (!KnownKeys[section].Contains(key))
                    throw new CliError($"{tomlPath}: 미지 키 [{section}].{key}",
                        $"지원 키: {string.Join(", ", KnownKeys[section])}");
        }

        string? workbookRel = GetString(doc, "excel", "workbook");
        if (string.IsNullOrWhiteSpace(workbookRel))
            throw new CliError($"{tomlPath}: [excel].workbook이 없음",
                "workbook = \"파일명.xlsm\" (프로젝트 루트 기준 상대경로)");

        var workbookPath = Path.GetFullPath(Path.Combine(root, workbookRel));
        var srcRel = GetString(doc, "src", "root") ?? "src";
        var name = GetString(doc, "project", "name")
                   ?? Path.GetFileNameWithoutExtension(workbookPath);
        bool visible = GetBool(doc, "excel", "visible") ?? true;
        int backupKeep = GetBackupKeep(doc, tomlPath);

        return new ProjectConfig
        {
            RootDir = root,
            TomlPath = tomlPath,
            ProjectName = name,
            WorkbookPath = workbookPath,
            Visible = visible,
            SrcRoot = Path.GetFullPath(Path.Combine(root, srcRel)),
            DistVersion = GetString(doc, "dist", "version"),
            DistExclude = GetStringList(doc, "dist", "exclude"),
            DistBootstrap = GetString(doc, "dist", "bootstrap"),
            DistStripComments = GetBool(doc, "dist", "stripComments") ?? false,
            BackupKeep = backupKeep,
        };
    }

    /// <summary>[backup] keep: 0 이상 정수만. 타입 오류는 조용히 기본값으로 떨어지지 않고 실패한다 —
    /// keep = "0"(문자열)을 무시하고 백업을 계속 쌓으면 사용자의 비활성 의도를 어기기 때문.</summary>
    private static int GetBackupKeep(Dictionary<string, Dictionary<string, object>> doc, string tomlPath)
    {
        if (!doc.TryGetValue("backup", out var s) || !s.TryGetValue("keep", out var v))
            return DefaultBackupKeep;
        if (v is not int keep)
            throw new CliError($"{tomlPath}: [backup].keep은 정수여야 함", "keep = 10 (0이면 push 전 백업 비활성)");
        if (keep < 0)
            throw new CliError($"{tomlPath}: [backup].keep은 0 이상이어야 함 (현재 {keep})", "keep = 10 (0이면 push 전 백업 비활성)");
        return keep;
    }

    private static string? GetString(Dictionary<string, Dictionary<string, object>> doc, string section, string key)
        => doc.TryGetValue(section, out var s) && s.TryGetValue(key, out var v) ? v as string : null;

    private static bool? GetBool(Dictionary<string, Dictionary<string, object>> doc, string section, string key)
        => doc.TryGetValue(section, out var s) && s.TryGetValue(key, out var v) && v is bool b ? b : null;

    private static IReadOnlyList<string> GetStringList(Dictionary<string, Dictionary<string, object>> doc, string section, string key)
        => doc.TryGetValue(section, out var s) && s.TryGetValue(key, out var v) && v is List<string> l
            ? l : Array.Empty<string>();
}
