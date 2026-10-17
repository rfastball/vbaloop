namespace VbaLoop.Core;

/// <summary>
/// vbaloop.toml 로드 결과(설계 §5). 탐색 규칙: cwd에서 상위로 올라가며 최초로 만나는
/// vbaloop.toml의 위치가 프로젝트 루트(git rev-parse 방식). --config로 오버라이드 가능.
/// </summary>
public sealed class ProjectConfig
{
    public required string RootDir { get; init; }
    /// <summary>실제로 읽은 설정 파일(--config로 다른 이름을 줄 수 있다). 경로를 조립하지 말고 이걸 쓴다.</summary>
    public string TomlPath { get; init; } = "";
    public required string ProjectName { get; init; }
    /// <summary>워크북 절대 경로.</summary>
    public required string WorkbookPath { get; init; }
    /// <summary>[excel].visible — persistent 세션 Excel의 표시 여부. 일회용 인스턴스는 항상 숨긴다.</summary>
    public bool Visible { get; init; } = true;
    /// <summary>[excel].session — 세션 필요 명령이 쓰는 Excel 수명(설계 §4 세션 모드). 기본은 일회용.</summary>
    public SessionMode SessionMode { get; init; } = SessionMode.Ephemeral;
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

    // ---- [push] (설계 §4 push, 선택) ----
    /// <summary>저장 전 VBE 컴파일 게이트. 일회용 인스턴스에서만 동작한다. --skip-compile로 1회 생략.</summary>
    public bool PushCompile { get; init; } = true;

    public string DbgDir => Path.Combine(RootDir, "_dbg");
    public string SessionMetaPath => Path.Combine(DbgDir, "session.json");
    public string ToolLogPath => Path.Combine(DbgDir, "vbaloop.jsonl");
    /// <summary>격리 마커(설계 §11) — 존재하는 동안 세션 필요 명령은 exit 3.</summary>
    public string RecoveryMarkerPath => Path.Combine(DbgDir, "recovery.json");
    public string BackupsDir => Path.Combine(DbgDir, "backups");
    /// <summary>일회용 인스턴스 메타(명령 하나 수명). 남아 있으면 비정상 종료한 명령의 잔재.</summary>
    public string EphemeralDir => Path.Combine(DbgDir, "ephemeral");

    // ---- [vbaloop] 관리 스탬프(설계 §12, init이 기록) ----
    /// <summary>관리 파일을 마지막으로 쓴 vbaloop 버전. 스탬프 도입 이전 프로젝트는 null.</summary>
    public string? ManagedBy { get; init; }

    /// <summary>
    /// Load 성공 직후 호출되는 훅. Program이 프로젝트 레지스트리 기록·exe-older 경고에 쓴다 —
    /// 순수 로더에 사용자 단위 부수효과를 넣지 않기 위해 주입한다(테스트는 null).
    /// </summary>
    public static Action<ProjectConfig>? Loaded { get; set; }

    private static readonly HashSet<string> KnownSections = new(StringComparer.OrdinalIgnoreCase)
        { "project", "excel", "src", "dist", "backup", "push", ManagedStamp.Section };
    private static readonly Dictionary<string, HashSet<string>> KnownKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        ["project"] = new(StringComparer.OrdinalIgnoreCase) { "name" },
        ["excel"] = new(StringComparer.OrdinalIgnoreCase) { "workbook", "visible", "session" },
        ["src"] = new(StringComparer.OrdinalIgnoreCase) { "root" },
        ["dist"] = new(StringComparer.OrdinalIgnoreCase) { "version", "exclude", "bootstrap", "stripcomments" },
        ["backup"] = new(StringComparer.OrdinalIgnoreCase) { "keep" },
        ["push"] = new(StringComparer.OrdinalIgnoreCase) { "compile" },
        // [vbaloop]은 도구 소유 섹션 — 미지 키를 거부하지 않는다(미래 exe가 키를 늘려도 이 exe가
        // 스탬프 비교까지는 도달해 "exe가 오래됨"을 말할 수 있게).
        [ManagedStamp.Section] = new(StringComparer.OrdinalIgnoreCase) { ManagedStamp.Key },
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
            dir = dir.Parent;
        }
        return null;
    }

    public static ProjectConfig Load(string? configOverride, string cwd)
    {
        var tomlPath = FindTomlPath(configOverride, cwd)
            ?? throw new CliError(
                configOverride != null
                    ? $"no file at the --config path: {configOverride}"
                    : $"vbaloop.toml not found (search started at: {cwd})",
                "run vbaloop init at the project root (onboarding bootstrap)");

        var root = Path.GetDirectoryName(tomlPath)!;
        var text = File.ReadAllText(tomlPath);
        Dictionary<string, Dictionary<string, object>> doc;
        try
        {
            doc = TomlMini.Parse(text, tomlPath);
            foreach (var section in doc.Keys)
            {
                if (!KnownSections.Contains(section))
                    throw new CliError($"{tomlPath}: unknown section [{section}]",
                        "supported sections: [project], [excel], [src], [dist], [backup], [push], [vbaloop]");
                if (string.Equals(section, ManagedStamp.Section, StringComparison.OrdinalIgnoreCase))
                    continue;
                foreach (var key in doc[section].Keys)
                    if (!KnownKeys[section].Contains(key))
                        throw new CliError($"{tomlPath}: unknown key [{section}].{key}",
                            $"supported keys: {string.Join(", ", KnownKeys[section])}");
            }
        }
        catch (CliError e) when (ExeOlderThan(text) is string stamp)
        {
            // 더 새 exe가 관리하는 프로젝트의 문법/키를 이 exe가 모를 수 있다 — 원인은 toml이 아니라 exe다
            throw new CliError($"{ManagedStamp.DescribeExeOlder(stamp, ExeVersion.Text)} ({e.Message})",
                ManagedStamp.UpdateHint, environment: true);
        }

        string? workbookRel = GetString(doc, "excel", "workbook");
        if (string.IsNullOrWhiteSpace(workbookRel))
            throw new CliError($"{tomlPath}: [excel].workbook is missing",
                "workbook = \"filename.xlsm\" (path relative to the project root)");

        var workbookPath = Path.GetFullPath(Path.Combine(root, workbookRel));
        var srcRel = GetString(doc, "src", "root") ?? "src";
        var name = GetString(doc, "project", "name")
                   ?? Path.GetFileNameWithoutExtension(workbookPath);
        bool visible = GetBool(doc, "excel", "visible") ?? true;
        int backupKeep = GetBackupKeep(doc, tomlPath);
        var sessionMode = GetSessionMode(doc, tomlPath);
        bool pushCompile = GetStrictBool(doc, "push", "compile", tomlPath) ?? true;

        var cfg = new ProjectConfig
        {
            RootDir = root,
            TomlPath = tomlPath,
            ProjectName = name,
            WorkbookPath = workbookPath,
            Visible = visible,
            SessionMode = sessionMode,
            SrcRoot = Path.GetFullPath(Path.Combine(root, srcRel)),
            DistVersion = GetString(doc, "dist", "version"),
            DistExclude = GetStringList(doc, "dist", "exclude"),
            DistBootstrap = GetString(doc, "dist", "bootstrap"),
            DistStripComments = GetBool(doc, "dist", "stripComments") ?? false,
            BackupKeep = backupKeep,
            PushCompile = pushCompile,
            ManagedBy = GetString(doc, ManagedStamp.Section, ManagedStamp.Key),
        };
        Loaded?.Invoke(cfg);
        return cfg;
    }

    /// <summary>toml 스탬프가 이 exe보다 새로우면 그 스탬프, 아니면 null.</summary>
    private static string? ExeOlderThan(string tomlText)
    {
        var stamp = ManagedStamp.Read(tomlText);
        return ManagedStamp.Decide(ExeVersion.Current, stamp) == StampState.ExeOlder ? stamp : null;
    }

    /// <summary>[backup] keep: 0 이상 정수만. 타입 오류는 조용히 기본값으로 떨어지지 않고 실패한다 —
    /// keep = "0"(문자열)을 무시하고 백업을 계속 쌓으면 사용자의 비활성 의도를 어기기 때문.</summary>
    private static int GetBackupKeep(Dictionary<string, Dictionary<string, object>> doc, string tomlPath)
    {
        if (!doc.TryGetValue("backup", out var s) || !s.TryGetValue("keep", out var v))
            return DefaultBackupKeep;
        if (v is not int keep)
            throw new CliError($"{tomlPath}: [backup].keep must be an integer", "keep = 10 (0 disables the pre-push backup)");
        if (keep < 0)
            throw new CliError($"{tomlPath}: [backup].keep must be >= 0 (currently {keep})", "keep = 10 (0 disables the pre-push backup)");
        return keep;
    }

    /// <summary>[excel].session: "ephemeral" | "persistent". 오타를 조용히 기본값으로 떨어뜨리지 않는다.</summary>
    private static SessionMode GetSessionMode(Dictionary<string, Dictionary<string, object>> doc, string tomlPath)
    {
        if (!doc.TryGetValue("excel", out var s) || !s.TryGetValue("session", out var v))
            return SessionMode.Ephemeral;
        return (v as string)?.ToLowerInvariant() switch
        {
            "ephemeral" => SessionMode.Ephemeral,
            "persistent" => SessionMode.Persistent,
            _ => throw new CliError($"{tomlPath}: [excel].session must be \"ephemeral\" or \"persistent\"",
                "session = \"ephemeral\" (hidden Excel per command) | \"persistent\" (keep a session Excel)"),
        };
    }

    private static bool? GetStrictBool(Dictionary<string, Dictionary<string, object>> doc, string section, string key,
        string tomlPath)
    {
        if (!doc.TryGetValue(section, out var s) || !s.TryGetValue(key, out var v)) return null;
        return v as bool? ?? throw new CliError($"{tomlPath}: [{section}].{key} must be true/false");
    }

    private static string? GetString(Dictionary<string, Dictionary<string, object>> doc, string section, string key)
        => doc.TryGetValue(section, out var s) && s.TryGetValue(key, out var v) ? v as string : null;

    private static bool? GetBool(Dictionary<string, Dictionary<string, object>> doc, string section, string key)
        => doc.TryGetValue(section, out var s) && s.TryGetValue(key, out var v) && v is bool b ? b : null;

    private static IReadOnlyList<string> GetStringList(Dictionary<string, Dictionary<string, object>> doc, string section, string key)
        => doc.TryGetValue(section, out var s) && s.TryGetValue(key, out var v) && v is List<string> l
            ? l : Array.Empty<string>();
}

/// <summary>세션 필요 명령(push/pull/test/run/diff)이 쓰는 Excel 수명(설계 §4 세션 모드).</summary>
public enum SessionMode
{
    /// <summary>명령마다 숨은 Excel을 띄우고 끝나면 닫는다. 살아 있는 명시 세션(session start/attach)이 있으면 그것을 쓴다.</summary>
    Ephemeral,
    /// <summary>auto-ensure가 세션 Excel을 띄워 명령 사이에 유지한다(구 기본 동작).</summary>
    Persistent,
}
