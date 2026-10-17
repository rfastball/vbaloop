using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 사용자 단위 프로젝트 목록(%LOCALAPPDATA%\vbaloop\projects.json, 설계 §12.6). 설치본 exe는 소스 레포
/// 옆에 있지 않아 형제 탐색(RepoLocator)으로 프로젝트를 찾을 수 없다 — 그래서 프로젝트 설정을 읽는 모든
/// 명령이 그 루트를 여기 기록하고, init --all이 이 목록으로 일괄 수렴한다. 기록은 best-effort이고
/// (실패해도 명령은 계속), 이미 있는 루트면 파일을 쓰지 않는다. 프로세스 간 경합은 네임드 뮤텍스로 직렬화.
/// </summary>
public static class ProjectRegistry
{
    private const string MutexName = @"Local\vbaloop-project-registry";
    private const int MaxEntries = 500;

    public sealed class Entry
    {
        [JsonPropertyName("root")] public string Root { get; set; } = "";
        [JsonPropertyName("addedAt")] public string AddedAt { get; set; } = "";
    }

    private sealed class Doc
    {
        [JsonPropertyName("projects")] public List<Entry> Projects { get; set; } = new();
    }

    private static readonly JsonSerializerOptions DocJson = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static string DefaultPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop", "projects.json");

    /// <summary>비교·저장용 정규 경로(전체 경로, 끝 구분자 제거). 대소문자는 보존하고 비교에서 무시한다.</summary>
    public static string Normalize(string root) => Path.GetFullPath(root).TrimEnd('\\', '/');

    public static bool SameRoot(string a, string b)
        => string.Equals(Normalize(a), Normalize(b), StringComparison.OrdinalIgnoreCase);

    public static List<Entry> Load(string path)
    {
        try
        {
            if (!File.Exists(path)) return new();
            return JsonSerializer.Deserialize<Doc>(File.ReadAllText(path))?.Projects ?? new();
        }
        catch
        {
            return new();
        }
    }

    // ---- 순수 연산(테스트 대상) ----

    /// <summary>root가 없으면 추가. (새 목록, 추가 여부). 상한 초과 시 가장 오래된 항목부터 버린다.</summary>
    public static (List<Entry> Entries, bool Added) Upsert(IReadOnlyList<Entry> entries, string root, DateTimeOffset now)
    {
        var list = entries.ToList();
        if (list.Any(e => SameRoot(e.Root, root))) return (list, false);
        list.Add(new Entry { Root = Normalize(root), AddedAt = now.ToString("o") });
        if (list.Count > MaxEntries) list = list.Skip(list.Count - MaxEntries).ToList();
        return (list, true);
    }

    /// <summary>폴더나 vbaloop.toml이 사라진 항목을 걸러낸다. (남은 것, 제거된 것). 중복 루트도 하나로 합친다.</summary>
    public static (List<Entry> Kept, List<Entry> Removed) Prune(IReadOnlyList<Entry> entries, Func<string, bool> isProject)
    {
        var kept = new List<Entry>();
        var removed = new List<Entry>();
        foreach (var e in entries)
        {
            if (string.IsNullOrWhiteSpace(e.Root) || !isProject(e.Root)) removed.Add(e);
            else if (kept.Any(k => SameRoot(k.Root, e.Root))) continue;
            else kept.Add(e);
        }
        return (kept, removed);
    }

    /// <summary>기본 판정: 폴더와 그 안의 vbaloop.toml이 모두 있어야 프로젝트.</summary>
    public static bool IsProject(string root)
    {
        try { return File.Exists(Path.Combine(root, "vbaloop.toml")); }
        catch { return false; }
    }

    /// <summary>
    /// init --all 대상 합집합: 형제 탐색 결과(있으면) ∪ 레지스트리. 같은 루트는 하나로, 이름은 폴더 이름.
    /// 형제 쪽 순서를 먼저 유지하고 레지스트리에만 있는 것을 이름순으로 덧붙인다.
    /// </summary>
    public static List<(string Name, string Root)> Union(
        IReadOnlyList<(string Name, string Root)> siblings, IEnumerable<string> registryRoots)
    {
        var result = siblings.ToList();
        var extra = registryRoots
            .Where(r => !result.Any(s => SameRoot(s.Root, r)))
            .Select(Normalize)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(r => (Name: Path.GetFileName(r), Root: r))
            .OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
            .ThenBy(p => p.Root, StringComparer.OrdinalIgnoreCase);
        result.AddRange(extra);
        return result;
    }

    // ---- 파일 I/O ----

    /// <summary>root가 목록에 없을 때만 잠금 후 기록. 예외를 던지지 않는다.</summary>
    public static void TryAdd(string path, string root)
    {
        try
        {
            if (Load(path).Any(e => SameRoot(e.Root, root))) return; // 빠른 경로: 쓰기 없음
            Locked(() =>
            {
                var (entries, added) = Upsert(Load(path), root, DateTimeOffset.UtcNow);
                if (added) Write(path, entries);
            });
        }
        catch
        {
            // 보조 기록 — 실패가 명령을 깨면 안 된다
        }
    }

    /// <summary>사라진 항목을 지우고 남은 루트와 제거된 루트를 돌려준다(init --all). 기록 실패는 무시.</summary>
    public static (List<string> Kept, List<string> Removed) PruneFile(string path, Func<string, bool>? isProject = null)
    {
        isProject ??= IsProject;
        List<Entry> kept = new(), removed = new();
        try
        {
            Locked(() =>
            {
                var all = Load(path);
                (kept, removed) = Prune(all, isProject);
                if (removed.Count > 0 || kept.Count != all.Count) Write(path, kept);
            });
        }
        catch
        {
            (kept, removed) = Prune(Load(path), isProject);
        }
        return (kept.Select(e => e.Root).ToList(), removed.Select(e => e.Root).ToList());
    }

    private static void Locked(Action action)
    {
        using var mx = new Mutex(false, MutexName);
        bool held;
        try { held = mx.WaitOne(2000); }
        catch (AbandonedMutexException) { held = true; }
        try { action(); }
        finally { if (held) mx.ReleaseMutex(); }
    }

    private static void Write(string path, List<Entry> entries)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var tmp = $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(new Doc { Projects = entries }, DocJson),
            new System.Text.UTF8Encoding(false));
        File.Move(tmp, path, overwrite: true);
    }
}
