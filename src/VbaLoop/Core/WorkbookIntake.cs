namespace VbaLoop.Core;

/// <summary>
/// setup에 워크북 파일을 직접 넘긴 경우(exe에 드래그 앤 드롭, 파일 선택 대화상자, --workbook)의
/// 프로젝트 폴더 결정(설계 §10.2). 원본은 건드리지 않는다 — 원본 옆에 &lt;이름&gt;\ 폴더를 만들어
/// 사본을 두고 그 폴더를 프로젝트로 쓴다. 사용자가 "빈 폴더를 만들고 워크북을 혼자 넣는" 준비
/// 단계를 없애는 것이 목적이고, setup 재실행은 같은 폴더를 재사용해 멱등이다.
/// </summary>
public static class WorkbookIntake
{
    public enum Kind
    {
        /// <summary>원본이 이미 vbaloop 프로젝트 안에 있다 — 그 자리에서 그대로 쓴다.</summary>
        InPlace,
        /// <summary>이전 setup이 만든 사본 폴더가 있다 — 다시 복사하지 않고 재사용한다.</summary>
        Reuse,
        /// <summary>새 폴더(또는 지정 폴더)에 사본을 만든다.</summary>
        Copy,
    }

    /// <param name="Folder">프로젝트 폴더 절대 경로.</param>
    /// <param name="WorkbookName">그 폴더 안의 워크북 파일명.</param>
    /// <param name="Source">원본 워크북 절대 경로(Copy일 때 복사 원천).</param>
    public sealed record Plan(Kind Kind, string Folder, string WorkbookName, string Source);

    /// <summary>탐색기 드래그 앤 드롭 인자인가 — 명령(verb)은 경로 구분자나 점을 포함하지 않는다.</summary>
    public static bool LooksLikePath(string arg)
        => !arg.StartsWith('-') && arg.IndexOfAny(new[] { '\\', '/', '.', ':' }) >= 0;

    /// <summary>
    /// 사용 가능한 워크북인지 검사하고 절대 경로를 반환한다. 콘솔 창에 파일을 끌어다 놓으면
    /// 따옴표로 감싼 경로가 입력되므로 바깥 따옴표를 벗긴다.
    /// </summary>
    public static string Validate(string rawPath)
    {
        string path = Path.GetFullPath(rawPath.Trim().Trim('"'));
        string name = Path.GetFileName(path);
        string ext = Path.GetExtension(path).ToLowerInvariant();

        if (name.StartsWith("~$"))
            throw new CliError($"this is an Excel lock file (starts with ~$): {name}",
                "pick the real workbook file without the ~$ prefix");
        switch (ext)
        {
            case ".xlsm":
                break;
            case ".xlsx" or ".xls" or ".xlsb":
                throw new CliError($"{ext} workbooks are not supported: {name}",
                    "in Excel, use 'Save As' > Excel Macro-Enabled Workbook (*.xlsm), then rerun with that file");
            case ".xlam":
                throw new CliError($"an add-in (.xlam) cannot be used as the development source: {name}",
                    "run setup with the add-in's source .xlsm, and build the add-in with vbaloop dist --out <name>.xlam");
            default:
                throw new CliError($"not an Excel macro workbook (.xlsm): {name}",
                    "vbaloop setup --workbook <path.xlsm>");
        }
        if (!File.Exists(path))
            throw new CliError($"workbook file not found: {path}", "check the path, then rerun");
        return path;
    }

    /// <summary>
    /// 프로젝트 폴더 결정(파일시스템 읽기만, 쓰기 없음).
    /// <list type="number">
    /// <item>폴더 지정 없음 + 원본 폴더에 vbaloop.toml → InPlace(이미 개발 중인 워크북)</item>
    /// <item>폴더 지정 = 원본 폴더 → InPlace</item>
    /// <item>후보 폴더(지정 폴더, 또는 &lt;원본폴더&gt;\&lt;이름&gt;, &lt;이름&gt; (2), ...)를 차례로 본다:
    ///   없거나 비었으면 Copy, 같은 이름 워크북이 있고 vbaloop.toml이 있거나 내용이 원본과 같으면
    ///   Reuse(이전 setup), 그 외에는 다음 후보. 지정 폴더는 하나뿐이라 충돌이면 실패한다.</item>
    /// </list>
    /// </summary>
    public static Plan Decide(string sourcePath, string? folderOverride)
    {
        string source = Path.GetFullPath(sourcePath);
        string name = Path.GetFileName(source);
        string sourceDir = Path.GetDirectoryName(source)!;

        if (folderOverride == null && HasToml(sourceDir))
            return InPlace(sourceDir, name, source);

        if (folderOverride != null)
        {
            string folder = Path.GetFullPath(folderOverride);
            if (SamePath(folder, sourceDir))
                return InPlace(folder, name, source);
            var verdict = Inspect(folder, name, source);
            return verdict switch
            {
                Slot.Empty => new Plan(Kind.Copy, folder, name, source),
                Slot.Reusable => new Plan(Kind.Reuse, folder, name, source),
                Slot.Free => new Plan(Kind.Copy, folder, name, source),
                Slot.OtherWorkbook => throw new CliError(
                    $"the specified folder contains another workbook (*.xlsm): {folder}",
                    "specify an empty folder, or omit --folder to create a new folder next to the original"),
                Slot.NotADirectory => throw new CliError(
                    $"the specified path is a file, not a folder: {folder}",
                    "specify a folder path, or omit --folder to create a new folder next to the original"),
                _ => throw new CliError(
                    $"the specified folder contains a different workbook with the same name: {Path.Combine(folder, name)}",
                    "will not overwrite — specify another folder or remove that file, then rerun"),
            };
        }

        string stem = Path.GetFileNameWithoutExtension(source);
        string baseFolder = Path.Combine(sourceDir, stem);
        for (int n = 1; n < 100; n++)
        {
            string folder = n == 1 ? baseFolder : $"{baseFolder} ({n})";
            switch (Inspect(folder, name, source))
            {
                case Slot.Empty:
                    return new Plan(Kind.Copy, folder, name, source);
                case Slot.Reusable:
                    return new Plan(Kind.Reuse, folder, name, source);
            }
            // Free(관계없는 파일이 든 폴더)·OtherWorkbook·Conflict는 남의 폴더일 수 있어 건너뛴다
        }
        throw new CliError($"could not choose a copy folder name: {baseFolder} (2)~(99) are all in use",
            "vbaloop setup --workbook <path> --folder <empty folder>");
    }

    /// <summary>Copy 계획을 실행한다. 사본의 읽기 전용 특성은 해제한다 — push가 저장할 수 없게 되므로.</summary>
    public static void Execute(Plan plan)
    {
        if (plan.Kind != Kind.Copy) return;
        Directory.CreateDirectory(plan.Folder);
        string dest = Path.Combine(plan.Folder, plan.WorkbookName);
        try
        {
            File.Copy(plan.Source, dest, overwrite: false);
        }
        catch (IOException e)
        {
            throw new CliError($"workbook copy failed: {e.Message}",
                "close the original in Excel, then rerun");
        }
        var attrs = File.GetAttributes(dest);
        if ((attrs & FileAttributes.ReadOnly) != 0)
            File.SetAttributes(dest, attrs & ~FileAttributes.ReadOnly);
    }

    private static Plan InPlace(string folder, string name, string source)
    {
        if (HasToml(folder))
        {
            if (!TomlPointsTo(folder, name))
                throw new CliError($"vbaloop.toml in this folder points to a different workbook: {folder}",
                    "check [excel].workbook in vbaloop.toml, or import separately with --folder <empty folder>");
        }
        else if (OtherWorkbooks(folder, name).Any())
        {
            throw new CliError($"the folder also contains another workbook (*.xlsm), so it cannot be used in place: {folder}",
                "omit --folder to proceed with a copy in a new folder next to the original");
        }
        return new Plan(Kind.InPlace, folder, name, source);
    }

    private enum Slot { Empty, Reusable, Free, OtherWorkbook, Conflict, NotADirectory }

    private static Slot Inspect(string folder, string name, string source)
    {
        // 그 경로에 폴더가 아닌 파일이 있으면(예: 원본 옆의 확장자 없는 "Book") 폴더를 만들 수 없다
        if (File.Exists(folder))
            return Slot.NotADirectory;
        if (!Directory.Exists(folder) || !Directory.EnumerateFileSystemEntries(folder).Any())
            return Slot.Empty;

        string candidate = Path.Combine(folder, name);
        if (File.Exists(candidate))
        {
            // 이전 setup의 흔적: 온보딩까지 끝났으면 toml이, 복사 직후 실패했으면 원본과 같은 사본이 있다
            if (HasToml(folder)) return TomlPointsTo(folder, name) ? Slot.Reusable : Slot.Conflict;
            return SameContent(candidate, source) ? Slot.Reusable : Slot.Conflict;
        }
        return OtherWorkbooks(folder, name).Any() || HasToml(folder) ? Slot.OtherWorkbook : Slot.Free;
    }

    private static IEnumerable<string> OtherWorkbooks(string folder, string name)
        => Directory.EnumerateFiles(folder, "*.xlsm", SearchOption.TopDirectoryOnly)
            .Select(p => Path.GetFileName(p))
            .Where(n => !n.StartsWith("~$") && !string.Equals(n, name, StringComparison.OrdinalIgnoreCase));

    private static bool TomlPointsTo(string folder, string name)
        => string.Equals(ProjectConfig.Load(null, folder).WorkbookPath, Path.Combine(folder, name),
            StringComparison.OrdinalIgnoreCase);

    private static bool HasToml(string dir) => File.Exists(Path.Combine(dir, "vbaloop.toml"));

    private static bool SamePath(string a, string b)
        => string.Equals(Path.TrimEndingDirectorySeparator(a), Path.TrimEndingDirectorySeparator(b),
            StringComparison.OrdinalIgnoreCase);

    private static bool SameContent(string a, string b)
    {
        var fa = new FileInfo(a);
        var fb = new FileInfo(b);
        if (fa.Length != fb.Length) return false;
        using var sa = new FileStream(a, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var sb = new FileStream(b, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var ba = new byte[81920];
        var bb = new byte[81920];
        while (true)
        {
            int na = sa.ReadAtLeast(ba, ba.Length, throwOnEndOfStream: false);
            int nb = sb.ReadAtLeast(bb, bb.Length, throwOnEndOfStream: false);
            if (na != nb) return false;
            if (na == 0) return true;
            if (!ba.AsSpan(0, na).SequenceEqual(bb.AsSpan(0, nb))) return false;
        }
    }
}
