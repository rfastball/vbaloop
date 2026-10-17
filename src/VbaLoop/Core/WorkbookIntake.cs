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
            throw new CliError($"Excel 잠금 파일임(~$로 시작): {name}",
                "~$ 없는 실제 워크북 파일을 고르세요");
        switch (ext)
        {
            case ".xlsm":
                break;
            case ".xlsx" or ".xls" or ".xlsb":
                throw new CliError($"{ext} 워크북은 지원하지 않음: {name}",
                    "Excel에서 '다른 이름으로 저장' > Excel 매크로 사용 통합 문서(*.xlsm)로 저장한 뒤 그 파일로 다시 실행");
            case ".xlam":
                throw new CliError($"애드인(.xlam)은 개발 원본으로 쓸 수 없음: {name}",
                    "애드인의 원본 .xlsm으로 setup하고, 애드인은 vbaloop dist --out <이름>.xlam으로 빌드");
            default:
                throw new CliError($"Excel 매크로 워크북(.xlsm)이 아님: {name}",
                    "vbaloop setup --workbook <경로.xlsm>");
        }
        if (!File.Exists(path))
            throw new CliError($"워크북 파일을 찾지 못함: {path}", "경로를 확인한 뒤 다시 실행");
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
                    $"지정 폴더에 다른 워크북(*.xlsm)이 있음: {folder}",
                    "빈 폴더를 지정하거나 --folder를 빼서 원본 옆에 새 폴더를 만들게 하세요"),
                _ => throw new CliError(
                    $"지정 폴더에 같은 이름의 다른 워크북이 있음: {Path.Combine(folder, name)}",
                    "덮어쓰지 않습니다 — 다른 폴더를 지정하거나 그 파일을 치운 뒤 다시 실행"),
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
        throw new CliError($"사본 폴더 이름을 정하지 못함: {baseFolder} (2)~(99)가 모두 사용 중",
            "vbaloop setup --workbook <경로> --folder <빈 폴더>");
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
            throw new CliError($"워크북 복사 실패: {e.Message}",
                "원본을 Excel에서 닫은 뒤 다시 실행");
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
                throw new CliError($"이 폴더의 vbaloop.toml은 다른 워크북을 가리킴: {folder}",
                    "vbaloop.toml의 [excel].workbook을 확인하거나, --folder <빈 폴더>로 따로 가져오세요");
        }
        else if (OtherWorkbooks(folder, name).Any())
        {
            throw new CliError($"폴더에 다른 워크북(*.xlsm)도 있어 그 자리에서 쓸 수 없음: {folder}",
                "--folder를 빼면 원본 옆 새 폴더에 사본을 만들어 진행합니다");
        }
        return new Plan(Kind.InPlace, folder, name, source);
    }

    private enum Slot { Empty, Reusable, Free, OtherWorkbook, Conflict }

    private static Slot Inspect(string folder, string name, string source)
    {
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

    private static bool HasToml(string dir)
        => File.Exists(Path.Combine(dir, "vbaloop.toml"))
           || File.Exists(Path.Combine(dir, LegacyNames.TomlFile));

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
