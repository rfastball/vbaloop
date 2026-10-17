namespace VbaLoop.Core;

/// <summary>
/// setup의 새 프로젝트 경로(설계 §10.2): 이름 하나만 받아 &lt;부모&gt;\&lt;이름&gt;\ 폴더를 정한다.
/// 이후 단계는 빈 폴더 흐름 그대로 — &lt;이름&gt;.xlsm 생성 + 샘플 테스트 + push/test 초록불.
/// 부모는 셸에서 부르면 cwd(`vbaloop setup --new Foo` = .\Foo), 탐색기 더블클릭이면 cwd가 exe
/// 위치(다운로드 폴더 등)라 의미가 없어 고정 폴더를 쓴다. OneDrive로 리디렉션되기 쉬운
/// 문서/바탕화면은 피한다 — 동기화 중인 워크북은 Excel이 경로를 URL로 보고하고 파일 잠금도 생긴다.
/// </summary>
public static class NewProjectFolder
{
    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON", "PRN", "AUX", "NUL",
        "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    };

    /// <summary>탐색기 더블클릭 실행의 기본 부모 폴더: %USERPROFILE%\vbaloop-projects.</summary>
    public static string ExplorerDefaultParent
        => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "vbaloop-projects");

    /// <summary>프로젝트 이름(= 폴더명 = 워크북 파일명)으로 쓸 수 있는지 검사하고 정리한 값을 반환한다.</summary>
    public static string ValidateName(string raw)
    {
        string name = raw.Trim().Trim('"').Trim();
        if (name.Length == 0)
            throw new CliError("프로젝트 이름이 비어 있음", "vbaloop setup --new <이름>");
        if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            throw new CliError($"프로젝트 이름에 쓸 수 없는 문자가 있음: {name}",
                "폴더 이름으로 쓸 수 있는 이름만 가능 (\\ / : * ? \" < > | 제외)");
        if (name is "." or ".." || name.EndsWith('.'))
            throw new CliError($"프로젝트 이름은 점으로 끝날 수 없음: {name}", "다른 이름으로 다시 실행");
        if (ReservedNames.Contains(name))
            throw new CliError($"Windows 예약 이름은 쓸 수 없음: {name}", "다른 이름으로 다시 실행");
        return name;
    }

    /// <summary>
    /// &lt;부모&gt;\&lt;이름&gt; 절대 경로. 이미 있고 비어 있지 않으면 거부한다 — 남의 폴더에 워크북을
    /// 만들지 않으며, 중단된 setup을 이어 가려는 경우는 --folder로 그 폴더를 직접 지정한다.
    /// </summary>
    public static string Resolve(string parent, string name)
    {
        string folder = Path.GetFullPath(Path.Combine(parent, ValidateName(name)));
        if (Directory.Exists(folder) && Directory.EnumerateFileSystemEntries(folder).Any())
            throw new CliError($"같은 이름의 폴더가 이미 있음: {folder}",
                $"다른 이름을 쓰거나, 그 폴더로 이어서 진행하려면 vbaloop setup --folder \"{folder}\"");
        return folder;
    }
}
