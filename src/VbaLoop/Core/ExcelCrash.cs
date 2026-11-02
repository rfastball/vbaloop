namespace VbaLoop.Core;

/// <summary>세션 Excel이 명령 도중 사라진 방식 — Alive(살아 있음), Exited(크래시 징후 없는 종료: 매크로의 Quit 등),
/// Crashed(크래시 징후 있음 — 크래시 의심 오류, 0 아닌 종료 코드, 또는 Office의 재기동).</summary>
public enum ExcelExitKind { Alive, Exited, Crashed }

/// <summary>크래시 1건 — 기록된 PID, 종료 코드(판독 불가면 null), Office 크래시 핸들러가 다시 띄운 Excel PID들,
/// 확인 시점에 정본의 `~$` 잠금 파일이 있었는지.</summary>
public sealed record ExcelCrashInfo(int Pid, int? ExitCode, IReadOnlyList<int> Relaunched, bool WorkbookLocked = false);

/// <summary>
/// Excel 크래시 판정과 안내 문구의 순수부(단위 테스트 대상). 실측(2026-11-02, 일회용 모드): 매크로의 액세스
/// 위반은 Application.Run에 DISP_E_EXCEPTION(0x80020009)으로 돌아오고 프로세스는 곧 사라진다. Office 크래시
/// 핸들러는 크래시한 Excel을 부모로 하는 새 EXCEL.EXE를 보이는 창으로 띄운다. 그 Excel은 정본을 쓰기로 다시 열고
/// (`~$` 잠금) "(version 1) [복구됨]" 사본도 열거나(첫 크래시), 복구 화면만 띄운다(같은 파일 반복 크래시) — 정본을
/// 쥐고 있는 동안 push는 read-only로 거부된다.
/// 그 Excel은 사용자에게 보이는 복구 상태를 들고 있으므로 도구가 kill하지 않는다(설계 §11).
/// </summary>
public static class ExcelCrash
{
    /// <summary>
    /// 종료 코드만으로는 못 가른다 — 실측: Office 크래시 핸들러를 거친 Excel도 종료 코드 0으로 끝난다. 그래서 크래시
    /// 징후(실행 중 호출이 크래시 의심 오류로 끝남, 0 아닌 종료 코드, 크래시한 Excel을 부모로 한 재기동 Excel)가
    /// 하나라도 있으면 Crashed, 없으면 Exited다. 종료 코드를 못 읽은 것(null)은 징후가 아니다.
    /// </summary>
    public static ExcelExitKind Classify(bool exited, int? exitCode, bool crashSuspectError, bool relaunched = false)
    {
        if (!exited) return ExcelExitKind.Alive;
        bool crashCode = exitCode is int c && c != 0;
        return crashSuspectError || crashCode || relaunched ? ExcelExitKind.Crashed : ExcelExitKind.Exited;
    }

    /// <summary>
    /// 크래시 핸들러가 다시 띄운 Excel 후보: 이름이 excel이고 부모 PID가 크래시한 PID이며, 시작 시각이 크래시한
    /// 프로세스보다 늦은 것(부모 PID 재사용으로 무관한 프로세스를 오인하지 않도록). 시작 시각을 모르면 제외한다.
    /// </summary>
    public static IReadOnlyList<int> FindRelaunched(
        IEnumerable<ProcessTree.ProcEntry> procs, int crashedPid, DateTime? crashedStartUtc, Func<int, DateTime?> startUtcOf)
    {
        var found = new List<int>();
        foreach (var p in procs)
        {
            if (p.ParentPid != crashedPid || p.Pid == crashedPid || p.Name != "excel") continue;
            var start = startUtcOf(p.Pid);
            if (start == null) continue;
            if (crashedStartUtc != null && start.Value < crashedStartUtc.Value) continue;
            found.Add(p.Pid);
        }
        found.Sort();
        return found;
    }

    public static string FormatExitCode(int? exitCode)
        => exitCode is int c ? $"0x{unchecked((uint)c):X8}" : "unknown";

    /// <summary>실패 detail — what은 매크로/테스트 전체 이름 또는 명령 단계. 종료 코드 0은 크래시 핸들러를 거친
    /// 정상 값이라(실측) 오해를 부르지 않도록 0이 아닐 때만 붙인다.</summary>
    public static string Detail(string what, int pid, int? exitCode)
        => $"Excel crashed while running {what} (PID {pid} exited" + CodeSuffix(exitCode) + ")";

    /// <summary>0이 아닌 종료 코드만 " with code 0x..."로.</summary>
    public static string CodeSuffix(int? exitCode)
        => exitCode is int c && c != 0 ? $" with code {FormatExitCode(c)}" : "";

    /// <summary>종료 코드 0으로 사라진 경우(매크로의 Application.Quit 등) — 크래시가 아니지만 세션은 끝났다.</summary>
    public static string ExitedDetail(string what, int pid)
        => $"Excel exited while running {what} (PID {pid}, no crash signs — did the code call Application.Quit?)";

    /// <summary>재기동된 Excel이 있을 때의 사람용 경고 한 줄. 재기동 Excel이 정본을 다시 여는지는 경우마다 달라(실측)
    /// 확인 시점의 `~$` 잠금 여부로 말을 가른다.</summary>
    public static string RelaunchNote(IReadOnlyList<int> relaunched, string workbookPath, bool workbookLocked)
    {
        if (relaunched.Count == 0) return "";
        string pids = string.Join(", ", relaunched);
        string name = Path.GetFileName(workbookPath);
        return $"Office relaunched Excel after the crash (PID {pids}, visible) to recover documents — "
               + (workbookLocked
                   ? $"{name} is locked (~${name} present), likely reopened there with a recovered copy; "
                   : $"it may reopen {name} with a recovered copy; ")
               + "while it holds the workbook, push is refused (read-only) until the user closes it";
    }

    /// <summary>
    /// 크래시 뒤 할 일. 재기동 Excel은 사용자 화면의 복구 상태라 kill 금지 — 사용자에게 알리고 닫아 달라고 한다.
    /// restartNote는 세션 모드별 꼬리(일회용이면 재실행, persistent면 session start).
    /// </summary>
    public static string Hint(IReadOnlyList<int> relaunched, string restartNote)
    {
        const string fix = "the code (or the VBE operation) crashed Excel itself — look for API Declare calls, "
                           + "memory/pointer access or object use after release in what just ran, fix src, push, then ";
        if (relaunched.Count == 0)
            return fix + restartNote;
        return $"Do NOT kill the relaunched Excel (PID {string.Join(", ", relaunched)}) yourself; inform the user: it may show the workbook "
               + "and a recovered copy, and push stays refused while it holds the workbook — the user closes that Excel and decides "
               + "whether to keep the recovered copy. Meanwhile, " + fix + restartNote;
    }
}
