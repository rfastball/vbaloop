namespace VbaLoop.Core;

/// <summary>
/// 세션 Excel 프로브 결과(설계 §11). 프로브는 자식 프로세스에서 벽시계 제한 하에 돈다 —
/// Windowless: 프로세스에 보이는 최상위 창도 EXCEL7(워크북) 창도 없음.
/// Responsive: COM 응답, Workbooks 목록 확보. Busy: RPC_E_CALL_REJECTED 계열(셀 편집·모달).
/// Unresponsive: 시간 초과/attach 실패/판독 불가.
/// </summary>
public enum ProbeKind { Windowless, Responsive, Busy, Unresponsive }

public sealed record ProbeResult(ProbeKind Kind, IReadOnlyList<string> Workbooks, string? Detail = null)
{
    public static ProbeResult Unresponsive(string detail) => new(ProbeKind.Unresponsive, Array.Empty<string>(), detail);

    /// <summary>COM 목록을 못 얻었을 때(Busy/Unresponsive) Win32로 읽은 EXCEL7 창 제목(설계 §14.3).
    /// 메시지를 보내지 않는 판독이라 매크로 실행 중인 Excel에서도 읽힌다.</summary>
    public IReadOnlyList<string> WindowTitles { get; init; } = Array.Empty<string>();
}

public enum AutoAction { Clear, Kill, Quarantine }
public sealed record AutoDecision(AutoAction Action, string Why);

public enum ClearAction { Clear, ForceClear, Refuse }
public sealed record ClearDecision(ClearAction Action, string Why);

public enum DiscardAction { Detach, CleanOnly, Kill, Refuse }
public sealed record DiscardDecision(DiscardAction Action, string Why);

/// <summary>
/// 격리/kill 판정의 순수부 (설계 §11 kill 규칙). 부작용 없음 — 실행은 Excel.RecoveryGate.
/// 규칙: 자동 경로의 kill은 (1) 도구 기동 세션 AND (2) PID+StartTime 신원 일치 AND
/// (3) 세션 워크북 외 워크북이 없음을 적극적으로 확인(Windowless 또는 Responsive 목록)한 경우만.
/// 일회용 인스턴스는 Busy/Unresponsive여도 창 제목이 세션 워크북뿐이면 (3)을 충족한다(설계 §14.3).
/// 그 밖의 모든 불확실성(Busy, Unresponsive, 신원 Unknown, external)은 격리다.
/// </summary>
public static class RecoveryPolicy
{
    // EXCEL7 제목 = 워크북 이름 + 다중 창 접미사(":2") + 현지화된 상태 접미사("  [읽기 전용]") — 2026-11-03 실측
    private static readonly System.Text.RegularExpressions.Regex SessionTitleTail =
        new(@"^(:\d+)?(\s+\[[^\]]*\])*$");

    public static AutoDecision DecideAuto(bool external, IdentityState identity, ProbeResult? probe,
        string sessionWorkbook, bool ephemeral = false)
    {
        if (identity == IdentityState.Gone)
            return new(AutoAction.Clear, "recorded Excel process already exited");
        if (external)
            return new(AutoAction.Quarantine, "an attached Excel belongs to the user and is never auto-cleaned");
        if (identity == IdentityState.Unknown)
            return new(AutoAction.Quarantine, "process identity (start time) cannot be verified");
        if (probe == null)
            return new(AutoAction.Quarantine, "no probe result");

        switch (probe.Kind)
        {
            case ProbeKind.Windowless:
                return new(AutoAction.Kill, "windowless Excel (0 workbook windows)");
            case ProbeKind.Responsive:
                var foreign = ForeignWorkbooks(probe.Workbooks, sessionWorkbook);
                return foreign.Count == 0
                    ? new(AutoAction.Kill, probe.Workbooks.Count == 0
                        ? "no open workbooks"
                        : "only the session workbook is open")
                    : new(AutoAction.Quarantine, $"non-session workbooks open: {string.Join(", ", foreign)}");
            case ProbeKind.Busy when ephemeral && OnlySessionWindows(probe.WindowTitles, sessionWorkbook):
            case ProbeKind.Unresponsive when ephemeral && OnlySessionWindows(probe.WindowTitles, sessionWorkbook):
                return new(AutoAction.Kill, "ephemeral Excel is busy, but its workbook windows show only the session workbook");
            case ProbeKind.Busy:
                return new(AutoAction.Quarantine, "Excel is busy (suspected cell edit mode/modal dialog) — cannot list workbooks"
                                                  + Suffix(probe.Detail) + Windows(probe.WindowTitles));
            default:
                return new(AutoAction.Quarantine, "Excel is not responding — cannot list workbooks" + Suffix(probe.Detail)
                                                  + Windows(probe.WindowTitles));
        }
    }

    /// <summary>EXCEL7 창 제목이 세션 워크북의 창인가. 이름이 다르거나(코드가 Caption을 바꾼 경우 포함)
    /// 모르는 꼬리가 붙으면 아니다 — 불확실하면 격리 쪽으로 기운다.</summary>
    public static bool IsSessionWindowTitle(string title, string sessionWorkbook)
    {
        string name = Path.GetFileName(sessionWorkbook);
        return name.Length > 0
            && title.StartsWith(name, StringComparison.OrdinalIgnoreCase)
            && SessionTitleTail.IsMatch(title[name.Length..]);
    }

    public static bool OnlySessionWindows(IReadOnlyList<string> titles, string sessionWorkbook)
        => titles.Count > 0 && titles.All(t => IsSessionWindowTitle(t, sessionWorkbook));

    /// <summary>recovery clear: 기록된 프로세스 소멸 확인 시에만 해제, --force는 경고와 함께 해제.</summary>
    public static ClearDecision DecideClear(bool corrupt, IdentityState identity, bool force)
    {
        if (corrupt)
            return force
                ? new(ClearAction.ForceClear, "corrupt marker cleared with --force")
                : new(ClearAction.Refuse, "marker is corrupt; cannot verify the recorded process");
        if (identity == IdentityState.Gone)
            return new(ClearAction.Clear, "recorded Excel process confirmed exited");
        return force
            ? new(ClearAction.ForceClear, "recorded Excel process still alive, cleared with --force")
            : new(ClearAction.Refuse, identity == IdentityState.Match
                ? "recorded Excel process still alive"
                : "recorded PID is alive and its identity cannot be verified");
    }

    /// <summary>
    /// session stop --discard: external은 분리만(kill 금지). tool은 신원 일치 시 kill —
    /// 단 프로브가 세션 외 워크북을 적극적으로 보여주면 거부(사람 작업 보호). 명시 명령이므로
    /// Busy/Unresponsive는 거부 사유가 아니다(그것이 discard가 존재하는 이유).
    /// </summary>
    public static DiscardDecision DecideDiscard(bool external, IdentityState identity, ProbeResult? probe,
        string sessionWorkbook)
    {
        if (external)
            return new(DiscardAction.Detach, "attached session — detach only, the user's Excel is left as is");
        if (identity == IdentityState.Gone)
            return new(DiscardAction.CleanOnly, "recorded Excel process already exited");
        if (identity == IdentityState.Unknown)
            return new(DiscardAction.Refuse, "process identity (start time) cannot be verified; not killing");
        if (probe is { Kind: ProbeKind.Responsive })
        {
            var foreign = ForeignWorkbooks(probe.Workbooks, sessionWorkbook);
            if (foreign.Count > 0)
                return new(DiscardAction.Refuse, $"non-session workbooks open; not killing: {string.Join(", ", foreign)}");
        }
        return new(DiscardAction.Kill, "tool-launched session, identity matches");
    }

    public static List<string> ForeignWorkbooks(IEnumerable<string> workbooks, string sessionWorkbook)
        => workbooks.Where(w => !PathDrift.SamePath(w, sessionWorkbook)).ToList();

    private static string Suffix(string? detail) => string.IsNullOrEmpty(detail) ? "" : $" ({detail})";

    private static string Windows(IReadOnlyList<string> titles)
        => titles.Count == 0 ? "" : $"; workbook windows: {string.Join(", ", titles)}";
}
