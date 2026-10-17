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
/// 그 밖의 모든 불확실성(Busy, Unresponsive, 신원 Unknown, external)은 격리다.
/// </summary>
public static class RecoveryPolicy
{
    public static AutoDecision DecideAuto(bool external, IdentityState identity, ProbeResult? probe,
        string sessionWorkbook)
    {
        if (identity == IdentityState.Gone)
            return new(AutoAction.Clear, "기록된 Excel 프로세스가 이미 종료됨");
        if (external)
            return new(AutoAction.Quarantine, "attach한 사용자 Excel은 자동 정리 대상이 아님");
        if (identity == IdentityState.Unknown)
            return new(AutoAction.Quarantine, "프로세스 신원(시작 시각)을 확인할 수 없음");
        if (probe == null)
            return new(AutoAction.Quarantine, "프로브 결과 없음");

        switch (probe.Kind)
        {
            case ProbeKind.Windowless:
                return new(AutoAction.Kill, "창 없는 Excel(워크북 창 0개)");
            case ProbeKind.Responsive:
                var foreign = ForeignWorkbooks(probe.Workbooks, sessionWorkbook);
                return foreign.Count == 0
                    ? new(AutoAction.Kill, probe.Workbooks.Count == 0
                        ? "열린 워크북 없음"
                        : "세션 워크북만 열려 있음")
                    : new(AutoAction.Quarantine, $"세션 외 워크북이 열려 있음: {string.Join(", ", foreign)}");
            case ProbeKind.Busy:
                return new(AutoAction.Quarantine, "Excel이 busy(셀 편집/모달 대화상자 의심) — 워크북 목록 확인 불가"
                                                  + Suffix(probe.Detail));
            default:
                return new(AutoAction.Quarantine, "Excel이 응답하지 않음 — 워크북 목록 확인 불가" + Suffix(probe.Detail));
        }
    }

    /// <summary>recovery clear: 기록된 프로세스 소멸 확인 시에만 해제, --force는 경고와 함께 해제.</summary>
    public static ClearDecision DecideClear(bool corrupt, IdentityState identity, bool force)
    {
        if (corrupt)
            return force
                ? new(ClearAction.ForceClear, "손상된 마커를 --force로 해제")
                : new(ClearAction.Refuse, "마커가 손상돼 기록된 프로세스를 확인할 수 없음");
        if (identity == IdentityState.Gone)
            return new(ClearAction.Clear, "기록된 Excel 프로세스 종료 확인");
        return force
            ? new(ClearAction.ForceClear, "기록된 Excel 프로세스가 아직 살아있으나 --force로 해제")
            : new(ClearAction.Refuse, identity == IdentityState.Match
                ? "기록된 Excel 프로세스가 아직 살아있음"
                : "기록된 PID가 살아있고 신원을 확인할 수 없음");
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
            return new(DiscardAction.Detach, "attach 세션 — 분리만 하고 사용자 Excel은 그대로 둠");
        if (identity == IdentityState.Gone)
            return new(DiscardAction.CleanOnly, "기록된 Excel 프로세스가 이미 종료됨");
        if (identity == IdentityState.Unknown)
            return new(DiscardAction.Refuse, "프로세스 신원(시작 시각)을 확인할 수 없어 종료하지 않음");
        if (probe is { Kind: ProbeKind.Responsive })
        {
            var foreign = ForeignWorkbooks(probe.Workbooks, sessionWorkbook);
            if (foreign.Count > 0)
                return new(DiscardAction.Refuse, $"세션 외 워크북이 열려 있어 종료하지 않음: {string.Join(", ", foreign)}");
        }
        return new(DiscardAction.Kill, "도구 기동 세션, 신원 일치");
    }

    public static List<string> ForeignWorkbooks(IEnumerable<string> workbooks, string sessionWorkbook)
        => workbooks.Where(w => !PathDrift.SamePath(w, sessionWorkbook)).ToList();

    private static string Suffix(string? detail) => string.IsNullOrEmpty(detail) ? "" : $" ({detail})";
}
