using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 실행 전에 찍어 두는 Application 설정(설계 §14.8). 매크로가 끝나도 Excel이 되돌리지 않는 것만 담는다 —
/// ScreenUpdating·DisplayAlerts·EnableCancelKey는 매크로가 끝나면 Excel이 스스로 되돌린다.
/// StatusBar null = Excel이 표시줄을 쥔 상태(StatusBar = False).
/// </summary>
public sealed record AppSettings(bool EnableEvents, int Calculation, string? StatusBar, bool Interactive, int Cursor)
{
    /// <summary>Application.StatusBar 값 → 매크로가 쓴 문구, Excel이 쥐었으면 null. 한 번이라도 문구를 쓴 뒤
    /// StatusBar = False로 돌려주면 Excel은 Boolean이 아니라 문자열 "FALSE"를 돌려준다(VBA에서도 같다, 실측) — 같은 상태로 본다.</summary>
    public static string? StatusBarText(object? value)
        => value is string s && !s.Equals("FALSE", StringComparison.OrdinalIgnoreCase) ? s : null;

    /// <summary>before와 값이 달라진 설정 이름(되돌릴 대상).</summary>
    public static IReadOnlyList<string> Changed(AppSettings before, AppSettings after)
    {
        var names = new List<string>();
        if (before.EnableEvents != after.EnableEvents) names.Add("EnableEvents");
        if (before.Calculation != after.Calculation) names.Add("Calculation");
        if (before.StatusBar != after.StatusBar) names.Add("StatusBar");
        if (before.Interactive != after.Interactive) names.Add("Interactive");
        if (before.Cursor != after.Cursor) names.Add("Cursor");
        return names;
    }
}

/// <summary>
/// 중간에 끊긴 매크로(취소·런타임 오류 = VBA End)의 뒷정리 결과(설계 §14.8). 매크로의 정리 코드가 돌지 않아
/// 남은 것 중 vbaloop가 알 수 있는 것만 다룬다: 되돌린 Application 설정과 매크로가 연 채 남긴 통합 문서.
/// </summary>
public sealed record RunAftermath(
    IReadOnlyList<string> Restored, IReadOnlyList<string> Closed, IReadOnlyList<string> LeftOpen,
    IReadOnlyList<string> Failed)
{
    [JsonIgnore]
    public bool IsEmpty => Restored.Count == 0 && Closed.Count == 0 && LeftOpen.Count == 0 && Failed.Count == 0;

    /// <summary>사람용 한 줄(payload에도 summary로 싣는다).</summary>
    public string Summary
    {
        get
        {
            var parts = new List<string>();
            if (Restored.Count > 0)
                parts.Add($"restored Application.{string.Join(", Application.", Restored)} to the value before the run");
            if (Closed.Count > 0)
                parts.Add($"closed without saving {Closed.Count} workbook(s) the macro left open ({string.Join(", ", Closed)})");
            if (LeftOpen.Count > 0)
                parts.Add($"the macro left {LeftOpen.Count} workbook(s) open ({string.Join(", ", LeftOpen)}) - not closed "
                          + "(a kept session, or a saved file with unsaved changes)");
            if (Failed.Count > 0)
                parts.Add($"could not undo: {string.Join(", ", Failed)}");
            return string.Join("; ", parts);
        }
    }

    /// <summary>실행 뒤에 새로 보이는 통합 문서(FullName, 대소문자 무시). 매크로가 닫은 기존 문서는 다루지 않는다.</summary>
    public static IReadOnlyList<string> NewWorkbooks(IEnumerable<string> before, IEnumerable<string> after)
    {
        var known = new HashSet<string>(before, StringComparer.OrdinalIgnoreCase);
        return after.Where(n => !known.Contains(n)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }
}
