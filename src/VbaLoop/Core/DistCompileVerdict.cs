namespace VbaLoop.Core;

/// <summary>
/// dist 게시 전 VBE 컴파일 검증의 판정(순수, 설계 §9.8). 관측값은 DistCommand가 임시 Excel에서
/// 모은다: VBE CommandBars의 "VBAProject 컴파일"(FindControl Id=578) 존재/Enabled, Execute 예외,
/// ModalWatchdog이 잡은 VBE 모달 본문, 실행 후 Enabled, 타임아웃 kill 여부.
/// 판정 원칙: 컴파일 오류는 VBE 모달(#32770 "Microsoft Visual Basic ...")로만 드러나므로 모달이 곧 오류다.
/// Enabled=false는 "컴파일할 것 없음(이미 컴파일됨)". 관측 불능(컨트롤 없음/Execute 예외/타임아웃)은
/// 게시하지 않는다(fail closed — 탈출구는 --skip-compile). 모달 없이 끝났는데 Enabled가 그대로면
/// 오류 신호는 없으므로 게시하되 경고한다(숨은 VBE에서 상태 갱신이 늦는 경우 대비).
/// </summary>
public static class DistCompileVerdict
{
    /// <summary>VBE 메뉴 디버그 &gt; "VBAProject 컴파일" 컨트롤 Id(언어 무관).</summary>
    public const int CompileControlId = 578;

    public enum Kind { Compiled, AlreadyCompiled, Unverified, CompileError, CheckFailed }

    public sealed record Verdict(Kind Kind, string? Detail)
    {
        public bool Publishable => Kind is Kind.Compiled or Kind.AlreadyCompiled or Kind.Unverified;
        public string Label => Kind switch
        {
            Kind.Compiled => "compiled",
            Kind.AlreadyCompiled => "already_compiled",
            Kind.Unverified => "unverified",
            Kind.CompileError => "compile_error",
            _ => "check_failed",
        };
    }

    public static Verdict Decide(
        bool controlFound, bool enabledBefore, IReadOnlyList<string> modalTexts,
        string? executeError, bool? enabledAfter, bool timedOut)
    {
        // 모달이 최우선 — 워치독이 모달을 닫으면 Execute가 예외로 돌아올 수 있으므로 예외보다 먼저 본다
        var texts = modalTexts.Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
        if (modalTexts.Count > 0)
            return new Verdict(Kind.CompileError,
                texts.Count > 0 ? string.Join(" || ", texts) : "VBE 컴파일 오류 모달(본문 수집 실패)");
        if (timedOut)
            return new Verdict(Kind.CheckFailed, "컴파일 타임아웃 — 임시 Excel kill(모달 해제 실패 의심)");
        if (!controlFound)
            return new Verdict(Kind.CheckFailed, $"VBE 컴파일 컨트롤(Id {CompileControlId})을 찾지 못함");
        if (!enabledBefore)
            return new Verdict(Kind.AlreadyCompiled, null);
        if (executeError != null)
            return new Verdict(Kind.CheckFailed, $"컴파일 명령 실행 실패: {executeError}");
        if (enabledAfter == null)
            return new Verdict(Kind.Unverified, "컴파일 후 컨트롤 상태 조회 실패 — 오류 모달은 없었음");
        if (enabledAfter == true)
            return new Verdict(Kind.Unverified, "컴파일 후에도 컴파일 컨트롤이 활성 — 오류 모달은 없었음");
        return new Verdict(Kind.Compiled, null);
    }
}
