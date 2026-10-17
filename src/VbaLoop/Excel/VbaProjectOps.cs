using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>VBComponents 열거/공용 상수 + 코드 주입 프리미티브.</summary>
public static class VbaProjectOps
{
    public const int TypeStdModule = 1;   // vbext_ct_StdModule
    public const int TypeClassModule = 2; // vbext_ct_ClassModule
    public const int TypeMsForm = 3;      // vbext_ct_MSForm
    public const int TypeDocument = 100;  // vbext_ct_Document

    public sealed record ComponentInfo(string Name, int Type, dynamic Comp);

    public static List<ComponentInfo> ListComponents(dynamic vbProject)
    {
        var list = new List<ComponentInfo>();
        dynamic comps = vbProject.VBComponents;
        int n = Convert.ToInt32(comps.Count);
        for (int i = 1; i <= n; i++)
        {
            dynamic c = comps.Item(i);
            list.Add(new ComponentInfo(Convert.ToString(c.Name) ?? "", Convert.ToInt32(c.Type), c));
        }
        return list;
    }

    public static HashSet<string> ComponentNames(dynamic vbProject)
    {
        List<ComponentInfo> comps = ListComponents(vbProject);
        return new HashSet<string>(comps.Select(c => c.Name), StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>빈 모듈 가드 포함 본문 읽기: CountOfLines=0이면 Lines 호출 생략(설계 §4 pull).</summary>
    public static string ReadModuleBody(dynamic codeModule)
    {
        int count = Convert.ToInt32(codeModule.CountOfLines);
        if (count == 0) return "";
        return Convert.ToString(codeModule.Lines[1, count]) ?? "";
    }

    // ---- 코드 주입 프리미티브 (설계 §4 push 2단계 로직 공유 — push는 자체 사본 유지,
    //      dist는 여기를 재사용. 순수 파이프라인 로직 이원화 금지 원칙은 Core.Dist*가 담당) ----

    /// <summary>
    /// 원자적 코드 교체: 스냅샷 → DeleteLines+AddFromString+팬텀 스윕 → 실패 시 스냅샷 재주입 원복.
    /// trailing CRLF를 붙이지 않아 왕복마다 끝 빈 줄이 증식하는 것을 막는다(실측).
    /// </summary>
    public static void ReplaceCode(dynamic codeModule, IReadOnlyList<string> bodyLines)
    {
        string newCode = string.Join("\r\n", bodyLines);
        int count = Convert.ToInt32(codeModule.CountOfLines);
        string snapshot = count > 0 ? Convert.ToString(codeModule.Lines[1, count]) ?? "" : "";
        try
        {
            InjectWithPhantomSweep(codeModule, newCode);
        }
        catch (Exception ex)
        {
            try
            {
                int c = Convert.ToInt32(codeModule.CountOfLines);
                if (c > 0) codeModule.DeleteLines(1, c);
                if (snapshot.Length > 0) codeModule.AddFromString(snapshot);
            }
            catch { /* 원복 실패 — 상위에서 전체 실패 처리 */ }
            throw new CliError($"코드 주입 실패(원복 시도함): {ex.Message}");
        }
    }

    /// <summary>
    /// DeleteLines+AddFromString 주입 후, VBE가 <c>_</c> 컨티뉴에이션 <c>Declare</c> 문을 오파싱해
    /// 모듈 끝에 덧붙이는 <b>컴파일 파손형</b> 고아 <c>()</c> 팬텀 라인을 제거한다(버그리포트
    /// 2026-07-10). 말미 빈 줄은 AddFromString의 무해한 산출물(pull TrimTrailingEmpty가 흡수)이라
    /// 보존하고, 빈 줄을 건너뛴 최하단 줄이 <c>()</c>면 삭제한다. 그 뒤 <b>말미 빈 줄을 제외한</b>
    /// 논리행수를 주입분과 대조 — 불일치면 loud 실패로 조용한 컴파일 파손을 fail-fast 전환한다.
    /// 행수 기반이라 VBE 식별자 재케이싱(2026-07-09)에 면역.
    /// </summary>
    private static void InjectWithPhantomSweep(dynamic codeModule, string newCode)
    {
        ComRetry.Invoke(() =>
        {
            int c = Convert.ToInt32(codeModule.CountOfLines);
            if (c > 0) codeModule.DeleteLines(1, c);
            if (newCode.Length > 0) codeModule.AddFromString(newCode);
        });
        if (newCode.Length == 0) return;

        while (SweepTrailingParenPhantom(codeModule)) { }

        int expected = MeaningfulCount(newCode.Split("\r\n")); // 주입분(말미 빈 줄 제외 논리행수)
        int actual = MeaningfulCountModule(codeModule);
        if (actual != expected)
            throw new CliError(
                $"코드 주입 후 논리행수 불일치(기대 {expected}, 실제 {actual}) — VBE 주입 파손 의심",
                "해당 모듈의 컨티뉴에이션 Declare 등을 점검하거나 수동 Export/Import로 관리");
    }

    /// <summary>
    /// push 실패 복구용 스냅샷 원복: 본문 전체를 스냅샷으로 되돌리고(팬텀 스윕 포함) 다시 읽어 검증한다.
    /// 검증 일치면 true. ReplaceCode의 내부 원복과 달리 결과를 숨기지 않는다 — false/예외면 호출자가
    /// 세션을 poison한다(PushRecovery). 스냅샷 자체가 <c>()</c>로 끝나면 스윕하지 않는다.
    /// </summary>
    public static bool RestoreSnapshot(dynamic codeModule, string snapshot)
    {
        ComRetry.Invoke(() =>
        {
            int c = Convert.ToInt32(codeModule.CountOfLines);
            if (c > 0) codeModule.DeleteLines(1, c);
            if (snapshot.Length > 0) codeModule.AddFromString(snapshot);
        });
        var expectedLines = CanonicalText.TrimTrailingEmpty(CanonicalText.SplitLines(snapshot));
        bool snapshotEndsWithParen = expectedLines.Length > 0 && expectedLines[^1].Trim() == "()";
        if (snapshot.Length > 0 && !snapshotEndsWithParen)
            while (SweepTrailingParenPhantom(codeModule)) { }
        string now = ReadModuleBody(codeModule);
        return PushRecovery.SnapshotMatches(snapshot, now);
    }

    /// <summary>말미 빈 줄을 건너뛴 최하단 비어있지 않은 줄이 <c>()</c>면 물리 삭제하고 true.</summary>
    private static bool SweepTrailingParenPhantom(dynamic codeModule)
    {
        int i = Convert.ToInt32(codeModule.CountOfLines);
        while (i >= 1 && LineTrim(codeModule, i).Length == 0) i--;
        if (i < 1 || LineTrim(codeModule, i) != "()") return false;
        int at = i;
        ComRetry.Invoke(() => { codeModule.DeleteLines(at, 1); }); // 블록 람다 = Action (dynamic 오버로드 모호 방지)
        return true;
    }

    private static string LineTrim(dynamic codeModule, int idx)
        => (Convert.ToString(codeModule.Lines[idx, 1]) ?? "").Trim();

    /// <summary>말미 빈 줄을 제외한 논리행수(src 정규형은 말미 빈 줄이 없어 실질 본문 길이).</summary>
    private static int MeaningfulCount(IReadOnlyList<string> lines)
    {
        int end = lines.Count;
        while (end > 0 && lines[end - 1].Trim().Length == 0) end--;
        return end;
    }

    private static int MeaningfulCountModule(dynamic codeModule)
    {
        int n = Convert.ToInt32(codeModule.CountOfLines);
        while (n >= 1 && LineTrim(codeModule, n).Length == 0) n--;
        return n;
    }

    /// <summary>표준 모듈 신규 생성 + 이름 지정 + 본문 주입. 이름 지정 실패 시 CliError.</summary>
    public static void AddStdModule(dynamic vbProject, string name, IReadOnlyList<string> bodyLines)
    {
        dynamic comp = ComRetry.Invoke(() => vbProject.VBComponents.Add(TypeStdModule));
        ComRetry.Invoke(() => { comp.Name = name; });
        string actual = Convert.ToString(comp.Name) ?? "";
        if (!actual.Equals(name, StringComparison.OrdinalIgnoreCase))
            throw new CliError($"{name}: 신규 모듈 이름 지정 실패 (실제: {actual})");
        ReplaceCode(comp.CodeModule, bodyLines);
    }

    /// <summary>검증형 제거: Remove 후 이름 부재 확인(비동기 지연 방어).</summary>
    public static void VerifiedRemove(dynamic vbProject, dynamic comp, string name)
    {
        ComRetry.Invoke(() => { vbProject.VBComponents.Remove(comp); });
        if (ComponentNames(vbProject).Contains(name))
            throw new CliError($"{name}: Remove 후에도 컴포넌트가 남아 있음(비동기 지연)");
    }
}
