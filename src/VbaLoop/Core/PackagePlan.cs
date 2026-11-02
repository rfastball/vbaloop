namespace VbaLoop.Core;

/// <summary>
/// 패키지 적용의 모듈 계획과 같은 버전 코드 비교(설계 §13.3 5단계, §13.4). 워크북 컴포넌트 목록과 패키지
/// 모듈만 보는 순수 로직이다 — 주입은 명령 계층이 이 계획대로 한다.
/// <list type="bullet">
/// <item>통합 문서 모듈은 이름이 아니라 종류로 맞춘다(CodeName이 UI 언어에 따라 다르다).</item>
/// <item>시트 모듈은 CodeName으로 맞추고, 패키지에 있는데 워크북에 없으면 오류다(만들 수 없다).</item>
/// <item>패키지에 없는 표준·클래스·폼은 제거하고, 워크북에만 있는 시트 모듈은 건드리지 않는다.</item>
/// </list>
/// </summary>
public static class PackagePlan
{
    /// <param name="Kind">null이면 지원하지 않는 컴포넌트 타입 — 계획·비교에서 뺀다.</param>
    /// <param name="IsWorkbook">통합 문서 모듈(Type 100이면서 이름이 Workbook.CodeName).</param>
    public sealed record WorkbookComponent(string Name, VbaComponentKind? Kind, bool IsWorkbook, string Body);

    public enum StepAction
    {
        /// <summary>표준·클래스 모듈 신규 생성.</summary>
        Add,
        /// <summary>표준·클래스 모듈 코드 교체.</summary>
        Replace,
        /// <summary>문서 모듈(시트·통합 문서) 코드 교체.</summary>
        Inject,
        /// <summary>폼 신규 Import.</summary>
        ImportForm,
        /// <summary>기존 폼 Remove 후 Import — 레이아웃을 비교할 근거가 없어 항상 다시 넣는다.</summary>
        ReimportForm,
        /// <summary>패키지에 없는 표준·클래스·폼 제거.</summary>
        Remove,
        /// <summary>워크북에만 있는 시트 모듈 — 데이터 영역이라 건드리지 않는다(경고).</summary>
        IgnoreDocument,
    }

    /// <param name="Target">워크북 쪽 컴포넌트 이름(통합 문서 모듈은 그 워크북의 CodeName).</param>
    /// <param name="Module">패키지 모듈(Remove/IgnoreDocument는 null).</param>
    public sealed record Step(string Target, StepAction Action, PackageFormat.Module? Module,
        IReadOnlyList<string> Header, IReadOnlyList<string> Body);

    public sealed record Result(IReadOnlyList<Step> Steps, IReadOnlyList<string> Errors);

    public static Result Build(IReadOnlyList<WorkbookComponent> workbook, IReadOnlyList<PackageFormat.Module> package)
    {
        var steps = new List<Step>();
        var errors = new List<string>();
        var byName = new Dictionary<string, WorkbookComponent>(StringComparer.OrdinalIgnoreCase);
        foreach (var c in workbook) byName.TryAdd(c.Name, c);
        var workbookModule = workbook.FirstOrDefault(c => c.IsWorkbook);
        var targeted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var m in package)
        {
            // strict: 코드 시작 이후 Attribute는 CodeModule 경로로 왕복 불가 — 주입 전에 거부(CliError)
            var split = VbaHeader.Split(m.Lines.ToArray(), m.FileName, strict: true);
            string target = m.Kind == PackageModuleKind.Workbook ? workbookModule?.Name ?? m.Name : m.Name;
            if (!targeted.Add(target))
            {
                errors.Add($"{m.FileName}: two package modules map to workbook component {target}");
                continue;
            }
            byName.TryGetValue(target, out var comp);
            Step Make(StepAction a) => new(target, a, m, split.Header, split.Body);

            switch (m.Kind)
            {
                case PackageModuleKind.Workbook:
                    if (workbookModule == null) errors.Add($"{m.FileName}: workbook has no ThisWorkbook document module");
                    else steps.Add(Make(StepAction.Inject));
                    break;
                case PackageModuleKind.Document:
                    if (comp == null)
                        errors.Add($"{m.FileName}: sheet module {target} is not in the workbook — sheets cannot be created by a package");
                    else if (comp.Kind != VbaComponentKind.Document || comp.IsWorkbook)
                        errors.Add(Mismatch(m, comp));
                    else steps.Add(Make(StepAction.Inject));
                    break;
                case PackageModuleKind.Form:
                    if (comp == null) steps.Add(Make(StepAction.ImportForm));
                    else if (comp.Kind != VbaComponentKind.MsForm) errors.Add(Mismatch(m, comp));
                    else steps.Add(Make(StepAction.ReimportForm));
                    break;
                default:
                    var want = m.Kind == PackageModuleKind.Std ? VbaComponentKind.StdModule : VbaComponentKind.ClassModule;
                    if (comp == null) steps.Add(Make(StepAction.Add));
                    else if (comp.Kind != want) errors.Add(Mismatch(m, comp));
                    else steps.Add(Make(StepAction.Replace));
                    break;
            }
        }

        foreach (var c in workbook)
        {
            if (targeted.Contains(c.Name) || c.Kind == null) continue;
            if (c.Kind == VbaComponentKind.Document)
                steps.Add(new Step(c.Name, StepAction.IgnoreDocument, null, Array.Empty<string>(), Array.Empty<string>()));
            else
                steps.Add(new Step(c.Name, StepAction.Remove, null, Array.Empty<string>(), Array.Empty<string>()));
        }
        return new Result(steps, errors);
    }

    private static string Mismatch(PackageFormat.Module m, WorkbookComponent comp)
        => $"{m.FileName}: package kind {PackageFormat.KindText(m.Kind)} does not match workbook component {comp.Name} "
           + $"({(comp.IsWorkbook ? "workbook" : comp.Kind?.ToString() ?? "unknown")})";

    /// <summary>
    /// 같은 버전 판정용 "코드가 같다"(설계 §13.4): 컴포넌트 집합이 같고(추가·제거할 것이 없고) 모든 모듈의
    /// 코드가 같다. 폼은 코드비하인드만 본다 — 레이아웃은 Export가 비결정적이라 비교하지 않는다. 식별자
    /// 대소문자만 다른 행은 같다(VBE 프로젝트 전역 재케이싱). 합성 modBuildInfo는 비교하지 않는다.
    /// 반환: 다른 컴포넌트 이름들(비면 같다).
    /// </summary>
    public static IReadOnlyList<string> CodeDifferences(Result plan, IReadOnlyList<WorkbookComponent> workbook)
    {
        var diffs = new List<string>();
        diffs.AddRange(plan.Errors.Count > 0 ? new[] { "(plan errors)" } : Array.Empty<string>());
        var byName = workbook.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase);
        foreach (var s in plan.Steps)
        {
            // 합성 스탬프는 빌드마다 BuiltAt이 달라 코드가 아니다 — 같은 src를 다시 빌드한 패키지를 "다른 코드"로 거부하지 않게
            if (s.Target.Equals(VersionStamp.ModuleName, StringComparison.OrdinalIgnoreCase)) continue;
            switch (s.Action)
            {
                case StepAction.Add or StepAction.ImportForm or StepAction.Remove:
                    diffs.Add(s.Target);
                    break;
                case StepAction.Replace or StepAction.Inject or StepAction.ReimportForm:
                    var wbLines = SplitBody(byName[s.Target].Body);
                    bool form = s.Action == StepAction.ReimportForm;
                    if (!BodyEquivalent(Normalize(wbLines, form), Normalize(s.Body, form)))
                        diffs.Add(s.Target);
                    break;
            }
        }
        return diffs;
    }

    private static IReadOnlyList<string> SplitBody(string body)
        => body.Length > 0 ? CanonicalText.SplitLines(body) : Array.Empty<string>();

    /// <summary>끝 빈 줄 제거. 폼은 선행 빈 줄과 멤버 Attribute도 뺀다(ModuleDiff 폼 비교와 같은 규칙).</summary>
    private static IReadOnlyList<string> Normalize(IReadOnlyList<string> lines, bool form)
    {
        var t = CanonicalText.TrimTrailingEmpty(lines);
        if (!form) return t;
        int lead = 0;
        while (lead < t.Length && t[lead].Trim().Length == 0) lead++;
        return CanonicalText.StripMemberAttributes(t.Skip(lead).ToArray());
    }

    private static bool BodyEquivalent(IReadOnlyList<string> a, IReadOnlyList<string> b)
    {
        if (a.Count != b.Count) return false;
        for (int i = 0; i < a.Count; i++)
            if (!string.Equals(a[i], b[i], StringComparison.Ordinal) && !VbaCasing.LineEquivalent(a[i], b[i]))
                return false;
        return true;
    }
}
