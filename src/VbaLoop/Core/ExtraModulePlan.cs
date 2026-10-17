namespace VbaLoop.Core;

/// <summary>
/// push의 미추적 워크북 모듈 분류(설계 §4). src에 없는
/// 워크북 모듈을 프루너블(표준/클래스/폼)과 무시할 문서모듈로 나눈다.
///
/// 문서/시트 모듈(vbext_ct_Document)은 VBComponents.Add/Remove로 생성·삭제할 수 없어
/// 워크북에만 존재하는 것이 정상 상태다. 매크로가 Worksheets.Add로 시트를 만들면 Sheet2~
/// 코드모듈이 자동 생성되는데, 이것이 무관한 표준모듈 push를 프루닝 가드로 막으면 안 된다.
/// 따라서 문서모듈은 차단/프루닝에서 제외하고 무시(IgnoredDocuments)로 보고만 한다.
/// </summary>
public static class ExtraModulePlan
{
    // VBA 컴포넌트 타입(vbext_ComponentType) — Excel COM 비참조 유지를 위해 값만 미러.
    private const int TypeStdModule = 1;
    private const int TypeClassModule = 2;
    private const int TypeMsForm = 3;
    private const int TypeDocument = 100;

    public sealed record Result(
        IReadOnlyList<string> Prunable,
        IReadOnlyList<string> IgnoredDocuments);

    /// <summary>워크북 컴포넌트 (이름, 타입) 목록과 src 이름 집합을 받아, src에 없는 것을
    /// 프루너블/무시-문서로 분류. src에 있는 것·관리 대상 아닌 타입은 결과에서 제외.</summary>
    public static Result Classify(
        IEnumerable<(string Name, int Type)> components,
        ISet<string> srcNames)
    {
        var prunable = new List<string>();
        var ignoredDocs = new List<string>();
        foreach (var (name, type) in components)
        {
            if (srcNames.Contains(name)) continue;
            switch (type)
            {
                case TypeDocument:
                    ignoredDocs.Add(name);
                    break;
                case TypeStdModule:
                case TypeClassModule:
                case TypeMsForm:
                    prunable.Add(name);
                    break;
                // 그 외 타입은 push 관리 대상 아님 — 무시
            }
        }
        return new Result(prunable, ignoredDocs);
    }
}
