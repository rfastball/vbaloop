namespace VbaLoop.Core;

/// <summary>
/// init이 정본을 소유하는 벤더링 VBA 모듈(현재 LogShim). 진실은 워크북이 아니라 init
/// 임베디드 자산이므로(하드 불변식: 직접 수정 금지, 갱신은 vbaloop init) pull은 이들의
/// src 파일을 워크북 본문으로 덮어쓰지 않고, diff는 워크북↔src 차이를 changed로 보고하지
/// 않는다. src 정본 무결성은 init --check가 담당한다.
/// push는 예외 — src 정본을 워크북에 주입하는 것이 정상 경로다.
/// </summary>
public static class VendoredModules
{
    public static readonly IReadOnlySet<string> Names =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "LogShim" };

    public static bool IsVendored(string moduleName) => Names.Contains(moduleName);
}
