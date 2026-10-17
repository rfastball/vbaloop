namespace VbaLoop.Core;

/// <summary>
/// 사용자에게 그대로 보여줄 한 줄 에러 + hint. 잡히면 exit 2로 종결.
/// environment=true는 환경 오류(세션 부재/stale/오염, Excel 기동/COM 치명) — exit 3으로
/// 구분되어 에이전트가 JSON 파싱 없이 "session start 후 재시도"를 분기할 수 있다.
/// </summary>
public sealed class CliError : Exception
{
    public string? Hint { get; }
    public bool Environment { get; }

    public CliError(string message, string? hint = null, bool environment = false) : base(message)
    {
        Hint = hint;
        Environment = environment;
    }
}
