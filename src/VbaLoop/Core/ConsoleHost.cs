using VbaLoop.Excel;

namespace VbaLoop.Core;

/// <summary>
/// 사람 대면 표면(setup 마법사)의 콘솔 입출력. 에이전트가 부르는 CLI 표면은 전부 비대화형
/// <c>--json</c>이므로 이 클래스는 setup에서만 쓴다 — 다른 명령에서 프롬프트를 띄우면
/// 에이전트가 무한 대기에 걸린다.
/// </summary>
public static class ConsoleHost
{
    /// <summary>
    /// 콘솔을 이 프로세스가 단독 소유하는가 = 탐색기에서 더블클릭해 새 콘솔이 할당된 경우.
    /// 셸에서 실행하면 셸 프로세스도 같은 콘솔에 붙어 있어 2 이상이다.
    /// </summary>
    public static bool IsOwnConsole()
    {
        try
        {
            var buf = new uint[4];
            uint n = NativeMethods.GetConsoleProcessList(buf, (uint)buf.Length);
            return n == 1;
        }
        catch
        {
            return false; // 콘솔 API 불가 환경 — 더블클릭으로 단정하지 않는다
        }
    }

    /// <summary>입력을 받을 수 있는 상태인가(리다이렉트/파이프면 false — 프롬프트 금지).</summary>
    public static bool CanPrompt => !Console.IsInputRedirected;

    /// <summary>
    /// 예/아니오 확인. 프롬프트가 불가능하면 <paramref name="fallback"/>을 그대로 반환한다
    /// (비대화형 실행에서 조용히 멈추지 않게 하는 계약).
    /// </summary>
    public static bool Confirm(string question, bool defaultYes, bool fallback)
    {
        if (!CanPrompt) return fallback;
        string suffix = defaultYes ? "[Y/n]" : "[y/N]";
        Console.Write($"{question} {suffix} ");
        string? line = Console.ReadLine();
        if (line == null) return fallback;
        line = line.Trim();
        if (line.Length == 0) return defaultYes;
        return line[0] is 'y' or 'Y';
    }

    /// <summary>한 줄 입력. 빈 줄/입력 불가면 null.</summary>
    public static string? Ask(string question)
    {
        if (!CanPrompt) return null;
        Console.Write($"{question} ");
        string? line = Console.ReadLine()?.Trim();
        return string.IsNullOrEmpty(line) ? null : line;
    }

    /// <summary>더블클릭 실행이면 콘솔 창이 즉시 닫히므로 결과를 읽을 시간을 준다.</summary>
    public static void PauseIfOwnConsole()
    {
        if (!IsOwnConsole() || !CanPrompt) return;
        Console.WriteLine();
        Console.Write("Press Enter to continue...");
        try { Console.ReadLine(); } catch { /* 무시 */ }
    }
}
