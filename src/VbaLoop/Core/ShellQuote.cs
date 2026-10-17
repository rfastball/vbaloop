using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>
/// 에이전트에게 돌려주는 명령 문자열의 PowerShell 인용(설계 §10.8). 큰따옴표는 PowerShell에서 <c>$(...)</c>·
/// <c>$var</c>·백틱을 해석하므로, 경로에 그런 문자가 있으면 실행 시 식이 평가되거나 대상이 바뀐다. 작은따옴표
/// 문자열은 해석이 없고 따옴표만 겹쳐 쓰면 된다. PowerShell은 둥근 작은따옴표(U+2018~U+201B)도 작은따옴표로
/// 취급하므로 함께 겹친다.
/// </summary>
public static partial class ShellQuote
{
    [GeneratedRegex(@"^[A-Za-z0-9_\-.:\\/=]+$")]
    private static partial Regex SafeBare();

    /// <summary>그대로 둬도 안전한 인자는 그대로, 아니면 작은따옴표로 감싼다.</summary>
    public static string PowerShell(string arg)
    {
        if (arg.Length > 0 && SafeBare().IsMatch(arg)) return arg;
        var sb = new System.Text.StringBuilder("'");
        foreach (char c in arg)
        {
            sb.Append(c);
            if (c is '\'' or '‘' or '’' or '‚' or '‛') sb.Append(c);
        }
        return sb.Append('\'').ToString();
    }

    /// <summary>명령과 인자들을 PowerShell에서 그대로 실행할 수 있는 한 줄로.</summary>
    public static string PowerShellCommand(string exe, IEnumerable<string> args)
        => string.Join(" ", new[] { exe }.Concat(args.Select(PowerShell)));
}
