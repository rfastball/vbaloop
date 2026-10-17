namespace VbaLoop.Core;

/// <summary>
/// LogShim 배포 무해화(설계 §9.3): LogShim은 어서션 싱크이자 런타임 로깅 파실리티라
/// 배포본에 <b>포함</b>하되(제외 아님 — modEngine 실측상 런타임 의존), 주입할 텍스트에서만
/// <c>IS_RELEASE</c> 상수를 True로 뒤집어 각 싱크를 no-op로 만든다. src·정본은 False 불변
/// (아티팩트는 빌드 산출물). 호출부는 그대로 남아 컴파일 안전, 비용은 상수 분기 1회.
/// </summary>
public static class LogShimRelease
{
    /// <summary>
    /// LogShim 라인들에서 <c>Private Const IS_RELEASE As Boolean = False</c>의 False를
    /// True로 재작성한 사본을 반환. 정본에 상수가 없으면(무해화 계약 위반) CliError —
    /// dist가 조용히 로깅 켜진 배포본을 내지 않도록 크게 실패한다.
    /// </summary>
    public static IReadOnlyList<string> FlipToRelease(IReadOnlyList<string> lines)
    {
        var result = new List<string>(lines.Count);
        bool flipped = false;
        foreach (var line in lines)
        {
            if (!flipped && IsReleaseConst(line, out string replaced))
            {
                result.Add(replaced);
                flipped = true;
            }
            else
            {
                result.Add(line);
            }
        }
        if (!flipped)
            throw new CliError(
                "LogShim has no IS_RELEASE constant — violates the distribution sanitize contract",
                "Add Private Const IS_RELEASE As Boolean = False to assets/LogShim.bas, then re-vendor with rollout.ps1");
        return result;
    }

    /// <summary>
    /// 코드부(주석 제거)에서 <c>[Private] Const IS_RELEASE As Boolean = False</c> 판정.
    /// 공백에 관용적이며 = 우측의 False 토큰만 True로 치환(원 라인의 주석은 보존).
    /// </summary>
    private static bool IsReleaseConst(string line, out string replaced)
    {
        replaced = line;
        var (code, comment) = VbaLine.SplitCodeComment(line);
        var words = code.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        // 최소 형태: Const IS_RELEASE As Boolean = False  (선두 Private 허용)
        int i = 0;
        if (i < words.Length && words[i].Equals("Private", StringComparison.OrdinalIgnoreCase)) i++;
        if (i + 5 >= words.Length) return false;
        if (!words[i].Equals("Const", StringComparison.OrdinalIgnoreCase)) return false;
        if (!words[i + 1].Equals("IS_RELEASE", StringComparison.OrdinalIgnoreCase)) return false;
        if (!words[i + 2].Equals("As", StringComparison.OrdinalIgnoreCase)) return false;
        if (!words[i + 3].Equals("Boolean", StringComparison.OrdinalIgnoreCase)) return false;
        if (!words[i + 4].Equals("=", StringComparison.Ordinal)) return false;
        if (!words[i + 5].Equals("False", StringComparison.OrdinalIgnoreCase)) return false;

        // 코드부의 마지막 False 토큰만 치환(주석부는 건드리지 않음)
        int falseAt = code.LastIndexOf("False", StringComparison.OrdinalIgnoreCase);
        string newCode = code[..falseAt] + "True" + code[(falseAt + "False".Length)..];
        replaced = newCode + (comment ?? "");
        return true;
    }
}
