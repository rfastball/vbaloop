using System.Text;

namespace VbaLoop.Core;

/// <summary>
/// 식별자 케이싱 등가관계: VBE 프로젝트 전역 케이스 테이블이 push 주입 시
/// 식별자 철자를 프로젝트에 등록된 철자로 재작성하므로, "식별자 토큰의 케이싱만 다른" 두 라인은
/// 동기화에서 동일로 취급한다. 문자열 리터럴 내용·주석 텍스트는 의미적이라 Ordinal로 구분한다 —
/// 전체-라인 대소문자 무시 비교(구 미봉책)가 이를 혼동해 실변경을 숨기던 결함을 닫는다. 판정은
/// 반드시 <see cref="VbaLine"/> 토크나이저 위에서만(정규식 금지). pull·diff·폼이 공유하는 바닥.
/// </summary>
public static class VbaCasing
{
    /// <summary>
    /// 두 물리 라인이 오직 식별자 토큰의 케이싱만 다르면 true. 주석부는 Ordinal 동일(유무·선행 공백
    /// 포함), 코드부의 토큰 사이 gap(연산자·공백·문자열 리터럴·<c>.</c>·숫자)은 Ordinal 동일,
    /// 식별자 토큰은 OrdinalIgnoreCase 동일이어야 한다. 토큰 수가 다르면 false. REM 주석 라인은
    /// 코드부 전체를 Ordinal 비교한다 — <see cref="VbaLine.SplitCodeComment"/>가 REM을 주석으로
    /// 보지 않아 REM 뒤 텍스트가 코드부에 남으므로, 케이싱을 무시하면 실제 주석 편집을 되돌린다.
    /// </summary>
    public static bool LineEquivalent(string a, string b)
    {
        var (codeA, comA) = VbaLine.SplitCodeComment(a);
        var (codeB, comB) = VbaLine.SplitCodeComment(b);

        if (!string.Equals(comA ?? "", comB ?? "", StringComparison.Ordinal)) return false;

        if (VbaLine.IsRemComment(codeA) || VbaLine.IsRemComment(codeB))
            return string.Equals(codeA, codeB, StringComparison.Ordinal);

        var toksA = VbaLine.Identifiers(codeA);
        var toksB = VbaLine.Identifiers(codeB);
        if (toksA.Count != toksB.Count) return false;

        int prevA = 0, prevB = 0;
        for (int k = 0; k < toksA.Count; k++)
        {
            var ta = toksA[k];
            var tb = toksB[k];
            if (!string.Equals(codeA[prevA..ta.Start], codeB[prevB..tb.Start], StringComparison.Ordinal))
                return false;
            if (!string.Equals(ta.Text, tb.Text, StringComparison.OrdinalIgnoreCase))
                return false;
            prevA = ta.End;
            prevB = tb.End;
        }
        return string.Equals(codeA[prevA..], codeB[prevB..], StringComparison.Ordinal);
    }

    /// <summary>
    /// <see cref="LineEquivalent"/>일 때 워크북 라인 내용에 src 식별자 철자를 splice한 라인 반환 —
    /// 워크북 gap·주석부는 그대로, 토큰 텍스트만 src 철자 채택. 비등가면 workbookLine 그대로(안전장치)
    /// 이므로 실변경(문자열/연산자/숫자)은 절대 되돌아가지 않는다. 같은 라인에 실변경+케이싱이 섞이면
    /// 비등가 → 워크북 그대로, 그 라인은 diff에서 changed로 드러나고 다음 왕복에서 자연 수렴한다.
    /// </summary>
    public static string ReconcileCasing(string workbookLine, string srcLine)
    {
        if (!LineEquivalent(workbookLine, srcLine)) return workbookLine;

        var (wbCode, wbCom) = VbaLine.SplitCodeComment(workbookLine);
        var (srcCode, _) = VbaLine.SplitCodeComment(srcLine);
        if (VbaLine.IsRemComment(wbCode)) return workbookLine; // 재케이싱할 식별자 없음

        var wbToks = VbaLine.Identifiers(wbCode);
        var srcToks = VbaLine.Identifiers(srcCode); // 토큰 수 동일(LineEquivalent 보장)

        var sb = new StringBuilder(workbookLine.Length);
        int prev = 0;
        for (int k = 0; k < wbToks.Count; k++)
        {
            sb.Append(wbCode, prev, wbToks[k].Start - prev); // 워크북 gap 그대로
            sb.Append(srcToks[k].Text);                       // src 철자
            prev = wbToks[k].End;
        }
        sb.Append(wbCode, prev, wbCode.Length - prev);        // trailing gap
        if (wbCom != null) sb.Append(wbCom);                  // 워크북 주석부 그대로
        return sb.ToString();
    }

    /// <summary>
    /// 라인 삽입/삭제에 견고한 케이싱 수렴: <see cref="LineEquivalent"/>를 동등 술어로 하는 LCS 정렬
    /// 후 매칭 쌍만 <see cref="ReconcileCasing"/>, 매칭 안 된 워크북 라인(삽입분)은 그대로 통과.
    /// 한쪽이 비면 workbook 그대로. 인덱스 정렬(구 <see cref="CanonicalText.PreserveCasing"/>)과 달리
    /// 앞선 삽입/삭제가 뒤 라인의 케이싱 수렴을 깨지 않는다.
    /// </summary>
    public static IReadOnlyList<string> ReconcileBlock(
        IReadOnlyList<string> workbook, IReadOnlyList<string> src)
    {
        if (workbook.Count == 0 || src.Count == 0) return workbook;

        var result = workbook.ToArray();
        foreach (var (wi, si) in AlignEquivalent(workbook, src))
            result[wi] = ReconcileCasing(workbook[wi], src[si]);
        return result;
    }

    /// <summary>
    /// <see cref="LineEquivalent"/> 기준 LCS 정렬 — 증가하는 (workbookIndex, srcIndex) 매칭 쌍.
    /// 공통 등가 prefix/suffix를 먼저 벗겨(흔한 경우 = 소수 라인만 재케이싱 → DP 박스 대폭 축소)
    /// 가운데 밴드만 O(n·m) DP. 모듈은 수백 라인이라 실측상 무시 가능.
    /// </summary>
    private static List<(int Wi, int Si)> AlignEquivalent(
        IReadOnlyList<string> a, IReadOnlyList<string> b)
    {
        int n = a.Count, m = b.Count;
        var pairs = new List<(int, int)>();

        int lo = 0;
        while (lo < n && lo < m && LineEquivalent(a[lo], b[lo]))
        {
            pairs.Add((lo, lo));
            lo++;
        }

        int hiA = n - 1, hiB = m - 1;
        var suffix = new List<(int, int)>();
        while (hiA >= lo && hiB >= lo && LineEquivalent(a[hiA], b[hiB]))
        {
            suffix.Add((hiA, hiB));
            hiA--;
            hiB--;
        }

        int rows = hiA - lo + 1, cols = hiB - lo + 1;
        if (rows > 0 && cols > 0)
        {
            var eq = new bool[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    eq[i, j] = LineEquivalent(a[lo + i], b[lo + j]);

            var dp = new int[rows + 1, cols + 1];
            for (int i = rows - 1; i >= 0; i--)
                for (int j = cols - 1; j >= 0; j--)
                    dp[i, j] = eq[i, j] ? dp[i + 1, j + 1] + 1 : Math.Max(dp[i + 1, j], dp[i, j + 1]);

            int x = 0, y = 0;
            while (x < rows && y < cols)
            {
                if (eq[x, y]) { pairs.Add((lo + x, lo + y)); x++; y++; }
                else if (dp[x + 1, y] >= dp[x, y + 1]) x++;
                else y++;
            }
        }

        suffix.Reverse();
        pairs.AddRange(suffix);
        return pairs;
    }
}
