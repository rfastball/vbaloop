namespace VbaLoop.Core;

/// <summary>
/// 주석제거(설계 §9.5, opt-in): 파일 축소 + `@desc` 내부메모 은닉. <b>정규식 금지</b> —
/// `'`가 문자열 리터럴 안(`x = "it's"`)에 있을 수 있어 <see cref="VbaLine"/> 토크나이저로만
/// 판정한다. 식별자 난독화는 기각(문자열 결합 규약이 조용히 깨짐, §9.5) — 여기선 안 한다.
/// </summary>
public static class VbaCommentStripper
{
    /// <summary>
    /// 본문 라인에서 주석 제거: (a) 온전한 주석 라인(`'`/`REM`)은 삭제, (b) 인라인 주석은
    /// 코드부만 남기고 우측 공백 트림. 문자열 리터럴 내부의 `'`는 절대 건드리지 않는다.
    /// 헤더(Attribute 블록)에는 적용하지 않는다 — 호출부가 본문만 넘긴다.
    /// </summary>
    public static IReadOnlyList<string> Strip(IReadOnlyList<string> body)
    {
        var result = new List<string>(body.Count);
        foreach (var line in body)
        {
            var (code, comment) = VbaLine.SplitCodeComment(line);

            // 온전한 주석 라인(코드부가 공백뿐이고 주석이 있음) → 삭제
            if (comment != null && code.Trim().Length == 0) continue;
            // REM 주석 라인 → 삭제
            if (comment == null && VbaLine.IsRemComment(code)) continue;

            if (comment != null)
                result.Add(code.TrimEnd()); // 인라인 주석 제거 + 꼬리 공백 정리
            else
                result.Add(line);
        }
        return result;
    }
}
