using System.Text;

namespace VbaLoop.Core;

/// <summary>
/// CP949 인코딩 가능성 사전 검사. VBE 내부 저장은 PROJECTCODEPAGE(이 환경 CP949)
/// MBCS이므로 CP949 밖 문자는 push 시 확정 손실(2026-07-08 실측: U+2014 → 0x3F).
/// 손실 전에 파일/라인/컬럼 단위로 탐지해 save 전에 실패시킨다.
/// </summary>
public static class Cp949Guard
{
    public sealed record Violation(int Line, int Column, string Char, string CodePoint);

    /// <summary>Program 기동 시 CodePagesEncodingProvider 등록 후 사용 가능.</summary>
    public static Encoding GetStrict(int codePage) =>
        Encoding.GetEncoding(codePage,
            EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

    public static List<Violation> Check(IReadOnlyList<string> lines, int codePage = 949)
    {
        var enc = GetStrict(codePage);
        var violations = new List<Violation>();

        for (int i = 0; i < lines.Count; i++)
        {
            string line = lines[i];
            bool ok;
            try
            {
                enc.GetByteCount(line);
                ok = true;
            }
            catch (EncoderFallbackException)
            {
                ok = false;
            }
            if (ok) continue;

            // 라인 전체가 불합격이면 문자 단위로 위치 특정 (서로게이트 쌍 단위)
            int col = 0;
            var e = System.Globalization.StringInfo.GetTextElementEnumerator(line);
            while (e.MoveNext())
            {
                string el = (string)e.Current;
                col = e.ElementIndex + 1;
                try
                {
                    enc.GetByteCount(el);
                }
                catch (EncoderFallbackException)
                {
                    int cp = char.ConvertToUtf32(el, 0);
                    violations.Add(new Violation(i + 1, col, el, $"U+{cp:X4}"));
                }
            }
        }

        return violations;
    }
}
