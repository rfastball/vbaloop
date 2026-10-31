namespace VbaLoop.Core;

/// <summary>
/// VBE가 컴파일 오류 모달 해제 직후 남기는 선택 영역(CodePane.GetSelection) — 모달 본문에는 위치가
/// 없으므로 이것이 유일한 위치 신호다(2026-10-31 실측). 구문 오류는 행 전체, 미정의 Sub/Function 같은
/// 의미 오류는 토큰만 선택된다. 열은 1부터, EndColumn은 반열림.
/// CodeLine은 코드 모듈 기준(Attribute 헤더 제외) — src 파일 행은 <see cref="MapToFileLine"/>.
/// </summary>
public sealed record CompileErrorLocation(
    string Module, int CodeLine, int StartColumn, int EndLine, int EndColumn, string LineText)
{
    /// <summary>한 행 안의 부분 선택이면 그 텍스트, 행 전체 선택이나 여러 행이면 null.</summary>
    public string? Token
    {
        get
        {
            if (EndLine != CodeLine || StartColumn < 1 || EndColumn <= StartColumn) return null;
            if (StartColumn > LineText.Length) return null;
            int len = Math.Min(EndColumn - StartColumn, LineText.Length - StartColumn + 1);
            bool wholeLine = LineText[..(StartColumn - 1)].Trim().Length == 0
                             && LineText[(StartColumn - 1 + len)..].Trim().Length == 0;
            if (wholeLine) return null;
            string token = LineText.Substring(StartColumn - 1, len).Trim();
            return token.Length > 0 ? token : null;
        }
    }

    /// <summary>
    /// 코드 모듈 행을 src 파일 행으로. 주입 본문(body)의 같은 위치 행이 VBE 행 텍스트와 같으면 헤더 행 수를
    /// 더한다. 어긋나면(새 모듈에 VBE가 Option Explicit을 자동 삽입한 경우 등) 본문에서 유일하게 일치하는 행을
    /// 찾고, 없거나 여럿이면 null — 틀린 행을 말하느니 코드 모듈 행만 보고한다.
    /// </summary>
    public int? MapToFileLine(int headerLineCount, IReadOnlyList<string> body)
    {
        string want = LineText.TrimEnd();
        int i = CodeLine - 1;
        if (i >= 0 && i < body.Count && body[i].TrimEnd() == want) return headerLineCount + CodeLine;
        int found = -1;
        for (int k = 0; k < body.Count; k++)
        {
            if (body[k].TrimEnd() != want) continue;
            if (found >= 0) return null;
            found = k;
        }
        return found >= 0 ? headerLineCount + found + 1 : null;
    }

    /// <summary>JSON payload 모양(push/dist 공용). file/line은 src로 사상된 경우만.</summary>
    public object ToPayload(string? file = null, int? fileLine = null) => new
    {
        module = Module, file, line = fileLine, codeLine = CodeLine,
        column = StartColumn, endLine = EndLine, endColumn = EndColumn,
        token = Token, text = LineText.Trim(),
    };

    /// <summary>사람/에이전트용 한 줄. file이 있으면 "PerfSynth.bas:16", 없으면 "PerfSynth (code line 15)".</summary>
    public string Describe(string? file = null, int? fileLine = null)
    {
        string where = file != null && fileLine != null ? $"{file}:{fileLine}"
            : file != null ? $"{file} (code line {CodeLine})"
            : $"{Module} (code line {CodeLine})";
        string text = LineText.Trim();
        return Token is { } t ? $"{where} '{t}' in: {text}" : $"{where}: {text}";
    }
}
