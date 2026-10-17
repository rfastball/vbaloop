namespace VbaLoop.Core;

/// <summary>
/// run 대상 매크로 이름 검증 — 항상 모듈 한정(Module.Proc)만 허용(비한정 호출은
/// 동명 프로시저 존재 시 1004 모호). 워크북 한정('워크북'!)은 러너가 자동으로 붙인다.
/// </summary>
public static class MacroName
{
    public static bool TryParse(string? input, out string module, out string proc, out string? error)
    {
        module = "";
        proc = "";
        error = null;

        string s = (input ?? "").Trim();
        if (s.Length == 0)
        {
            error = "macro name is empty";
            return false;
        }
        if (s.IndexOfAny(new[] { '!', '\'', '"', ' ', '\t', '[', ']' }) >= 0)
        {
            error = $"contains disallowed characters: {s} — workbook qualification ('...'!) is added automatically";
            return false;
        }
        var parts = s.Split('.');
        if (parts.Length != 2 || parts[0].Length == 0 || parts[1].Length == 0)
        {
            error = $"not in Module.Proc form: {s} — module-qualified call required";
            return false;
        }
        module = parts[0];
        proc = parts[1];
        return true;
    }
}
