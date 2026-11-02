using System.Text;
using Microsoft.Win32;

namespace VbaLoop.Core;

/// <summary>
/// Excel 복원력 "사용할 수 없는 항목"(<c>Resiliency\DisabledItems</c>) 중 문서 항목 판독(읽기 전용).
/// 여기 오른 경로는 사람이 Excel에서 열 때 "마지막으로 'X'을(를) 열었을 때 오류가 발생했습니다"를 띄운다.
/// 도구의 숨은 열기는 이 값과 별개의 기록으로 막히며 <c>Excel.WorkbookOpener</c>가 처리한다 — 값을
/// 지워도 그 기록은 남았다(실측 2026-11-02). 그래서 doctor는 보고만 하고 고치지 않는다.
/// </summary>
public static class ExcelDisabledItems
{
    private const int DocumentType = 2;

    /// <summary>값 형식(실측): DWORD 종류(2=문서), DWORD 경로 바이트 수(널 포함), DWORD, UTF-16 소문자 경로.
    /// 문서 항목이 아니거나 길이가 맞지 않으면 null — 순수 함수라 xUnit으로 검증한다.</summary>
    public static string? DecodeDocumentPath(byte[] value)
    {
        if (value.Length < 12 || BitConverter.ToInt32(value, 0) != DocumentType) return null;
        int len = BitConverter.ToInt32(value, 4);
        if (len < 2 || len % 2 != 0 || 12 + len > value.Length) return null;
        return Encoding.Unicode.GetString(value, 12, len).TrimEnd('\0');
    }

    /// <summary>workbookPath와 같은 경로의 항목 값 이름들(대소문자 무시). 키가 없으면 빈 목록.</summary>
    public static List<string> FindEntries(string officeVersion, string workbookPath)
    {
        var hits = new List<string>();
        using var key = Registry.CurrentUser.OpenSubKey(
            $@"Software\Microsoft\Office\{officeVersion}\Excel\Resiliency\DisabledItems");
        if (key == null) return hits;
        string target = Path.GetFullPath(workbookPath);
        foreach (string name in key.GetValueNames())
            if (key.GetValue(name) is byte[] v && DecodeDocumentPath(v) is { } p
                && string.Equals(p, target, StringComparison.OrdinalIgnoreCase))
                hits.Add(name);
        return hits;
    }
}
