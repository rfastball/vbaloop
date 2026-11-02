using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace VbaLoop.Core;

/// <summary>
/// 패키지 적용의 사용자 단위 저장소(설계 §13.5): <c>%LOCALAPPDATA%\vbaloop\backups\&lt;파일 이름&gt;-&lt;경로 해시 8자&gt;\</c>.
/// 업무 파일 옆에는 아무것도 쓰지 않는다 — 공유 드라이브에 백업이 쌓이거나 쓰기 권한이 없어 적용이 막히지 않게.
/// 경로 해시는 이름이 같은 다른 폴더의 파일을 구분한다. 같은 폴더에 백업 슬롯과 적용 기록(apply.jsonl)이 모인다.
/// </summary>
public static class PackageApplyStore
{
    public const string ApplyLogFileName = "apply.jsonl";

    public static string DefaultRoot => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop", "backups");

    /// <summary>업무 파일 하나의 폴더. 경로는 대소문자·끝 구분자를 무시하고 같은 파일로 본다.</summary>
    public static string FolderFor(string root, string workbookFullPath)
        => Path.Combine(root, Path.GetFileName(workbookFullPath) + "-" + PathHash(workbookFullPath)[..8]);

    public static string PathHash(string fullPath)
    {
        string norm = Path.GetFullPath(fullPath).TrimEnd('\\', '/').ToLowerInvariant();
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(norm))).ToLowerInvariant();
    }

    private static readonly JsonSerializerOptions Json = new()
    {
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>적용 기록 한 줄을 덧붙인다(최선 노력 — 기록 실패가 적용 결과를 바꾸지 않는다).</summary>
    public static bool TryAppend(string folder, object record)
    {
        try
        {
            Directory.CreateDirectory(folder);
            using var fs = new FileStream(Path.Combine(folder, ApplyLogFileName), FileMode.Append, FileAccess.Write, FileShare.Read);
            var bytes = CanonicalText.Utf8NoBom.GetBytes(JsonSerializer.Serialize(record, Json) + "\n");
            fs.Write(bytes, 0, bytes.Length);
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    /// <summary>
    /// Excel 잠금 파일(<c>~$&lt;파일 이름&gt;</c>)에서 연 사람 이름을 읽는다. 첫 바이트가 길이, 이어서 ANSI 이름이고,
    /// 0x36부터 UTF-16 이름이 다시 나온다(첫 바이트 길이 = UTF-16 글자 수). 형식이 안 맞으면 null.
    /// </summary>
    public static string? ParseLockOwner(byte[] bytes, Encoding ansi)
    {
        if (bytes.Length < 2) return null;
        int n = bytes[0];
        if (n == 0 || n > 54) return null;
        const int Utf16At = 0x36;
        if (bytes.Length >= Utf16At + 2 + n * 2)
        {
            int at = Utf16At + 2; // UTF-16 앞에도 길이 워드가 있다
            string wide = Encoding.Unicode.GetString(bytes, at, n * 2).Trim('\0', ' ');
            if (wide.Length > 0) return wide;
        }
        if (bytes.Length < 1 + n) return null;
        string narrow = ansi.GetString(bytes, 1, n).Trim('\0', ' ');
        return narrow.Length > 0 ? narrow : null;
    }
}
