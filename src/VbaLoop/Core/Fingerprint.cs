using System.Security.Cryptography;
using System.Text;

namespace VbaLoop.Core;

/// <summary>
/// push 해시 캐시용 src 소스 지문(설계 §4 push 0단계 "재검증 필요 여부를 도구가 판단").
/// 파일명(소문자)+내용 해시를 정렬 결합해 SHA-256 한 값으로.
/// </summary>
public static class Fingerprint
{
    public static readonly string[] SourceExtensions = { ".bas", ".cls", ".frm", ".frx" };

    public static string Compute(string srcRoot)
    {
        var entries = new List<string>();
        if (Directory.Exists(srcRoot))
        {
            foreach (var file in Directory.EnumerateFiles(srcRoot))
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                if (!SourceExtensions.Contains(ext)) continue;
                var name = Path.GetFileName(file).ToLowerInvariant();
                var hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file)));
                entries.Add($"{name}\0{hash}");
            }
        }
        entries.Sort(StringComparer.Ordinal);
        var combined = string.Join("\n", entries);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(combined)));
    }
}
