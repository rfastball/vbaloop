namespace VbaLoop.Core;

/// <summary>
/// SaveAs 오탈출(FullName 드리프트) 감지의 순수 판정부.
/// read-only 워크북에서 Excel의 Save가 오류 대신 기본 저장 위치로 조용히 SaveAs할 수 있어,
/// Save 후 FullName이 정본 경로와 같은지(SamePath), stop 시 동명 타경로 워크북이 있는지
/// (FindDriftSuspects)를 판정한다.
/// </summary>
public static class PathDrift
{
    /// <summary>경로 동일성: GetFullPath 정규화 + 끝 구분자 제거 + OrdinalIgnoreCase.</summary>
    public static bool SamePath(string a, string b)
        => string.Equals(Normalize(a), Normalize(b), StringComparison.OrdinalIgnoreCase);

    /// <summary>파일명은 대상과 같은데 경로가 다른 항목(SaveAs 오탈출 의심)을 반환.</summary>
    public static List<string> FindDriftSuspects(IEnumerable<string> fullNames, string targetPath)
    {
        string targetFile = FileNameOf(targetPath);
        return fullNames
            .Where(f => !SamePath(f, targetPath)
                        && string.Equals(FileNameOf(f), targetFile, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static string Normalize(string path)
    {
        // COM FullName은 경로가 아닐 수 있다(클라우드 파일의 https URL 등) — 실패 시 raw 비교로 폴백
        try
        {
            return Path.GetFullPath(path).TrimEnd('\\', '/');
        }
        catch
        {
            return path.TrimEnd('\\', '/');
        }
    }

    private static string FileNameOf(string path)
        => Path.GetFileName(path.TrimEnd('\\', '/')); // .NET Core의 GetFileName은 비경로 문자열에도 안전
}
