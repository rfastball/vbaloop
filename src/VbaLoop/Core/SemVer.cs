namespace VbaLoop.Core;

/// <summary>
/// semver 2.0 판독/비교(설계 §12). 관리 스탬프(vbaloop.toml [vbaloop] managed_by)와 업데이트 알림이
/// "exe가 프로젝트보다 오래됐나", "더 새 릴리스가 있나"를 판단하는 단위다.
/// 빌드 메타데이터(+...)는 우선순위에 쓰지 않으므로 파싱 시 버린다. 앞의 v는 허용(태그 이름).
/// </summary>
public sealed class SemVer : IComparable<SemVer>
{
    public int Major { get; }
    public int Minor { get; }
    public int Patch { get; }
    /// <summary>prerelease 식별자(예: rc.1 -> ["rc","1"]). 비어 있으면 정식 릴리스.</summary>
    public IReadOnlyList<string> Prerelease { get; }

    public bool IsPrerelease => Prerelease.Count > 0;

    private SemVer(int major, int minor, int patch, IReadOnlyList<string> pre)
    {
        Major = major; Minor = minor; Patch = patch; Prerelease = pre;
    }

    public static bool TryParse(string? text, out SemVer version)
    {
        version = null!;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var s = text.Trim();
        if (s.StartsWith('v') || s.StartsWith('V')) s = s[1..];
        int plus = s.IndexOf('+');
        if (plus >= 0) s = s[..plus];

        string core = s;
        var pre = Array.Empty<string>();
        int dash = s.IndexOf('-');
        if (dash >= 0)
        {
            core = s[..dash];
            var preText = s[(dash + 1)..];
            if (preText.Length == 0) return false;
            pre = preText.Split('.');
            foreach (var id in pre)
                if (id.Length == 0 || !id.All(c => char.IsAsciiLetterOrDigit(c) || c == '-')) return false;
        }

        var parts = core.Split('.');
        if (parts.Length != 3) return false;
        var nums = new int[3];
        for (int i = 0; i < 3; i++)
        {
            if (parts[i].Length == 0 || !parts[i].All(char.IsAsciiDigit)) return false;
            if (!int.TryParse(parts[i], out nums[i])) return false;
        }
        version = new SemVer(nums[0], nums[1], nums[2], pre);
        return true;
    }

    public static SemVer? ParseOrNull(string? text) => TryParse(text, out var v) ? v : null;

    /// <summary>semver §11 우선순위: 숫자 3개 -> prerelease 없음이 더 큼 -> 식별자별(숫자는 수치,
    /// 영숫자는 ordinal, 숫자 &lt; 영숫자) -> 식별자가 더 많은 쪽이 큼.</summary>
    public int CompareTo(SemVer? other)
    {
        if (other is null) return 1;
        int c = Major.CompareTo(other.Major);
        if (c != 0) return c;
        c = Minor.CompareTo(other.Minor);
        if (c != 0) return c;
        c = Patch.CompareTo(other.Patch);
        if (c != 0) return c;

        if (!IsPrerelease && !other.IsPrerelease) return 0;
        if (!IsPrerelease) return 1;
        if (!other.IsPrerelease) return -1;

        for (int i = 0; i < Math.Min(Prerelease.Count, other.Prerelease.Count); i++)
        {
            var a = Prerelease[i];
            var b = other.Prerelease[i];
            bool an = a.All(char.IsAsciiDigit), bn = b.All(char.IsAsciiDigit);
            if (an && bn)
            {
                c = a.Length != b.Length ? a.Length.CompareTo(b.Length) : string.CompareOrdinal(a, b);
            }
            else if (an) return -1;
            else if (bn) return 1;
            else c = string.CompareOrdinal(a, b);
            if (c != 0) return c < 0 ? -1 : 1;
        }
        return Prerelease.Count.CompareTo(other.Prerelease.Count);
    }

    public override string ToString()
        => IsPrerelease ? $"{Major}.{Minor}.{Patch}-{string.Join('.', Prerelease)}" : $"{Major}.{Minor}.{Patch}";

    public static int Compare(SemVer a, SemVer b) => a.CompareTo(b);
}
