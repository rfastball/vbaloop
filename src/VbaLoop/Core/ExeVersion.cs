using System.Reflection;

namespace VbaLoop.Core;

/// <summary>
/// 실행 exe(VbaLoop 어셈블리)의 버전. InformationalVersion = "&lt;semver&gt;+&lt;git describe&gt;" —
/// 릴리스는 태그에서 -p:Version 주입, 로컬 빌드는 마지막 v* 태그에서 유도(csproj StampGitRevision).
/// 관리 스탬프(§12)와 업데이트 알림이 비교하는 단위다. 엔트리 어셈블리가 아니라 이 어셈블리를 읽는다 —
/// 테스트 호스트에서도 vbaloop 자신의 버전이 나와야 하기 때문.
/// </summary>
public static class ExeVersion
{
    private static readonly (string Version, string Revision) Parts = SplitInformational(
        typeof(ExeVersion).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion);

    /// <summary>semver 문자열(예: 0.3.0). 판독 불가면 "unknown".</summary>
    public static string Text => Parts.Version;

    /// <summary>git describe 리비전(예: v0.3.0-5-gabc1234-dirty, 또는 nogit).</summary>
    public static string Revision => Parts.Revision;

    /// <summary>semver 판독. 판독 불가면 null.</summary>
    public static SemVer? Current => SemVer.ParseOrNull(Parts.Version);

    /// <summary>"0.3.0+v0.3.0-5-gabc-dirty" -> ("0.3.0", "v0.3.0-5-gabc-dirty"). '+'가 없으면 둘 다 전체 문자열
    /// (기존 동작 유지), null이면 unknown.</summary>
    public static (string Version, string Revision) SplitInformational(string? info)
    {
        if (info == null) return ("unknown", "unknown");
        var plus = info.IndexOf('+');
        if (plus < 0) return (info, info);
        return (info[..plus], plus + 1 < info.Length ? info[(plus + 1)..] : info);
    }
}
