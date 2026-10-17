using System.Reflection;
using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Commands;

/// <summary>
/// version: 실행 exe의 git 리비전 + 빌드 시각. "이 exe가 stale인가"를 한 줄로 답한다.
/// 리비전은 csproj의 StampGitRevision
/// 타깃이 SourceRevisionId로 주입한 git describe 결과 — dirty 접미는 미커밋 소스 빌드 표식.
/// </summary>
public static class VersionCommand
{
    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "version" };
        var exePath = Environment.ProcessPath;
        o.HumanLines.Add($"vbaloop {GetVersion()} ({GetRevision()})");
        o.HumanLines.Add($"exe: {exePath ?? "미상"}");
        o.HumanLines.Add($"built: {GetBuildTime() ?? "미상"}");
        o.Payload = new
        {
            version = GetVersion(),
            revision = GetRevision(),
            exePath,
            builtAt = GetBuildTime(),
        };
        return o;
    }

    /// <summary>InformationalVersion의 '+' 앞 semver(예: 1.2.3). 릴리스는 태그에서 주입되고,
    /// 로컬 빌드는 마지막 v* 태그에서 유도된다(csproj StampGitRevision) — 사용자가 "내 vbaloop가 몇 번인가"를
    /// 묻는 단위이자 관리 스탬프·업데이트 알림의 비교 단위.</summary>
    public static string GetVersion() => SplitInformational(ReadInformational()).Version;

    /// <summary>InformationalVersion의 '+' 뒤 리비전(예: v0.3.0-5-g8acb552-dirty). doctor/bugreport와 공유.</summary>
    public static string GetRevision() => SplitInformational(ReadInformational()).Revision;

    /// <summary>실행 exe 버전의 semver 판독. 판독 불가(unknown 등)면 null.</summary>
    public static SemVer? GetSemVer() => SemVer.ParseOrNull(GetVersion());

    private static string? ReadInformational() => Assembly.GetEntryAssembly()
        ?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;

    /// <summary>"0.3.0+v0.3.0-5-gabc-dirty" -> ("0.3.0", "v0.3.0-5-gabc-dirty"). '+'가 없으면 둘 다 전체 문자열
    /// (기존 동작 유지), null이면 unknown.</summary>
    public static (string Version, string Revision) SplitInformational(string? info)
    {
        if (info == null) return ("unknown", "unknown");
        var plus = info.IndexOf('+');
        if (plus < 0) return (info, info);
        return (info[..plus], plus + 1 < info.Length ? info[(plus + 1)..] : info);
    }

    /// <summary>exe 파일 mtime = publish 시각 (PublishSingleFile에서 신뢰 가능).</summary>
    public static string? GetBuildTime()
    {
        var exePath = Environment.ProcessPath;
        if (exePath == null || !File.Exists(exePath)) return null;
        return File.GetLastWriteTime(exePath).ToString("yyyy-MM-dd HH:mm:ss");
    }

    /// <summary>bugreport 템플릿의 "실행 exe 빌드" 라인용 요약.</summary>
    public static string GetBuildInfo()
        => $"{GetVersion()} ({GetRevision()}), built {GetBuildTime() ?? "미상"}";
}
