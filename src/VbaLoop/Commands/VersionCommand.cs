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

    /// <summary>InformationalVersion의 '+' 앞 semver(예: 1.2.3). 릴리스 태그에서 주입되며,
    /// 로컬 빌드는 csproj 기본값이다 — 사용자가 "내 vbaloop가 몇 번인가"를 묻는 단위.</summary>
    public static string GetVersion()
    {
        var info = Assembly.GetEntryAssembly()
            ?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (info == null) return "unknown";
        var plus = info.IndexOf('+');
        return plus >= 0 ? info[..plus] : info;
    }

    /// <summary>InformationalVersion의 '+' 뒤 리비전(예: 8acb552-dirty). doctor/bugreport와 공유.</summary>
    public static string GetRevision()
    {
        var info = Assembly.GetEntryAssembly()
            ?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (info == null) return "unknown";
        var plus = info.IndexOf('+');
        return plus >= 0 && plus + 1 < info.Length ? info[(plus + 1)..] : info;
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
