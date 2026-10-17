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

    /// <summary>실행 exe의 semver 문자열 — Core.ExeVersion 위임(관리 스탬프·업데이트 알림과 같은 값).</summary>
    public static string GetVersion() => ExeVersion.Text;

    /// <summary>InformationalVersion의 '+' 뒤 리비전(예: v0.3.0-5-g8acb552-dirty). doctor/bugreport와 공유.</summary>
    public static string GetRevision() => ExeVersion.Revision;

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
