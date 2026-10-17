using System.Diagnostics;
using System.Reflection;
using VbaLoop.Commands;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class SemVerTests
{
    [Theory]
    [InlineData("1.2.3", 1, 2, 3, "")]
    [InlineData("v0.3.0", 0, 3, 0, "")]
    [InlineData("0.1.0-rc.1", 0, 1, 0, "rc.1")]
    [InlineData("1.0.0+abc", 1, 0, 0, "")]               // 빌드 메타데이터는 버린다
    [InlineData("1.0.0-beta.2+v1.0.0-3-gabc-dirty", 1, 0, 0, "beta.2")]
    public void Parse_Valid(string text, int major, int minor, int patch, string pre)
    {
        Assert.True(SemVer.TryParse(text, out var v));
        Assert.Equal((major, minor, patch), (v.Major, v.Minor, v.Patch));
        Assert.Equal(pre, string.Join('.', v.Prerelease));
        Assert.Equal(pre.Length > 0, v.IsPrerelease);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("unknown")]
    [InlineData("1.2")]
    [InlineData("1.2.3.4")]
    [InlineData("1.2.x")]
    [InlineData("1.2.3-")]
    [InlineData("1.2.3-rc..1")]
    [InlineData("1.2.3-rc_1")]
    public void Parse_Invalid(string? text)
        => Assert.False(SemVer.TryParse(text, out _));

    [Theory]
    [InlineData("1.0.0", "1.0.0", 0)]
    [InlineData("1.0.0", "1.0.1", -1)]
    [InlineData("1.10.0", "1.9.0", 1)]
    [InlineData("2.0.0", "1.99.99", 1)]
    [InlineData("1.0.0-rc.1", "1.0.0", -1)]              // prerelease < 정식
    [InlineData("1.0.0-alpha", "1.0.0-alpha.1", -1)]     // 식별자가 더 많은 쪽이 큼
    [InlineData("1.0.0-alpha.1", "1.0.0-alpha.beta", -1)] // 숫자 < 영숫자
    [InlineData("1.0.0-beta.2", "1.0.0-beta.11", -1)]    // 숫자 식별자는 수치 비교
    [InlineData("1.0.0-rc.1", "1.0.0-beta.11", 1)]
    [InlineData("0.9.0", "1.0.0-rc.1", -1)]
    [InlineData("v0.3.0", "0.3.0+v0.3.0-5-gabc-dirty", 0)] // 태그+커밋 빌드는 태그와 같다
    public void Compare_FollowsSemverPrecedence(string a, string b, int expected)
    {
        var va = SemVer.ParseOrNull(a)!;
        var vb = SemVer.ParseOrNull(b)!;
        Assert.Equal(expected, Math.Sign(va.CompareTo(vb)));
        Assert.Equal(-expected, Math.Sign(vb.CompareTo(va)));
    }

    [Fact]
    public void ToString_RoundTrips()
    {
        Assert.Equal("0.1.0-rc.1", SemVer.ParseOrNull("v0.1.0-rc.1+x")!.ToString());
        Assert.Equal("1.2.3", SemVer.ParseOrNull("1.2.3")!.ToString());
    }

    [Theory]
    [InlineData("0.3.0+v0.3.0-5-gabc1234-dirty", "0.3.0", "v0.3.0-5-gabc1234-dirty")]
    [InlineData("1.2.3+8acb552", "1.2.3", "8acb552")]
    [InlineData("0.1.0", "0.1.0", "0.1.0")]
    [InlineData(null, "unknown", "unknown")]
    public void SplitInformational_SeparatesVersionAndRevision(string? info, string version, string revision)
        => Assert.Equal((version, revision), ExeVersion.SplitInformational(info));

    /// <summary>
    /// 로컬 빌드 버전 유도(csproj StampGitRevision)의 결과 검증: 이 테스트가 참조하는 VbaLoop 어셈블리는
    /// -p:Version 없이 빌드되므로 버전이 마지막 v* 태그와 같아야 한다. git/태그가 없으면 semver 판독만 본다.
    /// </summary>
    [Fact]
    public void LocalBuild_VersionDerivedFromLastVTag()
    {
        var info = typeof(VersionCommand).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()!.InformationalVersion;
        var (version, revision) = ExeVersion.SplitInformational(info);
        Assert.NotNull(SemVer.ParseOrNull(version));

        var tag = Git("describe --tags --match v* --abbrev=0");
        if (tag == null || SemVer.ParseOrNull(tag) == null) return; // nogit 폴백(0.1.0) 경로
        Assert.Equal(SemVer.ParseOrNull(tag)!.ToString(), version);
        Assert.StartsWith(tag, revision); // 리비전은 describe 원문(태그-커밋수-g해시[-dirty])
    }

    private static string? Git(string args)
    {
        try
        {
            var psi = new ProcessStartInfo("git", args)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                WorkingDirectory = AppContext.BaseDirectory,
            };
            using var p = Process.Start(psi)!;
            var output = p.StandardOutput.ReadToEnd().Trim();
            p.WaitForExit(10000);
            return p.ExitCode == 0 && output.Length > 0 ? output : null;
        }
        catch
        {
            return null;
        }
    }
}
