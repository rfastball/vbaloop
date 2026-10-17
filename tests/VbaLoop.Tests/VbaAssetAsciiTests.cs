using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>
/// vbaloop가 워크북에 주입하는 VBA 텍스트는 전부 순수 ASCII여야 한다. push는 시스템 ACP
/// (CP1252/CP932/CP936/CP949 ...) 밖 문자를 save 전에 거부하므로, 주석 한 줄의 한글만으로도
/// 비한국어 Windows에서 setup 마법사의 첫 push가 실패한다. ASCII는 모든 ACP의 공통 부분집합이다.
/// </summary>
public class VbaAssetAsciiTests
{
    public VbaAssetAsciiTests() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public static IEnumerable<object[]> InjectedVbaTexts()
    {
        yield return new object[] { "LogShim.bas (embedded)", ReadEmbedded("VbaLoop.assets.LogShim.bas") };
        yield return new object[] { "LogShim.bas (release flip)",
            string.Join("\r\n", LogShimRelease.FlipToRelease(
                CanonicalText.SplitLines(ReadEmbedded("VbaLoop.assets.LogShim.bas")))) };
        yield return new object[] { SampleProject.FileName, string.Join("\r\n", SampleProject.ModuleLines()) };
        yield return new object[] { "modBuildInfo (module)",
            VersionStamp.BuildInfoModule("1.2.3", "abc1234-dirty", "2026-01-01 00:00:00") };
        yield return new object[] { "modBuildInfo (body)",
            string.Join("\r\n", VersionStamp.BuildInfoBody("1.2.3", "abc1234-dirty", "2026-01-01 00:00:00")) };
    }

    [Theory]
    [MemberData(nameof(InjectedVbaTexts))]
    public void InjectedVbaText_IsAsciiOnly(string name, string text)
    {
        Assert.False(string.IsNullOrEmpty(text), $"{name}: 비어 있음");
        var lines = CanonicalText.SplitLines(text);
        for (int i = 0; i < lines.Length; i++)
        {
            int col = Array.FindIndex(lines[i].ToCharArray(), c => c > 127);
            Assert.True(col < 0,
                col < 0 ? "" : $"{name}:{i + 1}:{col + 1}: 비ASCII 문자 U+{(int)lines[i][col]:X4} in: {lines[i]}");
        }
    }

    [Theory]
    [InlineData(1252)]
    [InlineData(932)]
    [InlineData(936)]
    [InlineData(949)]
    public void InjectedVbaText_PassesCp949GuardOnCommonAcps(int codePage)
    {
        foreach (var row in InjectedVbaTexts())
        {
            var violations = Cp949Guard.Check(CanonicalText.SplitLines((string)row[1]), codePage);
            Assert.True(violations.Count == 0, $"{row[0]}: CP{codePage} 위반 {violations.Count}건");
        }
    }

    private static string ReadEmbedded(string logicalName)
    {
        using var stream = typeof(SampleProject).Assembly.GetManifestResourceStream(logicalName)
            ?? throw new InvalidOperationException($"임베드 리소스 없음: {logicalName}");
        using var reader = new StreamReader(stream, new UTF8Encoding(false));
        return reader.ReadToEnd();
    }
}
