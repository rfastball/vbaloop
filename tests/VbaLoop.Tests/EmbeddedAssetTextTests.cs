using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>
/// init이 사용자 프로젝트에 복사하는 임베드 자산에 제어 문자가 섞이지 않았는지 확인한다.
/// 스크립트로 문서를 고칠 때 경로의 \v·\b가 수직 탭(U+000B)·백스페이스(U+0008)로 바뀐 적이 있다.
/// 허용: CR, LF, TAB.
/// </summary>
public class EmbeddedAssetTextTests
{
    public static IEnumerable<object[]> AssetNames() => new[]
    {
        new object[] { "VbaLoop.assets.LogShim.bas" },
        new object[] { "VbaLoop.assets.AgentDocsBlock.md" },
        new object[] { "VbaLoop.assets.SkillMd" },
        new object[] { "VbaLoop.assets.PitfallsMd" },
        new object[] { "VbaLoop.assets.DistMd" },
        new object[] { "VbaLoop.assets.FormsMd" },
    };

    [Theory]
    [MemberData(nameof(AssetNames))]
    public void EmbeddedAsset_HasNoControlCharacters(string logicalName)
    {
        using var stream = typeof(SampleProject).Assembly.GetManifestResourceStream(logicalName)
            ?? throw new InvalidOperationException($"임베드 리소스 없음: {logicalName}");
        using var reader = new StreamReader(stream, new UTF8Encoding(false));
        var lines = CanonicalText.SplitLines(reader.ReadToEnd());
        for (int i = 0; i < lines.Length; i++)
        {
            int col = Array.FindIndex(lines[i].ToCharArray(), c => char.IsControl(c) && c != '\t');
            Assert.True(col < 0,
                col < 0 ? "" : $"{logicalName}:{i + 1}:{col + 1}: 제어 문자 U+{(int)lines[i][col]:X4} in: {lines[i]}");
        }
    }
}
