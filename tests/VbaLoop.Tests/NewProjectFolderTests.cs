using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class NewProjectFolderTests
{
    [Theory]
    [InlineData("MyTool", "MyTool")]
    [InlineData("  \"월별 정산\"  ", "월별 정산")]
    public void ValidateName_TrimsAndUnquotes(string raw, string expected)
        => Assert.Equal(expected, NewProjectFolder.ValidateName(raw));

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("a/b")]
    [InlineData(@"a\b")]
    [InlineData("a:b")]
    [InlineData("what?")]
    [InlineData("..")]
    [InlineData("name.")]
    [InlineData("CON")]
    [InlineData("lpt1")]
    [InlineData("CON.foo")]   // 장치 이름은 확장자가 붙어도 예약
    [InlineData("LPT1.test")]
    [InlineData("nul .txt")]
    public void ValidateName_Rejects(string raw)
        => Assert.Throws<CliError>(() => NewProjectFolder.ValidateName(raw));

    [Theory]
    [InlineData("Console")]
    [InlineData("COM10")]
    [InlineData("my.CON")]
    public void ValidateName_AllowsNamesMerelyContainingDeviceNames(string raw)
        => Assert.Equal(raw, NewProjectFolder.ValidateName(raw));

    [Fact]
    public void Resolve_NewOrEmptyFolder_ReturnsPath()
    {
        string parent = Path.Combine(Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        try
        {
            Assert.Equal(Path.Combine(parent, "Tool"), NewProjectFolder.Resolve(parent, "Tool"));
            Directory.CreateDirectory(Path.Combine(parent, "Empty"));
            Assert.Equal(Path.Combine(parent, "Empty"), NewProjectFolder.Resolve(parent, "Empty"));
        }
        finally
        {
            try { Directory.Delete(parent, recursive: true); } catch { }
        }
    }

    [Fact]
    public void Resolve_NonEmptyFolder_Throws()
    {
        string parent = Path.Combine(Path.GetTempPath(), "vbaloop-tests", Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(Path.Combine(parent, "Taken"));
            File.WriteAllText(Path.Combine(parent, "Taken", "x.txt"), "");
            var e = Assert.Throws<CliError>(() => NewProjectFolder.Resolve(parent, "Taken"));
            Assert.Contains("--folder", e.Hint);
        }
        finally
        {
            try { Directory.Delete(parent, recursive: true); } catch { }
        }
    }

    [Fact]
    public void ExplorerDefaultParent_IsUnderUserProfile()
        => Assert.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            NewProjectFolder.ExplorerDefaultParent);
}
