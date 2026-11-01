using VbaLoop.Commands;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>
/// doctor의 PATH shim 등록 판정과 공통 안내 순수 헬퍼 검증.
/// </summary>
public class DoctorCommandTests
{
    private const string Shim = @"C:\Users\me\AppData\Local\vbaloop\bin";

    [Theory]
    [InlineData(3, 2, false, true)]
    [InlineData(3, 2, true, false)]
    [InlineData(3, 0, false, false)]
    [InlineData(0, 0, false, false)]
    public void DescribeOrphans_ClarifiesCleanupScope_AndPreservesConfirmationHint(
        int count, int killable, bool killOrphans, bool expectsHint)
    {
        string detail = DoctorCommand.DescribeOrphans(count, killable, killOrphans);

        if (count == 0)
            Assert.Equal("no windowless orphan EXCEL processes", detail);
        else
        {
            Assert.StartsWith($"{count} windowless EXCEL processes ({killable} to clean up:", detail);
            Assert.Contains("in the current Windows session, excluding the current project's live session, /automation, not quarantined)", detail);
            Assert.DoesNotContain("not current session", detail);
        }
        Assert.Equal(expectsHint, detail.Contains(
            " — clean up only after the user confirms: vbaloop doctor --kill-orphans"));
    }

    [Fact]
    public void IsOnPath_ExactMatch_True()
    {
        Assert.True(DoctorCommand.IsOnPath($@"C:\other;{Shim};C:\more", Shim));
    }

    [Fact]
    public void IsOnPath_TrailingBackslashDifference_Tolerated()
    {
        Assert.True(DoctorCommand.IsOnPath($@"{Shim}\", Shim));
        Assert.True(DoctorCommand.IsOnPath(Shim, Shim + @"\"));
    }

    [Fact]
    public void IsOnPath_CaseInsensitive_True()
    {
        Assert.True(DoctorCommand.IsOnPath(Shim.ToUpperInvariant(), Shim.ToLowerInvariant()));
    }

    [Fact]
    public void IsOnPath_SurroundingSpaces_Tolerated()
    {
        Assert.True(DoctorCommand.IsOnPath($@"C:\a;  {Shim}  ;C:\b", Shim));
    }

    [Fact]
    public void IsOnPath_NotPresent_False()
    {
        Assert.False(DoctorCommand.IsOnPath(@"C:\other;C:\more", Shim));
    }

    [Fact]
    public void IsOnPath_EmptyOrNull_False()
    {
        Assert.False(DoctorCommand.IsOnPath(null, Shim));
        Assert.False(DoctorCommand.IsOnPath("", Shim));
    }
}
