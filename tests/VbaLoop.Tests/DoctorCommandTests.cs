using VbaLoop.Commands;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>
/// doctor의 PATH shim 등록 판정 순수 헬퍼 검증.
/// </summary>
public class DoctorCommandTests
{
    private const string Shim = @"C:\Users\me\AppData\Local\vbaloop\bin";

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
