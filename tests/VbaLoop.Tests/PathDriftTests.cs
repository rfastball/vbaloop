using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class PathDriftTests
{
    [Fact]
    public void SamePath_CaseAndTrailingSlash_Insensitive()
    {
        Assert.True(PathDrift.SamePath(@"C:\Work\Book.xlsm", @"c:\work\BOOK.XLSM"));
        Assert.True(PathDrift.SamePath(@"C:\Work\Sub\", @"C:\Work\Sub"));
    }

    [Fact]
    public void SamePath_RelativeSegments_Normalized()
    {
        Assert.True(PathDrift.SamePath(@"C:\Work\.\Sub\..\Book.xlsm", @"C:\Work\Book.xlsm"));
    }

    [Fact]
    public void SamePath_DifferentDirectory_False()
    {
        // 버그 실측 시나리오: Desktop 정본 vs Documents 오탈출 사본
        Assert.False(PathDrift.SamePath(
            @"C:\Users\u\Documents\SampleRpa.xlsm",
            @"C:\Users\u\Desktop\Projects\SampleRpa\SampleRpa.xlsm"));
    }

    [Fact]
    public void SamePath_NonPathString_FallsBackToRawCompare()
    {
        // 클라우드 파일의 FullName은 URL일 수 있다 — 예외 없이 raw 비교로 폴백
        Assert.True(PathDrift.SamePath(
            "https://d.docs.live.net/x/Book.xlsm", "HTTPS://d.docs.live.net/x/Book.xlsm"));
        Assert.False(PathDrift.SamePath(
            "https://d.docs.live.net/x/Book.xlsm", @"C:\Work\Book.xlsm"));
    }

    [Fact]
    public void FindDriftSuspects_SameNameOtherPath_Detected()
    {
        // 버그 실측: session stop의 foreign 목록에 자기 워크북이 Documents 경로로 나타남
        var suspects = PathDrift.FindDriftSuspects(
            new[] { @"C:\Users\u\Documents\SampleRpa.xlsm" },
            @"C:\Users\u\Desktop\Projects\SampleRpa\SampleRpa.xlsm");
        Assert.Equal(new[] { @"C:\Users\u\Documents\SampleRpa.xlsm" }, suspects);
    }

    [Fact]
    public void FindDriftSuspects_TargetItself_NotSuspect()
    {
        var suspects = PathDrift.FindDriftSuspects(
            new[] { @"C:\Work\Book.xlsm", @"c:\WORK\book.xlsm" },
            @"C:\Work\Book.xlsm");
        Assert.Empty(suspects);
    }

    [Fact]
    public void FindDriftSuspects_DifferentFileName_NotSuspect()
    {
        var suspects = PathDrift.FindDriftSuspects(
            new[] { @"C:\Users\u\Documents\Other.xlsm" },
            @"C:\Work\Book.xlsm");
        Assert.Empty(suspects);
    }

    [Fact]
    public void FindDriftSuspects_EmptyList_Empty()
    {
        Assert.Empty(PathDrift.FindDriftSuspects(Array.Empty<string>(), @"C:\Work\Book.xlsm"));
    }
}
