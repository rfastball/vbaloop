using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>session attach 대상 선택(설계 §11 L5).</summary>
public class AttachMatchTests
{
    private const string Canon = @"C:\Users\u\Desktop\Proj\Book.xlsm";

    [Theory]
    [InlineData(@"C:\Users\u\Desktop\Proj\Book.xlsm")]
    [InlineData(@"c:\users\U\desktop\proj\BOOK.XLSM")]
    [InlineData(@"C:/Users/u/Desktop/Proj/Book.xlsm")]
    [InlineData(@"C:\Users\u\Desktop\Proj\.\sub\..\Book.xlsm")]
    public void IsCanonical_NormalizesCaseSeparatorsAndSegments(string fullName)
    {
        Assert.True(AttachMatch.IsCanonical(fullName, Canon));
    }

    [Theory]
    [InlineData(@"C:\Users\u\Documents\Book.xlsm")]
    [InlineData(@"C:\Users\u\Desktop\Proj\Book.xlsx")]
    [InlineData("https://d.docs.live.net/abc123/Proj/Book.xlsm")]
    public void IsCanonical_RejectsOtherPaths(string fullName)
    {
        Assert.False(AttachMatch.IsCanonical(fullName, Canon));
    }

    [Fact]
    public void Pick_FindsWritableCanonical_AmongOtherWorkbooks()
    {
        var p = AttachMatch.Pick(new[]
        {
            new AttachCandidate(100, @"C:\Users\u\Documents\Budget.xlsx", false),
            new AttachCandidate(100, @"c:\users\u\desktop\proj\book.xlsm", false),
            new AttachCandidate(200, @"C:\Other.xlsx", false),
        }, Canon);
        Assert.Equal(AttachPickKind.Found, p.Kind);
        Assert.Equal(100, p.Target!.Pid);
    }

    [Fact]
    public void Pick_PrefersWritable_OverReadOnlyCopyInAnotherInstance()
    {
        var p = AttachMatch.Pick(new[]
        {
            new AttachCandidate(200, Canon, true),
            new AttachCandidate(100, Canon, false),
        }, Canon);
        Assert.Equal(AttachPickKind.Found, p.Kind);
        Assert.Equal(100, p.Target!.Pid);
    }

    [Fact]
    public void Pick_ReadOnlyOnly_IsRefused()
    {
        var p = AttachMatch.Pick(new[] { new AttachCandidate(100, Canon, true) }, Canon);
        Assert.Equal(AttachPickKind.ReadOnlyOnly, p.Kind);
    }

    [Fact]
    public void Pick_WritableInTwoInstances_IsAmbiguous()
    {
        var p = AttachMatch.Pick(new[]
        {
            new AttachCandidate(100, Canon, false),
            new AttachCandidate(200, Canon, false),
        }, Canon);
        Assert.Equal(AttachPickKind.Ambiguous, p.Kind);
        Assert.Null(p.Target);
    }

    [Fact]
    public void Pick_NotOpen_ReportsSameNameSuspects()
    {
        var p = AttachMatch.Pick(new[]
        {
            new AttachCandidate(100, "https://d.docs.live.net/abc123/Proj/Book.xlsm", false),
            new AttachCandidate(100, @"C:\x\Unrelated.xlsx", false),
        }, Canon);
        Assert.Equal(AttachPickKind.NotOpen, p.Kind);
        Assert.Single(p.Suspects);
        Assert.Contains("docs.live.net", p.Suspects[0]);
    }

    [Fact]
    public void Pick_NoCandidates_NotOpen()
    {
        var p = AttachMatch.Pick(Array.Empty<AttachCandidate>(), Canon);
        Assert.Equal(AttachPickKind.NotOpen, p.Kind);
        Assert.Empty(p.Suspects);
    }

    [Fact]
    public void SessionMeta_OwnerRoundTrip_AndLegacyIsTool()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vf-meta-" + Guid.NewGuid().ToString("N"));
        try
        {
            var path = Path.Combine(dir, "session.json");
            new SessionMeta { Pid = 1, Owner = SessionMeta.OwnerExternal }.Save(path);
            Assert.True(SessionMeta.Load(path)!.IsExternal);

            File.WriteAllText(path, "{\"pid\":1,\"excelStartTime\":\"x\",\"hwnd\":1,\"workbookFullName\":\"a\",\"createdAt\":\"t\"}");
            Assert.False(SessionMeta.Load(path)!.IsExternal);
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { /* best effort */ }
        }
    }
}
