using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class DistCompileScanTests
{
    private static DistModule Mod(string name, params string[] body)
        => new(name, ".bas", new[] { $"Attribute VB_Name = \"{name}\"" }, body);

    [Fact]
    public void PublicSymbols_ExtractsProceduresAndPublicData()
    {
        var body = new[]
        {
            "Public Sub RunAll()",
            "End Sub",
            "Function Helper() As Long",   // 무수식어 → Public
            "End Function",
            "Private Sub Hidden()",         // Private → 제외
            "End Sub",
            "Public Property Get Name() As String",
            "End Property",
            "Public Const MAX As Long = 10",
            "Public gState As Long",
            "Private mLocal As Long",       // Private 변수 → 제외
        };
        var syms = DistCompileScan.PublicSymbols(body).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains("RunAll", syms);
        Assert.Contains("Helper", syms);
        Assert.Contains("Name", syms);
        Assert.Contains("MAX", syms);
        Assert.Contains("gState", syms);
        Assert.DoesNotContain("Hidden", syms);
        Assert.DoesNotContain("mLocal", syms);
    }

    [Fact]
    public void Scan_CleanSeparation_NoDangling()
    {
        // Test_* 는 서로만 참조 → 제외해도 잔존 참조 0
        var included = new[] { Mod("modEngine", "Public Sub Go()", "    Dim x As Long", "End Sub") };
        var excluded = new[] { Mod("Test_Engine", "Public Sub RunAll()", "    Go", "End Sub") };
        var r = DistCompileScan.Scan(included, excluded);
        Assert.Empty(r);
    }

    [Fact]
    public void Scan_QualifiedRefToExcludedModule_Flagged()
    {
        var included = new[] { Mod("modEngine", "Sub Go()", "    Probe.Ping", "End Sub") };
        var excluded = new[] { Mod("Probe", "Public Sub Ping()", "End Sub") };
        var r = DistCompileScan.Scan(included, excluded);
        var hit = Assert.Single(r);
        Assert.Equal("Probe", hit.Symbol);
        Assert.Equal("module", hit.Kind);
        Assert.Equal("modEngine", hit.InModule);
    }

    [Fact]
    public void Scan_UnqualifiedProcRef_Flagged()
    {
        var included = new[] { Mod("modEngine", "Sub Go()", "    RunAll", "End Sub") };
        var excluded = new[] { Mod("Test_X", "Public Sub RunAll()", "End Sub") };
        var r = DistCompileScan.Scan(included, excluded);
        var hit = Assert.Single(r);
        Assert.Equal("RunAll", hit.Symbol);
        Assert.Equal("symbol", hit.Kind);
    }

    [Fact]
    public void Scan_RefInsideStringOrComment_NotFlagged()
    {
        var included = new[]
        {
            Mod("modEngine",
                "Sub Go()",
                "    Dim s As String",
                "    s = \"RunAll and Probe.Ping\"",  // 문자열 안
                "    ' RunAll here too",              // 주석 안
                "End Sub"),
        };
        var excluded = new[]
        {
            Mod("Probe", "Public Sub Ping()", "End Sub"),
            Mod("Test_X", "Public Sub RunAll()", "End Sub"),
        };
        var r = DistCompileScan.Scan(included, excluded);
        Assert.Empty(r);
    }

    [Fact]
    public void Scan_MemberAccessNotConfusedWithExcludedSymbol()
    {
        // obj.Ping 은 excluded 모듈 심볼이 아니라 멤버 접근 — precededByDot로 스킵
        var included = new[] { Mod("modEngine", "Sub Go()", "    obj.Ping", "End Sub") };
        var excluded = new[] { Mod("Test_X", "Public Sub Ping()", "End Sub") };
        var r = DistCompileScan.Scan(included, excluded);
        Assert.Empty(r);
    }
}
