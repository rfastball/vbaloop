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

    private static DistModule Cls(string name, params string[] body)
        => new(name, ".cls", new[] { $"Attribute VB_Name = \"{name}\"" }, body);

    [Fact]
    public void Scan_LocalShadowsExcludedSymbol_NotFlagged()
    {
        // 이슈 #4 재현: 프로시저 지역 Dim은 제외 모듈 Public보다 우선 바인딩
        var included = new[] { Mod("lkConfig", "Private Sub X(): Dim grid As Variant: grid = 1: End Sub") };
        var excluded = new[] { Mod("TestSupport", "Public Function Grid() As Variant", "End Function") };
        Assert.Empty(DistCompileScan.Scan(included, excluded));
    }

    [Fact]
    public void Scan_ParamsAndMultiLineLocals_NotFlagged()
    {
        var included = new[]
        {
            Mod("bkSchema",
                "Public Function Build(ByVal mode As Long, _",
                "        Optional ByRef grid As Variant) As Long",
                "    Dim a As Long, message(1 To 2, 3) As String",
                "    Static counter As Long",
                "    Const LIMIT As Long = 3",
                "    message(1, 3) = \"x\"",
                "    Build = mode + LIMIT + counter",
                "    grid = Empty",
                "End Function"),
        };
        var excluded = new[]
        {
            Mod("TestSupport",
                "Public Function Grid() As Variant", "End Function",
                "Public Message As String",
                "Public Const LIMIT As Long = 9",
                "Public Sub Counter()", "End Sub",
                "Public Mode As Long"),
        };
        Assert.Empty(DistCompileScan.Scan(included, excluded));
    }

    [Fact]
    public void Scan_LocalScopeEndsWithProcedure()
    {
        // Sub A의 지역 grid는 Sub B까지 가리지 못한다 → B의 grid는 제외 모듈 참조
        var included = new[]
        {
            Mod("lkConfig",
                "Sub A()", "    Dim grid As Variant", "    grid = 1", "End Sub",
                "Sub B()", "    x = grid()", "End Sub"),
        };
        var excluded = new[] { Mod("TestSupport", "Public Function Grid() As Variant", "End Function") };
        var hit = Assert.Single(DistCompileScan.Scan(included, excluded));
        Assert.Equal(6, hit.Line);
        Assert.Equal("grid", hit.Symbol);
    }

    [Fact]
    public void Scan_ExcludedClassMembers_NotCandidates()
    {
        // 클래스 멤버는 다른 모듈에서 비한정으로 참조 불가 → 후보 아님
        var included = new[] { Mod("bkBuilder", "Sub Go()", "    Debug.Print Message, Mode", "End Sub") };
        var excluded = new[] { Cls("TestValidator", "Public Message As String", "Public Mode As Long") };
        Assert.Empty(DistCompileScan.Scan(included, excluded));
    }

    [Fact]
    public void Scan_ExcludedClassNameAsType_Flagged()
    {
        var included = new[] { Mod("bkBuilder", "Sub Go()", "    Dim v As TestValidator", "    Set v = New TestValidator", "End Sub") };
        var excluded = new[] { Cls("TestValidator", "Public Message As String") };
        var r = DistCompileScan.Scan(included, excluded);
        Assert.Equal(new[] { 2, 3 }, r.Select(d => d.Line));
        Assert.All(r, d => Assert.Equal("module", d.Kind));
    }

    [Fact]
    public void Scan_OwnModuleMember_NotFlagged()
    {
        // 이슈 #4 재현: 잔존 모듈 자신의 Public Elapsed는 모듈 내부에서 자기 멤버로 바인딩
        var included = new[]
        {
            Mod("lkProgress",
                "Private mStart As Double",
                "Public Function Elapsed() As Double",
                "    Elapsed = Timer - mStart",
                "End Function",
                "Public Sub Report()",
                "    Debug.Print Elapsed",
                "End Sub"),
        };
        var excluded = new[] { Mod("TestSupport", "Public Function Elapsed() As Double", "End Function") };
        Assert.Empty(DistCompileScan.Scan(included, excluded));
    }

    [Fact]
    public void Scan_OtherIncludedStandardModulePublic_NotFlagged()
    {
        // 제외 후에도 같은 이름의 잔존 표준 모듈 Public이 남아 있으면 그쪽으로 바인딩
        var included = new[]
        {
            Mod("lkProgress", "Public Function Elapsed() As Double", "End Function"),
            Mod("lkReports", "Sub Go()", "    Debug.Print Elapsed", "End Sub"),
        };
        var excluded = new[] { Mod("TestSupport", "Public Function Elapsed() As Double", "End Function") };
        Assert.Empty(DistCompileScan.Scan(included, excluded));
    }

    [Fact]
    public void Scan_NamedArgumentNotTreatedAsStatementSeparator()
    {
        // `mode:=1`의 `:`는 문장 구분자가 아니다 — 지역 선언 파싱이 깨지지 않아야 함
        var included = new[]
        {
            Mod("bkSchema", "Sub Go()", "    Dim mode As Long: Call Foo(mode:=1): mode = 2", "End Sub"),
        };
        var excluded = new[] { Mod("TestSupport", "Public Mode As Long") };
        Assert.Empty(DistCompileScan.Scan(included, excluded));
    }
}
