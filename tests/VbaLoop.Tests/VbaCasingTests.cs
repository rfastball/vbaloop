using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class VbaCasingTests
{
    // ── LineEquivalent ──────────────────────────────────────────────────────

    [Theory]
    // 식별자 케이싱만 상이 → 등가
    [InlineData("    If LenB(ThisWorkbook.path) = 0 Then Exit Function",
                "    If LenB(ThisWorkbook.Path) = 0 Then Exit Function", true)]
    [InlineData("Dim path As String", "Dim Path As String", true)]
    [InlineData("x = a.name", "x = a.Name", true)]
    [InlineData("path1 = 0", "Path1 = 0", true)]
    [InlineData("", "", true)]
    // 문자열 리터럴 내용은 의미적 → 비등가
    [InlineData("x = \"Hello\"", "x = \"hello\"", false)]
    [InlineData("s = \"it's\"", "s = \"IT'S\"", false)]
    // 주석 텍스트는 의미적 → 비등가
    [InlineData("x = 1 ' Note", "x = 1 ' note", false)]
    // 코드 재케이싱 + 주석 동일 → 등가
    [InlineData("A.path = 1 ' keep", "A.Path = 1 ' keep", true)]
    // 연산자/숫자 변경 → 비등가
    [InlineData("n = n + 1", "n = n + 2", false)]
    // 공백/탭 차이(포맷) → 비등가(케이싱과 분리)
    [InlineData("x  = 1", "x = 1", false)]
    // 토큰 수 불일치 → 비등가
    [InlineData("Foo Bar", "Foo", false)]
    // 길이 다른 식별자 → 비등가
    [InlineData("path", "path1", false)]
    // REM 주석 텍스트는 의미적(가드) → 비등가
    [InlineData("REM Hello", "REM hello", false)]
    public void LineEquivalent_Cases(string a, string b, bool expected)
        => Assert.Equal(expected, VbaCasing.LineEquivalent(a, b));

    [Fact]
    public void LineEquivalent_IsSymmetric()
    {
        Assert.True(VbaCasing.LineEquivalent("a.Name", "a.name"));
        Assert.True(VbaCasing.LineEquivalent("a.name", "a.Name"));
    }

    // ── ReconcileCasing ─────────────────────────────────────────────────────

    [Fact]
    public void ReconcileCasing_Equivalent_AdoptsSrcSpelling_KeepsWhitespaceAndComment()
    {
        // 케이싱만 다른(gap·주석 동일) 등가 라인 → 식별자 철자만 src 채택, 들여쓰기·주석 보존
        var result = VbaCasing.ReconcileCasing(
            "    x = ThisWorkbook.path  ' note",
            "    x = ThisWorkbook.Path  ' note");
        Assert.Equal("    x = ThisWorkbook.Path  ' note", result);
    }

    [Fact]
    public void ReconcileCasing_NotEquivalent_ReturnsWorkbookUnchanged()
    {
        // 실변경(숫자)은 비등가 → 워크북 그대로(되돌리지 않음)
        Assert.Equal("n = n + 1", VbaCasing.ReconcileCasing("n = n + 1", "n = n + 2"));
    }

    [Fact]
    public void ReconcileCasing_RealChangePlusCasingSameLine_KeepsWorkbook()
    {
        // 같은 라인에 실변경 + 케이싱 혼재 → 비등가 → 워크북 전체 유지(실변경 2 보존)
        Assert.Equal("x = a.path + 2", VbaCasing.ReconcileCasing("x = a.path + 2", "x = a.Path + 1"));
    }

    [Fact]
    public void ReconcileCasing_StringLiteralCaseChange_NeverReverted()
    {
        // 문자열 케이싱 변경은 비등가 → 워크북 그대로(구 미봉책이 되돌리던 결함 방어)
        Assert.Equal("x = \"hello\"", VbaCasing.ReconcileCasing("x = \"hello\"", "x = \"Hello\""));
    }

    // ── ReconcileBlock (LCS 정렬 — 인덱스 정렬 미봉책이 표현 못 하던 케이스) ────

    [Fact]
    public void ReconcileBlock_ReproducesBugReport_Path()
    {
        // .path → src 철자 .Path로 수렴
        var wb = new[] { "    If LenB(ThisWorkbook.path) = 0 Then Exit Function" };
        var src = new[] { "    If LenB(ThisWorkbook.Path) = 0 Then Exit Function" };
        Assert.Equal(src, VbaCasing.ReconcileBlock(wb, src));
    }

    [Fact]
    public void ReconcileBlock_InsertBeforeRecasedLine_StillConverges()
    {
        // 재케이싱 라인 앞에 실제 라인 삽입 — 인덱스 정렬이면 어긋나 실패하던 케이스
        var wb = new[] { "Dim q As Long", "A.path = 1" };
        var src = new[] { "A.Path = 1" };
        Assert.Equal(new[] { "Dim q As Long", "A.Path = 1" }, VbaCasing.ReconcileBlock(wb, src));
    }

    [Fact]
    public void ReconcileBlock_DeleteBeforeRecasedLine_StillConverges()
    {
        // src에만 있던 앞 라인이 워크북에서 삭제됨 — 남은 재케이싱 라인은 여전히 수렴
        var wb = new[] { "A.path = 1" };
        var src = new[] { "Dim gone As Long", "A.Path = 1" };
        Assert.Equal(new[] { "A.Path = 1" }, VbaCasing.ReconcileBlock(wb, src));
    }

    [Fact]
    public void ReconcileBlock_MixedCasingRealChangeAndInsert()
    {
        var wb = new[] { "A.path = 1", "n = n + 9", "Dim inserted", "B.name" };
        var src = new[] { "A.Path = 1", "n = n + 1", "B.Name" };
        // 1행: 케이싱 수렴 / 2행: 실변경 → 워크북 유지 / 3행: 삽입 → 통과 / 4행: 케이싱 수렴
        Assert.Equal(new[] { "A.Path = 1", "n = n + 9", "Dim inserted", "B.Name" },
            VbaCasing.ReconcileBlock(wb, src));
    }

    [Fact]
    public void ReconcileBlock_EmptyEitherSide_ReturnsWorkbook()
    {
        Assert.Empty(VbaCasing.ReconcileBlock(System.Array.Empty<string>(), new[] { "a" }));
        Assert.Equal(new[] { "a" }, VbaCasing.ReconcileBlock(new[] { "a" }, System.Array.Empty<string>()));
    }

    [Fact]
    public void ReconcileBlock_Idempotent_AfterConvergence()
    {
        var wb = new[] { "Dim q", "A.path", "n = 1" };
        var src = new[] { "A.Path", "n = 1" };
        var once = VbaCasing.ReconcileBlock(wb, src);
        var twice = VbaCasing.ReconcileBlock(once, src);
        Assert.Equal(once, twice);
    }

    // ── 왕복 바이트 항등 시뮬레이션 (COM 없이 VBE 재케이싱 모사) ─────────────

    private static string Canon(IReadOnlyList<string> lines) => CanonicalText.ToCanonical(lines);

    [Fact]
    public void RoundTrip_RecasingOnly_ByteIdenticalToSrc()
    {
        // src 본문. 워크북은 VBE 전역 케이스 테이블이 Path/Name을 소문자로 재작성한 상태를 모사.
        var src = new[]
        {
            "    If LenB(ThisWorkbook.Path) = 0 Then Exit Function",
            "    nm = ThisWorkbook.Name",
            "    total = total + 1",
        };
        var workbook = new[]
        {
            "    If LenB(ThisWorkbook.path) = 0 Then Exit Function",
            "    nm = ThisWorkbook.name",
            "    total = total + 1",
        };
        var reconciled = VbaCasing.ReconcileBlock(workbook, src);
        // pull이 다시 써도 src 바이트가 동일 — 가짜 diff 없음
        Assert.Equal(Canon(src), Canon(reconciled));
        // 2회차 멱등
        Assert.Equal(Canon(src), Canon(VbaCasing.ReconcileBlock(reconciled, src)));
    }

    [Fact]
    public void RoundTrip_RecasingPlusStringEdit_StringSurvives_CasingConverges()
    {
        // 재케이싱 + 다른 라인의 문자열 리터럴 실편집이 동시에 일어남
        var src = new[]
        {
            "    p = ThisWorkbook.Path",
            "    MsgBox \"Old\"",
        };
        var workbook = new[]
        {
            "    p = ThisWorkbook.path",   // VBE 재케이싱
            "    MsgBox \"New\"",           // 사용자 실편집(문자열)
        };
        var reconciled = VbaCasing.ReconcileBlock(workbook, src);
        // 재케이싱 라인은 src 철자로 수렴, 문자열 편집은 워크북 그대로 생존
        Assert.Equal(new[] { "    p = ThisWorkbook.Path", "    MsgBox \"New\"" }, reconciled);
    }
}
