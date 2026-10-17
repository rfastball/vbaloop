using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class TestEnumeratorTests
{
    [Fact]
    public void FindsImplicitAndExplicitPublic_ExcludesPrivateFriend()
    {
        var lines = new[]
        {
            "Option Explicit",
            "Sub Test_Implicit()",
            "End Sub",
            "Public Sub Test_Explicit()",
            "End Sub",
            "Private Sub Test_Hidden()",
            "End Sub",
            "Friend Sub Test_Friend()",
            "End Sub",
            "Public Sub NotATest()",
            "End Sub",
            "Public Sub Test_WithArgs(x As Long)",
            "End Sub",
            "Public Function Test_Func() As Long",
            "End Function",
        };
        var found = TestEnumerator.FindTestSubs(lines);
        Assert.Equal(new[] { "Test_Implicit", "Test_Explicit" }, found);
    }

    [Fact]
    public void TrailingComment_Allowed()
    {
        var found = TestEnumerator.FindTestSubs(new[] { "Sub Test_A() ' 주석" });
        Assert.Equal(new[] { "Test_A" }, found);
    }

    [Fact]
    public void LineContinuation_Joined()
    {
        var found = TestEnumerator.FindTestSubs(new[] { "Public Sub Test_Cont _", "()" });
        Assert.Equal(new[] { "Test_Cont" }, found);
    }

    [Fact]
    public void CasePrefix_IsOrdinal()
    {
        // ^Test_ 매칭은 대소문자 구분 (설계 §4 test 열거 규칙)
        var found = TestEnumerator.FindTestSubs(new[] { "Sub test_lower()", "End Sub" });
        Assert.Empty(found);
    }
}
