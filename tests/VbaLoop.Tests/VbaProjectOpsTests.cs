using VbaLoop.Core;
using VbaLoop.Excel;
using Xunit;

namespace VbaLoop.Tests;

/// <summary>
/// ReplaceCode의 주입 후 컴파일 파손형 () 팬텀 스윕 + 말미 빈 줄 관용 + 논리행수 정수 가드
/// 검증. 실제 팬텀은 VBE AddFromString 런타임
/// 동작이라 순수 재현 불가 — CodeModule을 흉내내는 스텁으로 스윕/가드 로직만 격리 검증한다.
/// </summary>
public class VbaProjectOpsTests
{
    // vbext_CodeModule 오리 타이핑 스텁. dynamic 디스패치로 실제 COM 멤버명을 미러.
    // public이어야 dynamic 바인더가 다른 어셈블리(VbaLoop)에서 멤버를 본다.
    // Phantoms: AddFromString이 주입 본문 뒤에 덧붙이는 고아 라인들(VBE 오파싱 흉내).
    public sealed class FakeCodeModule
    {
        public readonly List<string> L = new();
        public string[] Phantoms = System.Array.Empty<string>();

        public int CountOfLines => L.Count;
        public FakeLines Lines => new(L);
        public void DeleteLines(int start, int count) => L.RemoveRange(start - 1, count);
        public void AddFromString(string s)
        {
            L.AddRange(s.Split("\r\n"));
            L.AddRange(Phantoms);
        }
    }

    public sealed class FakeLines
    {
        private readonly List<string> _l;
        public FakeLines(List<string> l) => _l = l;
        public string this[int start, int count] => string.Join("\r\n", _l.Skip(start - 1).Take(count));
    }

    [Fact]
    public void ReplaceCode_NoPhantom_InjectsAsIs()
    {
        var fake = new FakeCodeModule();
        dynamic cm = fake;
        VbaProjectOps.ReplaceCode(cm, new[] { "Public Function F() As Long", "End Function" });
        Assert.Equal(new[] { "Public Function F() As Long", "End Function" }, fake.L);
    }

    [Fact]
    public void ReplaceCode_StripsTrailingPhantomParens()
    {
        // 컨티뉴에이션 Declare 주입 시 VBE가 끝에 붙이는 () 팬텀 한 줄을 제거
        var fake = new FakeCodeModule { Phantoms = new[] { "()" } };
        dynamic cm = fake;
        var body = new[]
        {
            "Private Declare PtrSafe Sub ReproMoveF Lib \"kernel32\" _",
            "    (ByVal a As LongPtr, ByVal b As LongPtr)",
            "Public Function F() As Long",
            "End Function",
        };
        VbaProjectOps.ReplaceCode(cm, body);
        Assert.Equal(body, fake.L);
        Assert.DoesNotContain("()", fake.L);
    }

    [Fact]
    public void ReplaceCode_ParenPhantomAboveTrailingBlank_StillStripped()
    {
        // () 팬텀 뒤에 무해한 빈 줄이 붙어도 최하단 빈 줄을 건너뛰고 ()를 제거, 빈 줄은 보존
        var fake = new FakeCodeModule { Phantoms = new[] { "()", "" } };
        dynamic cm = fake;
        VbaProjectOps.ReplaceCode(cm, new[] { "Sub A()", "End Sub" });
        Assert.Equal(new[] { "Sub A()", "End Sub", "" }, fake.L);
    }

    [Fact]
    public void ReplaceCode_TrailingBlankLine_ToleratedNotFailed()
    {
        // 실측(api_Compression): AddFromString이 말미에 무해한 빈 줄을 남긴다 — pull이 흡수하므로
        // 파손이 아니다. 정수 가드가 이를 오판(fail)하지 않아야 push가 막히지 않는다.
        var fake = new FakeCodeModule { Phantoms = new[] { "" } };
        dynamic cm = fake;
        VbaProjectOps.ReplaceCode(cm, new[] { "Sub A()", "End Sub" });
        Assert.Equal(new[] { "Sub A()", "End Sub", "" }, fake.L); // 빈 줄 보존
    }

    [Fact]
    public void ReplaceCode_NonPhantomMismatch_ThrowsAndRollsBack()
    {
        // 초과 라인이 알려진 팬텀도 빈 줄도 아니면 loud 실패(조용한 파손 차단) + 스냅샷 원복
        var fake = new FakeCodeModule { Phantoms = new[] { "Garbage" } };
        dynamic cm = fake;
        // 블록 람다 = Action — dynamic 반환(void)을 Func<object>로 잡으려는 바인더 오류 방지
        var ex = Assert.Throws<CliError>(() =>
        {
            VbaProjectOps.ReplaceCode(cm, new[] { "Sub A()", "End Sub" });
        });
        Assert.Contains("논리행수 불일치", ex.Message);
        Assert.Empty(fake.L); // 최초 빈 모듈로 원복
    }
}
