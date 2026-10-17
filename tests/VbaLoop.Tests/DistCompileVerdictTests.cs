using VbaLoop.Core;
using Xunit;
using static VbaLoop.Core.DistCompileVerdict;

namespace VbaLoop.Tests;

public class DistCompileVerdictTests
{
    private static readonly string[] NoModal = Array.Empty<string>();

    [Fact]
    public void CleanCompile_Publishable()
    {
        var v = Decide(controlFound: true, enabledBefore: true, NoModal, executeError: null, enabledAfter: false, timedOut: false);
        Assert.Equal(Kind.Compiled, v.Kind);
        Assert.True(v.Publishable);
    }

    [Fact]
    public void DisabledControl_TreatedAsAlreadyCompiled()
    {
        var v = Decide(true, enabledBefore: false, NoModal, null, null, false);
        Assert.Equal(Kind.AlreadyCompiled, v.Kind);
        Assert.True(v.Publishable);
    }

    [Fact]
    public void Modal_IsCompileError_NotPublishable_TextReported()
    {
        var modal = new[] { "Microsoft Visual Basic for Applications | 컴파일 오류: 변수가 정의되지 않았습니다. | 확인" };
        var v = Decide(true, true, modal, executeError: "Exception from HRESULT", enabledAfter: true, timedOut: false);
        Assert.Equal(Kind.CompileError, v.Kind); // Execute 예외보다 모달이 우선
        Assert.False(v.Publishable);
        Assert.Contains("변수가 정의되지 않았습니다", v.Detail);
    }

    [Fact]
    public void ModalWithEmptyBody_StillCompileError()
    {
        var v = Decide(true, true, new[] { "" }, null, false, false);
        Assert.Equal(Kind.CompileError, v.Kind);
        Assert.NotNull(v.Detail);
    }

    [Fact]
    public void Timeout_CheckFailed()
        => Assert.False(Decide(true, true, NoModal, "RPC 서버를 사용할 수 없습니다", null, timedOut: true).Publishable);

    [Fact]
    public void ControlMissing_CheckFailed()
    {
        var v = Decide(controlFound: false, false, NoModal, null, null, false);
        Assert.Equal(Kind.CheckFailed, v.Kind);
        Assert.Contains("578", v.Detail);
    }

    [Fact]
    public void ExecuteErrorWithoutModal_CheckFailed()
        => Assert.Equal(Kind.CheckFailed, Decide(true, true, NoModal, "boom", null, false).Kind);

    [Theory]
    [InlineData(true)]
    [InlineData(null)]
    public void NoModalButStateUnclear_PublishableWithWarning(bool? enabledAfter)
    {
        var v = Decide(true, true, NoModal, null, enabledAfter, false);
        Assert.Equal(Kind.Unverified, v.Kind);
        Assert.True(v.Publishable);
        Assert.NotNull(v.Detail);
    }

    [Theory]
    [InlineData(Kind.Compiled, "compiled")]
    [InlineData(Kind.AlreadyCompiled, "already_compiled")]
    [InlineData(Kind.Unverified, "unverified")]
    [InlineData(Kind.CompileError, "compile_error")]
    [InlineData(Kind.CheckFailed, "check_failed")]
    public void Labels_StableForJson(Kind kind, string label)
        => Assert.Equal(label, new Verdict(kind, null).Label);
}
