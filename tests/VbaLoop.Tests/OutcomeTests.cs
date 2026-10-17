using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class OutcomeTests
{
    [Fact]
    public void ExitCode_Success_Zero()
    {
        var o = new Outcome { Command = "test" };
        Assert.Equal(0, o.ExitCode);
    }

    [Fact]
    public void ExitCode_GeneralFailure_Two()
    {
        var o = new Outcome { Command = "test" };
        o.Fail("일반 실패");
        Assert.Equal(2, o.ExitCode);
    }

    [Fact]
    public void ExitCode_EnvironmentFailure_Three()
    {
        var o = new Outcome { Command = "test" };
        o.Fail("세션 없음", "session start", environment: true);
        Assert.Equal(3, o.ExitCode);
        Assert.Equal("session start", o.Hint);
    }

    [Fact]
    public void Fail_EnvironmentFlag_Sticky()
    {
        // 환경 오류가 한 번 기록되면 뒤따르는 일반 Fail이 등급을 낮추지 못한다
        var o = new Outcome { Command = "test" };
        o.Fail("환경 오류", environment: true);
        o.Fail("후속 일반 오류");
        Assert.Equal(3, o.ExitCode);
    }

    [Fact]
    public void CliError_DefaultsToGeneral()
    {
        var e = new CliError("메시지", "힌트");
        Assert.False(e.Environment);
        var env = new CliError("메시지", "힌트", environment: true);
        Assert.True(env.Environment);
    }

    [Fact]
    public void Envelope_Session_NullByDefault_ValueWhenAutoStarted()
    {
        // 평시(채택/기존 세션)엔 null — 출력 안정성. auto-start 시에만 표식이 실린다
        var plain = new Outcome { Command = "push" };
        string j1 = System.Text.Json.JsonSerializer.Serialize(
            new { session = plain.Session }, Output.JsonOptions);
        Assert.Contains("\"session\": null", j1);

        var started = new Outcome { Command = "push" };
        started.Session = new { action = "started", pid = 123, workbook = "C:\\a\\b.xlsm" };
        string j2 = System.Text.Json.JsonSerializer.Serialize(
            new { session = started.Session }, Output.JsonOptions);
        Assert.Contains("\"action\": \"started\"", j2);
        Assert.Contains("\"pid\": 123", j2);
    }
}
