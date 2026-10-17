using System.Text.Json;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class EnvelopeTests
{
    private static readonly string[] BaseKeys = { "command", "ok", "exitCode", "hint", "session", "errors", "payload" };

    [Fact]
    public void WithoutUpdate_KeysAndOrderUnchanged()
    {
        var o = new Outcome { Command = "status" };
        using var doc = JsonDocument.Parse(Output.Serialize(o));
        Assert.Equal(BaseKeys, doc.RootElement.EnumerateObject().Select(p => p.Name).ToArray());
        Assert.Equal(JsonValueKind.Null, doc.RootElement.GetProperty("session").ValueKind); // 기존 null 필드는 유지
        Assert.Equal(JsonValueKind.Null, doc.RootElement.GetProperty("hint").ValueKind);
    }

    [Fact]
    public void WithUpdate_AppendsOptionalField_HintUntouched()
    {
        var o = new Outcome { Command = "doctor", Update = new UpdateNotice("0.3.0", "0.4.0", "vbaloop update") };
        o.Fail("x", "기존 hint");
        using var doc = JsonDocument.Parse(Output.Serialize(o));
        var root = doc.RootElement;
        Assert.Equal(BaseKeys.Append("update").ToArray(), root.EnumerateObject().Select(p => p.Name).ToArray());
        var u = root.GetProperty("update");
        Assert.Equal("0.3.0", u.GetProperty("current").GetString());
        Assert.Equal("0.4.0", u.GetProperty("latest").GetString());
        Assert.Equal("vbaloop update", u.GetProperty("command").GetString());
        Assert.Equal("기존 hint", root.GetProperty("hint").GetString());
        Assert.Equal(2, root.GetProperty("exitCode").GetInt32());
    }
}
