using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class SelfInstallTests
{
    private const string Dir = @"C:\Users\me\AppData\Local\vbaloop\bin";

    [Fact]
    public void AddToPathValue_Empty_ReturnsDirOnly()
    {
        var (path, changed) = SelfInstall.AddToPathValue("", Dir);
        Assert.True(changed);
        Assert.Equal(Dir, path);
    }

    [Fact]
    public void AddToPathValue_Null_ReturnsDirOnly()
    {
        var (path, changed) = SelfInstall.AddToPathValue(null, Dir);
        Assert.True(changed);
        Assert.Equal(Dir, path);
    }

    [Fact]
    public void AddToPathValue_Appends_PreservingExistingEntriesVerbatim()
    {
        // %USERPROFILE% 같은 미확장 토큰이 원문 그대로 살아남아야 한다(ExpandString 기록 전제)
        const string raw = @"C:\Windows;%USERPROFILE%\bin";
        var (path, changed) = SelfInstall.AddToPathValue(raw, Dir);
        Assert.True(changed);
        Assert.Equal(raw + ";" + Dir, path);
    }

    [Theory]
    [InlineData(@"C:\Windows;C:\Users\me\AppData\Local\vbaloop\bin")]
    [InlineData(@"C:\Users\me\AppData\Local\vbaloop\bin\;C:\Windows")]   // 트레일링 백슬래시
    [InlineData(@"C:\USERS\ME\APPDATA\LOCAL\VBALOOP\BIN;C:\Windows")]     // 대소문자
    [InlineData(@"C:\Windows; C:\Users\me\AppData\Local\vbaloop\bin ")]   // 주변 공백
    public void AddToPathValue_AlreadyPresent_NoChange(string raw)
    {
        var (path, changed) = SelfInstall.AddToPathValue(raw, Dir);
        Assert.False(changed);
        Assert.Equal(raw, path);
    }

    [Fact]
    public void AddToPathValue_DropsEmptySegments_ButKeepsOrder()
    {
        var (path, changed) = SelfInstall.AddToPathValue(@"C:\a;;C:\b;", Dir);
        Assert.True(changed);
        Assert.Equal(@"C:\a;C:\b;" + Dir, path);
    }
}
