using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ShellQuoteTests
{
    [Theory]
    [InlineData("--json", "--json")]
    [InlineData(@"C:\work\Book.xlsm", @"C:\work\Book.xlsm")]
    [InlineData(@"C:\my work\정산.xlsm", @"'C:\my work\정산.xlsm'")]
    [InlineData(@"C:\$(calc)\b.xlsm", @"'C:\$(calc)\b.xlsm'")]
    [InlineData("a`b", "'a`b'")]
    [InlineData("it's", "'it''s'")]
    [InlineData("it\u2019s", "'it\u2019\u2019s'")] // PowerShell은 둥근 따옴표도 작은따옴표로 본다
    [InlineData("", "''")]
    public void PowerShell_QuotesOnlyWhenNeeded(string arg, string expected)
        => Assert.Equal(expected, ShellQuote.PowerShell(arg));

    [Fact]
    public void PowerShellCommand_JoinsQuotedArgs()
        => Assert.Equal(@"vbaloop setup --folder 'C:\a b' --json",
            ShellQuote.PowerShellCommand("vbaloop", new[] { "setup", "--folder", @"C:\a b", "--json" }));
}
