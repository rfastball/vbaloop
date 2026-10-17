using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ProjectConfigTests
{
    private static ProjectConfig LoadToml(string toml)
    {
        string dir = Path.Combine(Path.GetTempPath(), "vbaloop-cfgtest", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, "vbaloop.toml");
        File.WriteAllText(path, toml);
        try { return ProjectConfig.Load(path, dir); }
        finally { try { Directory.Delete(dir, true); } catch { } }
    }

    [Fact]
    public void DistSection_Parsed()
    {
        var cfg = LoadToml("""
            [excel]
            workbook = "Book.xlsm"

            [dist]
            version = "1.2.0"
            exclude = ["Test_*", "Smoke*"]
            bootstrap = "modBootstrap.Bootstrap"
            stripComments = true
            """);
        Assert.Equal("1.2.0", cfg.DistVersion);
        Assert.Equal(new[] { "Test_*", "Smoke*" }, cfg.DistExclude);
        Assert.Equal("modBootstrap.Bootstrap", cfg.DistBootstrap);
        Assert.True(cfg.DistStripComments);
    }

    [Fact]
    public void DistSection_Absent_Defaults()
    {
        var cfg = LoadToml("[excel]\nworkbook = \"Book.xlsm\"");
        Assert.Null(cfg.DistVersion);
        Assert.Empty(cfg.DistExclude);
        Assert.Null(cfg.DistBootstrap);
        Assert.False(cfg.DistStripComments);
    }

    [Fact]
    public void UnknownDistKey_FailsLoudly()
    {
        Assert.Throws<CliError>(() => LoadToml("[excel]\nworkbook = \"B.xlsm\"\n[dist]\nbogus = \"x\""));
    }

    [Fact]
    public void BackupSection_Absent_DefaultsToTen()
    {
        var cfg = LoadToml("[excel]\nworkbook = \"Book.xlsm\"");
        Assert.Equal(10, cfg.BackupKeep);
        Assert.Equal(ProjectConfig.DefaultBackupKeep, cfg.BackupKeep);
    }

    [Theory]
    [InlineData("keep = 3", 3)]
    [InlineData("keep = 0", 0)]
    public void BackupKeep_Parsed(string line, int expected)
    {
        var cfg = LoadToml("[excel]\nworkbook = \"Book.xlsm\"\n[backup]\n" + line);
        Assert.Equal(expected, cfg.BackupKeep);
    }

    [Theory]
    [InlineData("keep = -1")]      // 음수
    [InlineData("keep = \"5\"")]   // 문자열 - 조용히 기본값으로 떨어지면 안 됨
    [InlineData("keep = true")]
    [InlineData("bogus = 1")]      // 미지 키
    public void BackupKeep_Invalid_FailsLoudly(string line)
    {
        Assert.Throws<CliError>(() => LoadToml("[excel]\nworkbook = \"Book.xlsm\"\n[backup]\n" + line));
    }

    [Fact]
    public void BackupsDir_UnderDbg()
    {
        var cfg = LoadToml("[excel]\nworkbook = \"Book.xlsm\"");
        Assert.Equal(Path.Combine(cfg.RootDir, "_dbg", "backups"), cfg.BackupsDir);
    }


    [Fact]
    public void SessionMode_Absent_DefaultsToEphemeral()
    {
        var cfg = LoadToml("[excel]\nworkbook = \"Book.xlsm\"");
        Assert.Equal(SessionMode.Ephemeral, cfg.SessionMode);
        Assert.Equal(Path.Combine(cfg.RootDir, "_dbg", "ephemeral"), cfg.EphemeralDir);
    }

    [Theory]
    [InlineData("ephemeral", SessionMode.Ephemeral)]
    [InlineData("persistent", SessionMode.Persistent)]
    [InlineData("Persistent", SessionMode.Persistent)]
    public void SessionMode_Parsed(string value, SessionMode expected)
    {
        var cfg = LoadToml($"[excel]\nworkbook = \"Book.xlsm\"\nsession = \"{value}\"");
        Assert.Equal(expected, cfg.SessionMode);
    }

    [Theory]
    [InlineData("session = \"shared\"")]   // 오타/미지 값 - 조용히 기본값으로 떨어지면 안 됨
    [InlineData("session = true")]
    public void SessionMode_Invalid_FailsLoudly(string line)
    {
        Assert.Throws<CliError>(() => LoadToml("[excel]\nworkbook = \"Book.xlsm\"\n" + line));
    }

    [Fact]
    public void PushCompile_DefaultsOn_AndParses()
    {
        Assert.True(LoadToml("[excel]\nworkbook = \"Book.xlsm\"").PushCompile);
        Assert.False(LoadToml("[excel]\nworkbook = \"Book.xlsm\"\n[push]\ncompile = false").PushCompile);
    }

    [Theory]
    [InlineData("compile = \"no\"")]
    [InlineData("bogus = true")]
    public void PushSection_Invalid_FailsLoudly(string line)
    {
        Assert.Throws<CliError>(() => LoadToml("[excel]\nworkbook = \"Book.xlsm\"\n[push]\n" + line));
    }
}