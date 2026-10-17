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
}
