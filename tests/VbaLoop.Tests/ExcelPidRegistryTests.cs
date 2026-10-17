using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ExcelPidRegistryTests : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), "vbaloop-pids-" + Guid.NewGuid().ToString("N"), "excel-pids.json");

    public void Dispose()
    {
        try { Directory.Delete(Path.GetDirectoryName(_path)!, recursive: true); } catch { /* best effort */ }
    }

    private static bool NeverGone(ExcelPidRegistry.Entry _) => false;

    [Fact]
    public void Missing_LoadsEmpty()
    {
        Assert.Empty(ExcelPidRegistry.Load(_path));
    }

    [Fact]
    public void Record_ThenHas_ByPidStartAndKind()
    {
        ExcelPidRegistry.TryRecord(_path, 10, "t1", ExcelPidRegistry.KindLaunched, null, NeverGone);
        ExcelPidRegistry.TryRecord(_path, 10, "t1", ExcelPidRegistry.KindQuarantined, @"C:\p", NeverGone);
        var e = ExcelPidRegistry.Load(_path);

        Assert.True(ExcelPidRegistry.Has(e, 10, "t1", ExcelPidRegistry.KindLaunched));
        Assert.True(ExcelPidRegistry.Has(e, 10, "t1", ExcelPidRegistry.KindQuarantined));
        // PID 재사용(다른 시작 시각)은 같은 프로세스가 아니다
        Assert.False(ExcelPidRegistry.Has(e, 10, "t2", ExcelPidRegistry.KindLaunched));
        Assert.False(ExcelPidRegistry.Has(e, 10, null, ExcelPidRegistry.KindLaunched));
    }

    [Fact]
    public void Record_IsIdempotent_PerKey()
    {
        ExcelPidRegistry.TryRecord(_path, 10, "t1", ExcelPidRegistry.KindLaunched, null, NeverGone);
        ExcelPidRegistry.TryRecord(_path, 10, "t1", ExcelPidRegistry.KindLaunched, null, NeverGone);
        Assert.Single(ExcelPidRegistry.Load(_path));
    }

    [Fact]
    public void Record_PrunesGoneProcesses()
    {
        ExcelPidRegistry.TryRecord(_path, 10, "t1", ExcelPidRegistry.KindLaunched, null, NeverGone);
        ExcelPidRegistry.TryRecord(_path, 11, "t2", ExcelPidRegistry.KindLaunched, null, e => e.Pid == 10);
        var e = ExcelPidRegistry.Load(_path);
        Assert.Single(e);
        Assert.Equal(11, e[0].Pid);
    }

    [Fact]
    public void CorruptFile_LoadsEmpty_AndIsOverwrittenOnRecord()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
        File.WriteAllText(_path, "{{{");
        Assert.Empty(ExcelPidRegistry.Load(_path));
        ExcelPidRegistry.TryRecord(_path, 7, "t", ExcelPidRegistry.KindLaunched, null, NeverGone);
        Assert.Single(ExcelPidRegistry.Load(_path));
    }
}
