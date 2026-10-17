using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class SyncCacheTests
{
    private static string TempPath() =>
        Path.Combine(Path.GetTempPath(), "vbaloop-test", Path.GetRandomFileName(), "sync.json");

    [Fact]
    public void Roundtrip_PreservesFieldsAndForms()
    {
        var path = TempPath();
        try
        {
            var cache = new SyncCache
            {
                LastPushFingerprint = "ABC123",
                LastPushAt = "2026-07-08T00:00:00.0000000Z",
                LastPullFingerprint = "DEF456",
                LastPullAt = "2026-07-08T01:00:00.0000000Z",
            };
            cache.Forms["frmFilter01"] = new SyncCache.FormSync { FrmHash = "F1", FrxHash = "X1" };
            cache.Save(path);

            var loaded = SyncCache.Load(path);
            Assert.Equal("ABC123", loaded.LastPushFingerprint);
            Assert.Equal("2026-07-08T00:00:00.0000000Z", loaded.LastPushAt);
            Assert.Equal("DEF456", loaded.LastPullFingerprint);
            Assert.Equal("2026-07-08T01:00:00.0000000Z", loaded.LastPullAt);
            Assert.Equal("X1", loaded.Forms["frmFilter01"].FrxHash);
            // 폼 이름 조회는 대소문자 무시 (VBA 이름 규칙)
            Assert.Equal("F1", loaded.Forms["FRMFILTER01"].FrmHash);
        }
        finally
        {
            Directory.Delete(Path.GetDirectoryName(path)!, recursive: true);
        }
    }

    [Fact]
    public void Load_MissingOrCorrupt_ReturnsEmptyCache()
    {
        var missing = SyncCache.Load(TempPath());
        Assert.Null(missing.LastPushFingerprint);
        Assert.Empty(missing.Forms);

        var path = TempPath();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, "{ 깨진 json ");
            var corrupt = SyncCache.Load(path);
            Assert.Null(corrupt.LastPushFingerprint);
            Assert.Empty(corrupt.Forms);
        }
        finally
        {
            Directory.Delete(Path.GetDirectoryName(path)!, recursive: true);
        }
    }

    [Fact]
    public void SessionMeta_Load_IgnoresLegacyPushFields()
    {
        // 구버전 session.json에 남은 lastPush* 필드는 무시되고 로드는 성공해야 한다
        var path = Path.Combine(Path.GetTempPath(), "vbaloop-test-" + Guid.NewGuid().ToString("N") + ".json");
        try
        {
            File.WriteAllText(path, """
                {
                  "pid": 123,
                  "excelStartTime": "2026-07-08T00:00:00.0000000Z",
                  "hwnd": 456,
                  "workbookFullName": "C:\\x\\Book.xlsm",
                  "createdAt": "2026-07-08T00:00:00.0000000Z",
                  "lastPushFingerprint": "LEGACY",
                  "lastPushAt": "2026-07-08T00:00:00.0000000Z",
                  "poisoned": false
                }
                """);
            var meta = SessionMeta.Load(path);
            Assert.NotNull(meta);
            Assert.Equal(123, meta!.Pid);
            Assert.Equal(456, meta.Hwnd);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
