using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class FingerprintTests
{
    [Fact]
    public void SameContent_SameFingerprint_ChangeDetected()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vbaloop-test", Path.GetRandomFileName());
        Directory.CreateDirectory(dir);
        try
        {
            File.WriteAllText(Path.Combine(dir, "modA.bas"), "Attribute VB_Name = \"modA\"\r\n");
            File.WriteAllText(Path.Combine(dir, "clsB.cls"), "VERSION 1.0 CLASS\r\n");
            File.WriteAllText(Path.Combine(dir, "readme.md"), "무시되는 파일");

            var fp1 = Fingerprint.Compute(dir);
            var fp2 = Fingerprint.Compute(dir);
            Assert.Equal(fp1, fp2);

            File.WriteAllText(Path.Combine(dir, "readme.md"), "소스 외 파일 변경은 지문 불변");
            Assert.Equal(fp1, Fingerprint.Compute(dir));

            File.AppendAllText(Path.Combine(dir, "modA.bas"), "Option Explicit\r\n");
            Assert.NotEqual(fp1, Fingerprint.Compute(dir));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void MissingDir_StableEmptyFingerprint()
    {
        var fp = Fingerprint.Compute(Path.Combine(Path.GetTempPath(), "vbaloop-none-" + Guid.NewGuid()));
        Assert.False(string.IsNullOrEmpty(fp));
    }
}
