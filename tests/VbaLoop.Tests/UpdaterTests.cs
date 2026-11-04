using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.IO.Compression;
using VbaLoop.Core;
using VbaLoop.Commands;
using VbaLoop.Updater;
using Xunit;

namespace VbaLoop.Tests;

// Test delta: 새 GUI/CLI 경계만 소유. 기존 Package* 판정 매트릭스는 복제하지 않는다.
// 작은 임시 파일만 사용하고 Excel/자식 프로세스/빌드 추가 기동은 없다.
public sealed class UpdaterTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "vbaloop-updater-test-" + Guid.NewGuid().ToString("N"));
    public UpdaterTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, recursive: true);

    // Test delta: 중단 후 회수/개인 경로와 예외 본문 제외/용량 제한은 기존 테스트가 검증하지 않는다.
    // 기존 GUI 경계 소유자에 작은 파일 기반 3건 추가. Excel/자식 프로세스/CI lane 추가 없음.
    // diagnostics.2: 같은 사례에 저장 오류 본문의 선택적 보존/경로 제거를 확장. 수집 사례 수 +0.
    [Fact]
    public void Diagnostics_PersistPhases_RedactPaths_AndExcludeExceptionMessages()
    {
        string workbook = Path.Combine(_root, "confidential-book.xlsb");
        var log = new UpdaterDiagnostics(_root, "worker", workbook, Path.Combine(_root, "kind"));
        var session = new PackageUiSession(null, null, log);
        session.Report("saving");
        log.Write("dialog", new { title = workbook + " confidential-book.xlsb" });
        log.Error("failure", new InvalidOperationException("SECRET VBA OR CELL VALUE"));
        log.Error("save_error", new System.Runtime.InteropServices.COMException("Sharing violation: " + workbook, unchecked((int)0x800A03EC)), includeMessage: true);
        session.Report("saved");
        string text = File.ReadAllText(Path.Combine(_root, "worker.jsonl"));
        Assert.DoesNotContain("confidential-book", text);
        Assert.DoesNotContain("SECRET VBA OR CELL VALUE", text);
        Assert.Contains("Sharing violation: [redacted]", text);
        Assert.Contains("0x800A03EC", text);
        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries).Select(l => JsonDocument.Parse(l)).ToArray();
        try
        {
            var phases = lines.Where(l => l.RootElement.GetProperty("kind").GetString() == "phase").ToArray();
            Assert.Equal(new[] { "unknown", "saved" }, phases.Select(l => l.RootElement.GetProperty("data").GetProperty("saveState").GetString()));
            Assert.All(lines, l => Assert.True(l.RootElement.TryGetProperty("utc", out _)));
        }
        finally { foreach (var line in lines) line.Dispose(); }
    }

    [Fact]
    public void Diagnostics_LimitGrowth_AndDoNotInterruptWorkOnWriteFailure()
    {
        var log = new UpdaterDiagnostics(_root, "worker");
        string path = Path.Combine(_root, "worker.jsonl");
        using (var file = File.OpenWrite(path)) file.SetLength(UpdaterDiagnostics.MaxBytes - 1024);
        log.Write("next");
        long length = new FileInfo(path).Length;
        Assert.InRange(length, UpdaterDiagnostics.MaxBytes - 1023, UpdaterDiagnostics.MaxBytes);
        log.Write("ignored");
        Assert.Equal(length, new FileInfo(path).Length);
        var broken = new UpdaterDiagnostics(_root, "broken");
        string brokenPath = Path.Combine(_root, "broken.jsonl");
        File.Delete(brokenPath);
        Directory.CreateDirectory(brokenPath);
        broken.Write("cannot_write"); // 로그 장애를 원래 작업에 전파하지 않는다.
    }

    [Fact]
    public void DiagnosticExport_IncludesOnlyLogs_AndWorksWithActiveWriter()
    {
        string folder = Path.Combine(_root, "session", "apply");
        var log = new UpdaterDiagnostics(folder, "worker");
        File.WriteAllText(Path.Combine(folder, "private.xlsb"), "must not export workbook");
        using var active = new FileStream(Path.Combine(folder, "worker.jsonl"), FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
        string zip = Path.Combine(_root, "diagnostics.zip");
        CliClient.ExportFolder(Path.GetDirectoryName(folder)!, zip);
        using var archive = ZipFile.OpenRead(zip);
        Assert.Equal("apply/worker.jsonl", Assert.Single(archive.Entries).FullName);
        using var reader = new StreamReader(archive.Entries[0].Open());
        Assert.Contains("started", reader.ReadToEnd());
    }

    // Test delta: 공유 쓰기를 허용하는 관찰용 읽기 핸들의 오탐과 실제 쓰기 잠금 차단을
    // 기존 업데이터 경계에서 검증. 임시 파일 3개만 추가, Excel/프로세스 기동 없음.
    [Theory]
    [InlineData(FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, false)]
    [InlineData(FileAccess.Read, FileShare.Read, true)]
    [InlineData(FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete, true)]
    public void WritableCheck_AllowsSharedReaders_ButRefusesWritersAndWriteDenial(
        FileAccess access, FileShare sharing, bool blocked)
    {
        string workbook = Path.Combine(_root, "book.xlsm");
        File.WriteAllText(workbook, "unchanged");
        PackageApplyCommand.CheckWritable(workbook);
        using (var holder = new FileStream(workbook, FileMode.Open, access, sharing))
        {
            if (blocked)
                Assert.Equal("file_locked", Assert.Throws<CliError>(() => PackageApplyCommand.CheckWritable(workbook)).Code);
            else
                PackageApplyCommand.CheckWritable(workbook);
        }
        Assert.Equal("unchanged", File.ReadAllText(workbook));
    }

    [Fact]
    public void Bundle_RepairsCorruptCache_AndRejectsCorruptEmbeddedBytes()
    {
        byte[] bytes = Encoding.UTF8.GetBytes("representative executable bytes");
        var manifest = new EngineManifest("1.0.0", Convert.ToHexString(SHA256.HashData(bytes)), bytes.Length, 1);
        string path = EngineBundle.ExtractVerified(new MemoryStream(bytes), manifest, _root);
        Assert.Equal(bytes, File.ReadAllBytes(path));
        // 길이가 같아도 내용이 다르면 실행하지 않고 원본으로 복구한다.
        File.WriteAllBytes(path, new byte[bytes.Length]);
        Assert.Equal(path, EngineBundle.ExtractVerified(new MemoryStream(bytes), manifest, _root));
        Assert.Equal(bytes, File.ReadAllBytes(path));
        var badManifest = manifest with { Sha256 = new string('A', 64) };
        Assert.Throws<InvalidDataException>(() => EngineBundle.ExtractVerified(new MemoryStream(bytes), badManifest, _root));
        Assert.Empty(Directory.GetFiles(_root, "*.tmp", SearchOption.AllDirectories));
    }

    [Theory]
    [InlineData("not_saved", false, true, false)]
    [InlineData("unknown", false, false, false)]
    [InlineData("saved", true, false, true)]
    [InlineData("saved", false, false, false)]
    public void Response_DoesNotConfuseInspectionOrUncertainSaveWithSuccess(string state, bool saved, bool canApply, bool applied)
    {
        string json = JsonSerializer.Serialize(new { protocol = 1, command = "package", ok = true, exitCode = 0,
            saveState = state, packageHash = new string('A', 64), workbookHash = new string('B', 64),
            payload = new { decision = "applied", saved } });
        var result = PackageResponse.Parse(json, 0);
        Assert.Equal(canApply, result.CanApply);
        Assert.Equal(applied, result.Applied);
        Assert.Throws<InvalidDataException>(() => PackageResponse.Parse(json, 3));
    }

    [Fact]
    public void ApplyArguments_RequireInspection_AndKeepPathsAsSingleArguments()
    {
        string package = Path.Combine(_root, "지출 관리 & $(file).txt");
        string workbook = Path.Combine(_root, "업무 파일.xlsm");
        Assert.Throws<InvalidOperationException>(() => CliClient.BuildArguments(package, workbook, false, false, null));
        var inspected = new PackageResponse { Ok = true, SaveState = "not_saved", PackageHash = new string('A', 64),
            WorkbookHash = new string('B', 64), Payload = new PackageInfo { Decision = "applied" } };
        var args = CliClient.BuildArguments(package, workbook, false, true, inspected);
        Assert.Equal(package, args[Array.IndexOf(args, "--package") + 1]);
        Assert.Equal(workbook, args[Array.IndexOf(args, "--workbook") + 1]);
        Assert.Contains("--adopt", args);
        var parsed = VbaLoop.Cli.CliArgs.Parse(args.Skip(1).ToArray());
        Assert.Equal(inspected.PackageHash, parsed.ExpectedPackageHash);
        Assert.Equal(inspected.WorkbookHash, parsed.ExpectedWorkbookHash);
    }

    [Fact]
    public void InspectionHashes_RejectChangedPackageAndWorkbook()
    {
        string workbook = Path.Combine(_root, "book.xlsm");
        File.WriteAllText(workbook, "original workbook");
        byte[] package = [1, 2, 3];
        var inspect = new PackageUiSession(null, null);
        inspect.Verify(package, workbook);
        var apply = new PackageUiSession(inspect.PackageHash, inspect.WorkbookHash);
        apply.Verify(package, workbook);
        Assert.Equal("files_changed", Assert.Throws<CliError>(() => apply.Verify([1, 2, 4], workbook)).Code);
        File.WriteAllText(workbook, "changed workbook");
        Assert.Equal("files_changed", Assert.Throws<CliError>(() => apply.Verify(package, workbook)).Code);
        Assert.Equal("files_changed", Assert.Throws<CliError>(() => inspect.VerifyOpenedWorkbook(workbook)).Code);
    }

    [Fact]
    public void GuiCompileGate_RequiresPositiveConfirmation_WithoutChangingCliPolicy()
    {
        Assert.True(PackageUiSession.CanSave(DistCompileVerdict.Kind.Compiled));
        Assert.True(PackageUiSession.CanSave(DistCompileVerdict.Kind.AlreadyCompiled));
        Assert.False(PackageUiSession.CanSave(DistCompileVerdict.Kind.Unverified));
        Assert.True(new DistCompileVerdict.Verdict(DistCompileVerdict.Kind.Unverified, null).Publishable);
    }

    [Fact]
    public void TimeoutCannotRaceIntoSave_AndNeverTerminatesSaving()
    {
        bool expired = false;
        var beforeSave = new PackageUiSession(null, null);
        beforeSave.CheckTimeout(DateTime.UtcNow.AddMinutes(3), () => expired = true);
        Assert.True(expired);
        Assert.Throws<CliError>(() => beforeSave.Report("saving"));
        var saving = new PackageUiSession(null, null);
        saving.Report("saving");
        saving.CheckTimeout(DateTime.UtcNow.AddMinutes(3), () => Assert.Fail("Must not kill during Save"));
        Assert.Equal("unknown", saving.SaveState);
        saving.Report("saved");
        Assert.Equal("saved", saving.SaveState);
    }
}
