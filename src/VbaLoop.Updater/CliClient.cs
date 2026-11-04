using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using VbaLoop.Core;

namespace VbaLoop.Updater;

public sealed class CliClient
{
    public string DiagnosticFolder { get; } = Path.Combine(UpdaterDiagnostics.Root, DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N")[..8]);

    public void ExportDiagnostics(string destination)
    {
        // 재시작한 뒤에도 직전 실행 로그를 회수한다. 사용자가 선택한 ZIP에만 복사한다.
        string folder = Directory.Exists(DiagnosticFolder) ? DiagnosticFolder :
            Directory.Exists(UpdaterDiagnostics.Root) ? Directory.GetDirectories(UpdaterDiagnostics.Root).OrderByDescending(p => p).FirstOrDefault()
                ?? throw new IOException("아직 기록된 진단 로그가 없습니다.") : throw new IOException("아직 기록된 진단 로그가 없습니다.");
        ExportFolder(folder, destination);
    }

    public static void ExportFolder(string folder, string destination)
    {
        using var output = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None);
        using var archive = new ZipArchive(output, ZipArchiveMode.Create);
        foreach (string path in Directory.EnumerateFiles(folder, "*.jsonl", SearchOption.AllDirectories))
        {
            using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var entry = archive.CreateEntry(Path.GetRelativePath(folder, path).Replace('\\', '/')).Open();
            // 진행 중에도 내보낼 수 있도록 열었을 때의 길이까지만 복사한다.
            long remaining = input.Length;
            byte[] buffer = new byte[8192];
            while (remaining > 0)
            {
                int count = input.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
                if (count == 0) break;
                entry.Write(buffer, 0, count);
                remaining -= count;
            }
        }
    }

    public async Task<PackageResponse> RunAsync(string package, string workbook, bool plan, bool adopt,
        PackageResponse? inspection, IProgress<string> progress)
    {
        string folder = Path.Combine(DiagnosticFolder, DateTime.Now.ToString("HHmmss-fff") + (plan ? "-inspect-" : "-apply-") + Guid.NewGuid().ToString("N")[..8]);
        var log = new UpdaterDiagnostics(folder, "gui", package, workbook);
        try
        {
            log.Write("request", new { plan, adopt });
            progress.Report("preparing");
            var bundle = await Task.Run(() => EngineBundle.Extract());
            var args = BuildArguments(package, workbook, plan, adopt, inspection);
            var result = await ExecuteAsync(bundle.Path, args, progress, folder, log);
            var response = PackageResponse.Parse(result.Output, result.ExitCode);
            log.Write("response", new { response.Ok, response.ErrorCode, response.SaveState });
            return response;
        }
        catch (Exception ex) { log.Error("gui_error", ex); throw; }
    }

    public static string[] BuildArguments(string package, string workbook, bool plan, bool adopt, PackageResponse? inspection)
    {
        var args = new List<string> { "__package-ui", "--json", "--package", Path.GetFullPath(package), "--workbook", Path.GetFullPath(workbook) };
        if (plan) args.Add("--plan");
        else
        {
            if (inspection?.CanApply != true) throw new InvalidOperationException("먼저 파일을 검사해 주세요.");
            args.AddRange(["--expected-package-hash", inspection.PackageHash!, "--expected-workbook-hash", inspection.WorkbookHash!]);
        }
        if (adopt) args.Add("--adopt");
        return args.ToArray();
    }

    public static async Task<(string Output, int ExitCode)> ExecuteAsync(string exe, IEnumerable<string> args, IProgress<string>? progress = null,
        string? diagnosticFolder = null, UpdaterDiagnostics? log = null)
    {
        var start = new ProcessStartInfo(exe)
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = Path.GetDirectoryName(exe)!,
        };
        foreach (var arg in args) start.ArgumentList.Add(arg);
        if (diagnosticFolder != null) start.Environment[UpdaterDiagnostics.FolderVariable] = diagnosticFolder;
        using var process = new Process { StartInfo = start };
        if (!process.Start()) throw new IOException("실행 도구를 시작할 수 없습니다.");
        log?.Write("worker_started", new { pid = process.Id });
        // 양쪽 파이프를 동시에 비운다. stdout은 최종 결과, stderr는 실제 작업 단계다.
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = ReadProgressAsync(process.StandardError, progress, log);
        await process.WaitForExitAsync();
        await stderr;
        log?.Write("worker_exited", new { process.ExitCode });
        return (await stdout, process.ExitCode);
    }

    private static async Task ReadProgressAsync(StreamReader reader, IProgress<string>? progress, UpdaterDiagnostics? log)
    {
        while (await reader.ReadLineAsync() is { } line)
        {
            try
            {
                using var json = JsonDocument.Parse(line);
                var root = json.RootElement;
                if (root.TryGetProperty("protocol", out var protocol) && protocol.GetInt32() == 1 &&
                    root.TryGetProperty("phase", out var phase) && phase.GetString() is { } value)
                {
                    log?.Write("worker_phase", new { phase = value });
                    progress?.Report(value);
                }
            }
            catch (JsonException) { log?.Write("worker_non_protocol_stderr"); /* 업무 내용을 포함할 수 있어 원문은 기록하지 않는다. */ }
        }
    }
}
