using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;

namespace VbaLoop.Updater;

public sealed class CliClient
{
    public async Task<PackageResponse> RunAsync(string package, string workbook, bool plan, bool adopt,
        PackageResponse? inspection, IProgress<string> progress)
    {
        progress.Report("preparing");
        var bundle = await Task.Run(() => EngineBundle.Extract());
        var args = BuildArguments(package, workbook, plan, adopt, inspection);
        var result = await ExecuteAsync(bundle.Path, args, progress);
        return PackageResponse.Parse(result.Output, result.ExitCode);
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

    public static async Task<(string Output, int ExitCode)> ExecuteAsync(string exe, IEnumerable<string> args, IProgress<string>? progress = null)
    {
        var start = new ProcessStartInfo(exe)
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = Path.GetDirectoryName(exe)!,
        };
        foreach (var arg in args) start.ArgumentList.Add(arg);
        using var process = new Process { StartInfo = start };
        if (!process.Start()) throw new IOException("실행 도구를 시작할 수 없습니다.");
        // 양쪽 파이프를 동시에 비운다. stdout은 최종 결과, stderr는 실제 작업 단계다.
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = ReadProgressAsync(process.StandardError, progress);
        await process.WaitForExitAsync();
        await stderr;
        return (await stdout, process.ExitCode);
    }

    private static async Task ReadProgressAsync(StreamReader reader, IProgress<string>? progress)
    {
        while (await reader.ReadLineAsync() is { } line)
        {
            try
            {
                using var json = JsonDocument.Parse(line);
                var root = json.RootElement;
                if (root.TryGetProperty("protocol", out var protocol) && protocol.GetInt32() == 1 &&
                    root.TryGetProperty("phase", out var phase) && phase.GetString() is { } value)
                    progress?.Report(value);
            }
            catch (JsonException) { /* stderr의 일반 진단 문구는 단계로 해석하지 않는다. */ }
        }
    }
}
