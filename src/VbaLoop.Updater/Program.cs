using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;

namespace VbaLoop.Updater;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 2 && args[0] == "--self-test") return SelfTestAsync(args[1]).GetAwaiter().GetResult();
        var app = new Application();
        app.DispatcherUnhandledException += (_, e) =>
        {
            MessageBox.Show(e.Exception.Message, "vbaloop 업데이트", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        };
        return app.Run(new MainWindow(args));
    }

    private static async Task<int> SelfTestAsync(string reportPath)
    {
        try
        {
            reportPath = Path.GetFullPath(reportPath);
            Directory.CreateDirectory(Path.GetDirectoryName(reportPath)!);
            var bundle = EngineBundle.Extract(Path.Combine(Path.GetDirectoryName(reportPath)!, "smoke-engine"));
            var version = FileVersionInfo.GetVersionInfo(bundle.Path).ProductVersion?.Split('+')[0];
            if (version != bundle.Manifest.Version) throw new InvalidDataException("Embedded CLI version mismatch");
            // help와 잘못된 package 요청은 Excel·네트워크·프로젝트를 사용하지 않는다.
            var help = await CliClient.ExecuteAsync(bundle.Path, ["help", "--json"]);
            using var json = JsonDocument.Parse(help.Output);
            if (help.ExitCode != 0 || !json.RootElement.GetProperty("ok").GetBoolean())
                throw new InvalidDataException("Embedded CLI did not start");
            var invalid = await CliClient.ExecuteAsync(bundle.Path, ["__package-ui", "--json"]);
            var response = PackageResponse.Parse(invalid.Output, invalid.ExitCode);
            if (response.ErrorCode != "invalid_request") throw new InvalidDataException("GUI protocol mismatch");
            File.WriteAllText(reportPath, JsonSerializer.Serialize(new { ok = true, version, cliSha256 = bundle.Manifest.Sha256 }));
            return 0;
        }
        catch (Exception ex)
        {
            File.WriteAllText(reportPath, JsonSerializer.Serialize(new { ok = false, error = ex.ToString() }));
            return 1;
        }
    }
}
