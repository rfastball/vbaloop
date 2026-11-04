using System.Text.Json;
using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Commands;

/// <summary>내장 GUI가 사용하는 로컬 자식 프로세스 진입. 네트워크·설치·프로젝트 설정 없음.</summary>
public static class PackageUiCommand
{
    public static int Run(CliArgs opts)
    {
        UpdaterDiagnostics? diagnostics = null;
        string? diagnosticFolder = Environment.GetEnvironmentVariable(UpdaterDiagnostics.FolderVariable);
        if (!string.IsNullOrWhiteSpace(diagnosticFolder))
        {
            try { diagnostics = new UpdaterDiagnostics(diagnosticFolder, "worker", opts.Package, opts.Workbook); }
            catch (Exception ex) { Console.Error.WriteLine($"diagnostic_unavailable: {ex.GetType().Name} (0x{ex.HResult:X8})"); }
        }
        var session = new PackageUiSession(opts.ExpectedPackageHash, opts.ExpectedWorkbookHash, diagnostics);
        diagnostics?.Write("request", new { plan = opts.Plan, adopt = opts.Adopt });
        if (opts.Workbook != null) diagnostics?.FileState("workbook_before", opts.Workbook);
        using var observer = diagnostics == null ? null : new PackageUiDiagnostics(session, diagnostics, opts.Workbook);
        var outcome = new Outcome { Command = "push" };
        string? errorCode = null;
        using var watchdog = session.Watch(() =>
        {
            WriteResult(false, 3, "operation_timeout", ["Excel operation timed out. Check the workbook and any remaining Excel windows."], null, null);
            ComCleanup.HardExit(3);
        });
        try
        {
            if (opts.Package == null || !opts.Json || (!opts.Plan &&
                (opts.ExpectedPackageHash == null || opts.ExpectedWorkbookHash == null)))
                throw new CliError("GUI apply requires the hashes returned by inspection", code: "invalid_request");
            if (opts.AllowDowngrade || opts.SkipCompile)
                throw new CliError("GUI does not allow downgrade or skipping compilation", code: "invalid_request");
            outcome = PackageApplyCommand.Run(opts, ui: session);
            errorCode = outcome.ErrorCode;
        }
        catch (CliError ex)
        {
            diagnostics?.Error("command_error", ex);
            outcome.Fail(ex.Message, ex.Hint, ex.Environment);
            errorCode = ex.Code;
        }
        catch (Exception ex)
        {
            diagnostics?.Error("unexpected_error", ex);
            outcome.Fail($"{ex.GetType().Name}: {ex.Message}");
            errorCode = "unexpected_error";
        }

        session.Finish(() => WriteResult(outcome.Ok, outcome.ExitCode, errorCode, outcome.Errors, outcome.Hint, outcome.Payload));
        // Apply의 finally가 도구 Excel을 정리했다. 응답 없는 RCW의 STA 해제에 다시 막히지 않는다.
        ComCleanup.HardExit(outcome.ExitCode);
        return outcome.ExitCode;

        void WriteResult(bool ok, int exitCode, string? code, IEnumerable<string> errors, string? hint, object? payload)
        {
            diagnostics?.Write("result", new { ok, exitCode, errorCode = code, phase = session.Phase, saveState = session.SaveState });
            if (opts.Workbook != null) diagnostics?.FileState("workbook_after", opts.Workbook);
            try
            {
                Console.WriteLine(JsonSerializer.Serialize(new
                {
                    protocol = 1, command = "package", ok, exitCode, errorCode = code, errors, hint, payload,
                    saveState = session.SaveState, phase = session.Phase,
                    copyMode = true, outputWorkbook = session.OutputWorkbook,
                    packageHash = session.PackageHash, workbookHash = session.WorkbookHash,
                }, Output.JsonOptions));
            }
            catch (IOException) { /* GUI를 닫은 뒤에도 Excel 저장·정리는 완료한다. */ }
        }
    }
}
