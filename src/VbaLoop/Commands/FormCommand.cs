using System.Diagnostics;
using System.Text;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

public static class FormCommand
{
    public static Outcome Run(CliArgs opts)
    {
        if (opts.Positional.Count != 1 || opts.Positional[0] != "build" || opts.Spec == null || opts.Code == null)
            throw new CliError("form build requires --spec <design.json> and --code <body.vba>",
                "vbaloop form build --spec forms/MyForm.json --code forms/MyForm.vba --json; then vbaloop push --json");
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        int codePage = NativeMethods.GetACP();
        var definition = FormDefinition.Parse(ReadUtf8(opts.Spec), ReadUtf8(opts.Code), codePage);
        foreach (string ext in new[] { ".frm", ".frx" })
        {
            string output = Path.Combine(cfg.SrcRoot, definition.Form.Name + ext);
            if (PathDrift.SamePath(Path.GetFullPath(opts.Spec), output) || PathDrift.SamePath(Path.GetFullPath(opts.Code), output))
                throw new CliError("form inputs must be separate from generated frm/frx output files");
        }
        using var cmdLock = CommandMutex.Acquire(cfg);
        FormSource.CheckDestination(cfg.SrcRoot, definition.Form.Name, opts.Force);
        var log = new ToolLog(cfg.ToolLogPath, "form");
        LeaseService.AcquireOrRenew(cfg, "form build", log);
        string stage = Path.Combine(cfg.DbgDir, "form-build", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stage);
        var filter = OleMessageFilter.Register();
        ExcelLauncher.LaunchResult? launch = null;
        bool published = false;
        try
        {
            launch = ExcelLauncher.CreateNew(visible: false);
            dynamic app = launch.App;
            app.DisplayAlerts = false;
            app.EnableEvents = false;
            app.AutomationSecurity = 3;
            dynamic workbook = app.Workbooks.Add();
            dynamic project = SessionService.GetVbProject(workbook);
            FormBuilder.Export(project, definition, stage, codePage);
            workbook.Close(false);
            FormSource.Publish(stage, cfg.SrcRoot, definition.Form.Name, opts.Force);
            published = true;
            var outcome = new Outcome { Command = "form", LogRecords = log.Records };
            string frm = Path.Combine(cfg.SrcRoot, definition.Form.Name + ".frm");
            outcome.Payload = new { action = "build", name = definition.Form.Name, frm, frx = Path.ChangeExtension(frm, ".frx"), next = "vbaloop push --json" };
            outcome.HumanLines.Add($"Built {definition.Form.Name} -> {frm}; apply with vbaloop push");
            log.Write("info", "form_built", new { definition.Form.Name, frm });
            return outcome;
        }
        finally
        {
            if (launch != null)
            {
                try { launch.App.Quit(); } catch { /* Owned process cleanup below. */ }
                if (ProcessIdentity.IsAlive(launch.Pid, launch.StartTimeIso))
                    try { Process.GetProcessById(launch.Pid).Kill(entireProcessTree: true); } catch { /* doctor can report leftovers */ }
            }
            filter.Revoke();
            // Keep failed staging/backup evidence for source recovery.
            if (published) try { Directory.Delete(stage, recursive: true); } catch { /* harmless staging residue */ }
        }
    }

    private static string ReadUtf8(string path)
    {
        try { return new UTF8Encoding(false, true).GetString(File.ReadAllBytes(path)).TrimStart('\uFEFF'); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or DecoderFallbackException)
        { throw new CliError($"cannot read UTF-8 input '{path}': {ex.Message}"); }
    }
}
