namespace VbaLoop.Core;

/// <summary>Publish a validated pair under the project command lock. Keep backups if restoration fails.</summary>
public static class FormSource
{
    public static void CheckDestination(string srcRoot, string name, bool force)
    {
        foreach (string ext in new[] { ".bas", ".cls" })
            if (File.Exists(Path.Combine(srcRoot, name + ext)))
                throw new CliError($"{name}: source component already exists as {ext}");
        if (!force && new[] { ".frm", ".frx" }.Any(ext => File.Exists(Path.Combine(srcRoot, name + ext))))
            throw new CliError($"{name}: form source already exists", "Use --force to regenerate both files from JSON and VBA (overwrites source edits)");
    }

    public static void Publish(string stage, string srcRoot, string name, bool force)
    {
        CheckDestination(srcRoot, name, force);
        string[] extensions = { ".frm", ".frx" };
        foreach (string ext in extensions)
            if (!File.Exists(Path.Combine(stage, name + ext))) throw new CliError($"generated form is missing {ext}");
        Directory.CreateDirectory(srcRoot);
        var written = new List<string>();
        foreach (string ext in extensions)
        {
            string dest = Path.Combine(srcRoot, name + ext);
            if (File.Exists(dest)) File.Copy(dest, Path.Combine(stage, name + ext + ".backup"));
        }
        try
        {
            foreach (string ext in extensions)
            {
                File.Move(Path.Combine(stage, name + ext), Path.Combine(srcRoot, name + ext), overwrite: force);
                written.Add(ext);
            }
        }
        catch (Exception error)
        {
            try
            {
                foreach (string ext in written)
                {
                    string backup = Path.Combine(stage, name + ext + ".backup"), dest = Path.Combine(srcRoot, name + ext);
                    if (File.Exists(backup)) File.Copy(backup, dest, overwrite: true);
                    else File.Delete(dest);
                }
            }
            catch (Exception rollback)
            {
                throw new CliError($"form publication failed ({error.Message}); source restoration failed ({rollback.Message})",
                    $"Do not push; restore the source pair using backups in {stage}");
            }
            throw new CliError($"form publication failed; previous source restored: {error.Message}");
        }
    }
}
