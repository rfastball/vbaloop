using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// exe 위에 패키지 .txt(와 워크북)를 끌어다 놓은 진입(설계 §13.6). 사람용이라 계획을 보여 주고 Y/N을 물은 뒤
/// 적용하고, 결과 창은 키 입력까지 유지한다. 셸에서 경로만 넘긴 경우(콘솔을 혼자 쓰지 않음)는 확인 없이
/// <c>push --package</c>로 해석한다 — 다른 CLI 경로처럼 비대화형이다.
/// </summary>
public static class PackageDrop
{
    public static bool IsPackagePath(string arg)
        => Path.GetExtension(arg.Trim().Trim('"')).Equals(".txt", StringComparison.OrdinalIgnoreCase);

    public static int Run(string[] args, List<string> dropped)
    {
        bool interactive = ConsoleHost.IsOwnConsole();
        bool json = CliArgs.RawArgsRequestJson(args);
        var o = new Outcome { Command = "push" };
        try
        {
            o = Execute(args, dropped, interactive);
        }
        catch (CliError e)
        {
            o.Fail(e.Message, e.Hint, e.Environment);
        }
        catch (Exception e)
        {
            o.Fail($"internal error: {e.GetType().Name}: {e.Message}");
        }
        Output.Print(o, json && !interactive);
        if (interactive) ConsoleHost.PauseIfOwnConsole();
        if (!ComCleanup.Run(TimeSpan.FromSeconds(5)))
            ComCleanup.HardExit(o.ExitCode);
        return o.ExitCode;
    }

    private static Outcome Execute(string[] args, List<string> dropped, bool interactive)
    {
        var packages = dropped.Where(IsPackagePath).ToList();
        var workbooks = dropped.Where(p => !IsPackagePath(p)).ToList();
        if (packages.Count > 1 || workbooks.Count > 1)
            throw new CliError($"drop one package .txt and at most one workbook ({packages.Count} packages, {workbooks.Count} other files received)",
                "Or run vbaloop push --package <file.txt> --workbook <file.xlsm>");

        string? workbook = workbooks.FirstOrDefault();
        if (workbook == null)
        {
            if (!interactive)
                throw new CliError("no workbook given for the package", "Run vbaloop push --package <file.txt> --workbook <file.xlsm>");
            workbook = FilePicker.PickWorkbook("Select the workbook to update with this package")
                ?? throw new CliError("no workbook selected — nothing changed");
        }

        // 위치 인자는 빼고 뒤따르는 옵션(--plan, --json 등)은 그대로 넘긴다
        var rest = args.Where(a => !dropped.Contains(a)).ToList();
        string[] Argv(bool adopt) => new[] { "--package", packages[0], "--workbook", workbook }
            .Concat(rest).Concat(adopt ? new[] { "--adopt" } : Array.Empty<string>()).ToArray();

        if (!interactive)
            return PackageApplyCommand.Run(CliArgs.Parse(Argv(false)));

        var o = PackageApplyCommand.Run(CliArgs.Parse(Argv(false)), Confirm);
        // 사람 진입만: dist로 만든 적 없는 업무 파일의 첫 적용은 채택 여부를 직접 묻는다
        if (!o.Ok && o.Hint == PackageVerdict.AdoptHint)
        {
            Flush(o);
            Console.WriteLine($"  {o.Errors.FirstOrDefault()}");
            if (!ConsoleHost.Confirm("This workbook was never updated by a vbaloop package. Use it as this package's target?", defaultYes: false, fallback: false))
            {
                var cancelled = new Outcome { Command = "push" };
                cancelled.HumanLines.Add("cancelled — nothing changed");
                return cancelled;
            }
            o = PackageApplyCommand.Run(CliArgs.Parse(Argv(true)), Confirm);
        }
        return o;
    }

    /// <summary>계획 요약을 지금 출력하고(결과 출력과 겹치지 않게 비운다) 적용 여부를 묻는다.</summary>
    private static bool Confirm(Outcome plan)
    {
        Flush(plan);
        return ConsoleHost.Confirm("Apply this package now?", defaultYes: false, fallback: false);
    }

    private static void Flush(Outcome o)
    {
        foreach (var line in o.HumanLines) Console.WriteLine(line);
        o.HumanLines.Clear();
    }
}
