using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop push --package &lt;txt&gt; --workbook &lt;파일&gt; (설계 §13.3): dist 텍스트 패키지로 업무 워크북의 VBA 코드만
/// 교체한다. 프로젝트 없이 돈다 — vbaloop.toml·lease·_dbg\·지문 캐시·레지스트리를 쓰지 않으므로 프로젝트 폴더
/// 안에서 실행해도 그 프로젝트의 동기화 상태는 바뀌지 않는다. Excel은 항상 이 명령 전용의 숨은 인스턴스이고,
/// 실패하면 저장 없이 버린다 — 디스크의 업무 파일은 적용 전 그대로 남는다. 판정·계획은 Core.Package*가 한다.
/// </summary>
public static class PackageApplyCommand
{
    private static readonly string[] MacroExtensions = { ".xlsm", ".xlsb", ".xlam" };

    /// <param name="confirm">드래그 앤 드롭 진입 전용(§13.6): 계획을 보여 준 뒤 적용 여부를 묻는다. CLI는 null(비대화형).</param>
    public static Outcome Run(CliArgs opts, Func<Outcome, bool>? confirm = null, PackageUiSession? ui = null)
    {
        var o = new Outcome { Command = "push" };
        var log = new ToolLog(null, "push");
        o.LogRecords = log.Records;

        // 0. 인자 계약 — 둘은 짝이다. --workbook 기본값을 정본으로 두지 않아 정본·업무 파일 혼동을 막는다
        if (opts.Workbook == null)
            throw new CliError("--package needs --workbook <file> (the business workbook to update)",
                "Run vbaloop push --package <file.txt> --workbook <file.xlsm>");
        if (opts.Config != null)
            throw new CliError("--config cannot be used with --package — package mode does not read a project");
        if (opts.Force)
            throw new CliError("--force cannot be used with --package — there is no fingerprint cache to bypass");
        if (opts.SkipCompile)
            throw new CliError("--skip-compile cannot be used with --package — the whole code is replaced, so a compile error is a package defect",
                "Fix it on the development PC, build the package again, then apply");

        string pkgPath = Path.GetFullPath(opts.Package!.Trim().Trim('"'));
        string wbPath = Path.GetFullPath(opts.Workbook.Trim().Trim('"'));
        if (!File.Exists(pkgPath)) throw new CliError($"package file not found: {pkgPath}");
        if (!File.Exists(wbPath)) throw new CliError($"workbook file not found: {wbPath}");
        string ext = Path.GetExtension(wbPath).ToLowerInvariant();
        if (!MacroExtensions.Contains(ext))
            throw new CliError($"workbook must be a macro-enabled file ({string.Join(", ", MacroExtensions)}): {wbPath}",
                "A .xlsx cannot hold VBA — saving it would drop the code");

        // 1. 패키지 파싱·해시 검증, ACP 가드(Excel 전). 대상 PC의 ACP가 개발 PC와 다를 수 있다
        ui?.Report("checking");
        var packageBytes = File.ReadAllBytes(pkgPath);
        PackageFormat.Package pkg;
        try { pkg = PackageFormat.Parse(packageBytes); }
        catch (CliError ex) { throw new CliError(ex.Message, ex.Hint, code: "invalid_package"); }
        int codePage = NativeMethods.GetACP();
        GuardEncoding(pkg, codePage);

        using var pathLock = PathMutex.Acquire(wbPath);
        FileUseDiagnostics.Capture(ui?.Diagnostics, wbPath, "before_open");
        if (ui == null) CheckWritable(wbPath);
        ui?.Verify(packageBytes, wbPath);
        using var workingCopy = ui == null ? null : new GuiWorkbookCopy(wbPath, ui.WorkbookHash!, opts.Plan);
        if (workingCopy != null)
        {
            ui!.WorkingWorkbook = workingCopy.WorkPath;
            ui.OutputWorkbook = workingCopy.OutputPath;
            ui.Diagnostics?.Write("working_copy_created", new { copyMode = true, outputAllocated = ui.OutputWorkbook != null });
        }

        string storeDir = PackageApplyStore.FolderFor(PackageApplyStore.DefaultRoot, wbPath);
        var record = new ApplyRecord(wbPath, pkgPath, pkg);
        var filter = OleMessageFilter.Register();
        ExcelLauncher.LaunchResult? launch = null;
        try
        {
            ui?.Report("opening");
            launch = ExcelLauncher.CreateNew(visible: false);
            ui?.OwnExcel(launch.Pid, launch.StartTimeIso);
            return Apply(o, opts, confirm, launch, wbPath, pkgPath, pkg, codePage, storeDir, record, log, ui);
        }
        catch (COMException ex) when (ComRetry.IsPoison(ex.HResult) || ex.HResult == unchecked((int)0x800706BA))
        {
            ui?.Diagnostics?.Error("excel_disconnected", ex);
            record.Result = "failed";
            throw new CliError(ui == null
                    ? $"Excel stopped responding or exited during the apply (0x{ex.HResult:X8}) — the workbook was not saved"
                    : $"Excel connection was lost (0x{ex.HResult:X8}); save state: {ui.SaveState}",
                ui == null ? "The file on disk is unchanged; retry once, and if it repeats check that the workbook opens in Excel by hand"
                    : "Check the workbook and its backup before inspecting again",
                environment: true, code: "excel_disconnected");
        }
        catch (InvalidComObjectException ex)
        {
            ui?.Diagnostics?.Error("excel_invalid_com_object", ex);
            record.Result = "failed";
            throw new CliError(ui == null ? "Excel connection was lost during the apply — the workbook was not saved"
                    : $"Excel connection was lost; save state: {ui.SaveState}",
                ui == null ? "The file on disk is unchanged; retry once" : "Check the workbook and its backup before inspecting again",
                environment: true, code: "excel_disconnected");
        }
        catch (CliError e)
        {
            ui?.Diagnostics?.Error("apply_error", e);
            record.Error ??= e.Message;
            throw;
        }
        finally
        {
            if (ui != null && ui.SaveState != "not_saved") workingCopy?.PreserveOutput();
            filter.Revoke();
            if (launch != null) KillIfAlive(launch, log, ui?.Diagnostics);
            // --plan은 적용이 아니다 — 이력은 실제 적용 시도(거부 포함)만
            if (record.Result != null && !opts.Plan)
            {
                record.Error ??= o.Errors.FirstOrDefault();
                if (!PackageApplyStore.TryAppend(storeDir, record.ToJson()))
                    log.Write("warn", "package_apply_log_failed", new { folder = storeDir });
            }
        }
    }

    private static Outcome Apply(Outcome o, CliArgs opts, Func<Outcome, bool>? confirm,
        ExcelLauncher.LaunchResult launch, string wbPath, string pkgPath, PackageFormat.Package pkg,
        int codePage, string storeDir, ApplyRecord record, ToolLog log, PackageUiSession? ui)
    {
        dynamic app = launch.App;
        app.DisplayAlerts = false;
        app.EnableEvents = false; // Workbook_Open·시트 이벤트 미실행
        app.AskToUpdateLinks = false;
        int prevSec = Convert.ToInt32(app.AutomationSecurity);
        app.AutomationSecurity = 1; // msoAutomationSecurityLow — 열기 중 매크로 경고로 멈추지 않게
        dynamic wb;
        string openedPath = ui?.WorkingWorkbook ?? wbPath;
        try { wb = WorkbookOpener.Open(app, launch.Pid, openedPath, false, log).Workbook; }
        finally { app.AutomationSecurity = prevSec; }
        ui?.VerifyOpenedWorkbook(openedPath);
        if (Convert.ToBoolean(wb.ReadOnly))
            throw new CliError($"Excel opened the workbook read-only: {wbPath}",
                "Check write-reservation/password, workbook protection, permissions and other programs using the file; no changes were saved",
                environment: true, code: "workbook_read_only");

        // 2. 대상·버전 판정
        var stamp = new PackageVerdict.WorkbookStamp(
            ReadDocProp(wb, "vbaloop_Target"), ReadDocProp(wb, "vbaloop_BuildVersion"), ReadDocProp(wb, "vbaloop_PackageSha256"));
        record.FromVersion = stamp.Version;
        dynamic proj = SessionService.GetVbProject(wb);
        List<PackagePlan.WorkbookComponent> comps = ReadComponents(proj, Convert.ToString(wb.CodeName) ?? "");
        var plan = PackagePlan.Build(comps, pkg.Modules);
        IReadOnlyList<string>? diffs = null;
        var verdict = PackageVerdict.Decide(stamp, pkg.Meta, pkg.Sha256, opts.Adopt, opts.AllowDowngrade,
            () => (diffs = PackagePlan.CodeDifferences(plan, comps)).Count == 0);
        record.Decision = verdict.Decision.ToString();
        log.Write("info", "package_verdict", new { decision = verdict.Decision.ToString(), verdict.Reason, from = stamp.Version, to = pkg.Meta.Version });

        var steps = plan.Steps.Select(s => new { target = s.Target, action = StepLabel(s.Action), file = s.Module?.FileName }).ToList();
        var warnings = plan.Steps.Where(s => s.Action == PackagePlan.StepAction.IgnoreDocument)
            .Select(s => $"{s.Target}: workbook-only sheet module left as is — if its code no longer fits, the compile check fails")
            .ToList();
        object Payload(string? backup, object? compile, bool saved) => new
        {
            workbook = wbPath, package = pkgPath, target = pkg.Meta.Target,
            version = pkg.Meta.Version, previousVersion = stamp.Version, sha256 = pkg.Sha256,
            decision = DecisionLabel(verdict.Decision), reason = verdict.Reason, adopting = verdict.Adopting,
            changedModules = diffs, plan = steps, planErrors = plan.Errors,
            backup, backupFolder = storeDir, applyLog = Path.Combine(storeDir, PackageApplyStore.ApplyLogFileName),
            compile, saved, warnings,
            outputWorkbook = ui?.OutputWorkbook, copyMode = ui != null,
        };

        o.HumanLines.Add($"package: {pkg.Meta.Target} {pkg.Meta.Version} -> {wbPath}");
        o.HumanLines.Add($"  workbook: target {stamp.Target ?? "(none)"}, version {stamp.Version ?? "(none)"}");
        o.HumanLines.Add($"  decision: {verdict.Reason}");
        if (verdict.Adopting)
            o.HumanLines.Add($"  adopting: this workbook will be recorded as target {pkg.Meta.Target}");

        if (diffs is { Count: > 0 })
            o.HumanLines.Add($"  differs from the package: {string.Join(", ", diffs)}");
        if (verdict.Decision == PackageVerdict.Decision.Refuse)
        {
            o.ErrorCode = verdict.Code;
            record.Result = "refused";
            o.Payload = Payload(null, null, false);
            o.Fail($"package refused: {verdict.Reason}", verdict.Hint);
            return o;
        }
        if (verdict.Decision == PackageVerdict.Decision.AlreadyApplied)
        {
            record.Result = "alreadyApplied";
            o.Payload = Payload(null, null, false);
            o.HumanLines.Add("  nothing to do (no backup, not saved)");
            return o;
        }
        if (plan.Errors.Count > 0)
        {
            o.ErrorCode = "incompatible_workbook";
            record.Result = "refused";
            o.Payload = Payload(null, null, false);
            o.Fail($"package does not fit this workbook: {string.Join(" | ", plan.Errors)}",
                "Check that this is the right workbook; sheets in the package must already exist in it");
            return o;
        }
        foreach (var s in plan.Steps.Where(s => s.Action != PackagePlan.StepAction.Replace && s.Action != PackagePlan.StepAction.Inject))
            o.HumanLines.Add($"  {StepLabel(s.Action)}: {s.Target}");
        foreach (var w in warnings) o.HumanLines.Add($"  warning: {w}");

        // 3. 계획만
        if (opts.Plan)
        {
            o.Payload = Payload(null, null, false);
            o.HumanLines.Add($"  plan only — {plan.Steps.Count(s => s.Action is PackagePlan.StepAction.Replace or PackagePlan.StepAction.Inject)} modules would be replaced; nothing changed");
            o.HumanLines.Add($"  backups go to {storeDir}");
            return o;
        }
        if (confirm != null && !confirm(o))
        {
            o.Payload = Payload(null, null, false);
            o.HumanLines.Add("  cancelled — nothing changed");
            return o;
        }

        // 4. 백업 — 첫 변경 직전. 실패하면 워크북을 건드리기 전에 멈춘다. 아직 아무것도 바꾸지 않았으므로
        // 디스크 파일이 곧 적용 전 상태다 — SaveCopyAs가 아니라 바이트 그대로 복사한다(Excel 쓰기 경로 무관)
        ui?.Report("backing_up");
        var slot = BackupStore.Allocate(storeDir, BackupStore.KindPackage, wbPath, DateTime.UtcNow);
        try
        {
            CopyShared(openedPath, slot.CopyPath);
            BackupStore.Complete(slot, new BackupStore.Metadata
            {
                Command = "push --package",
                WorkbookPath = wbPath,
                ExeVersion = VersionCommand.GetBuildInfo(),
                PackageSha256 = pkg.Sha256,
                Target = pkg.Meta.Target,
                FromVersion = stamp.Version,
                ToVersion = pkg.Meta.Version,
            });
        }
        catch (Exception ex) when (!IsFatalCom(ex))
        {
            record.Result = "failed";
            BackupStore.Discard(slot);
            throw new CliError($"backup failed — aborted before touching the workbook: {ex.Message}",
                $"Check free space and write permission for {storeDir}", code: "backup_failed");
        }
        record.Backup = slot.Id;
        record.Result = "failed"; // 아래에서 성공으로 바뀌지 않으면 실패 기록
        o.HumanLines.Add($"  backup: {slot.CopyPath}");

        // 5. 주입 → 무결성
        var before = comps.Select(c => c.Name).ToList();
        var added = new List<string>();
        var removed = new List<string>();
        try
        {
            ui?.Report("applying");
            Inject(proj, plan, codePage, added, removed, log);
            CheckIntegrity(proj, plan, before, added, removed);
        }
        catch (CliError e)
        {
            o.ErrorCode = "apply_failed";
            o.Payload = Payload(slot.Id, null, false);
            o.Fail($"package apply aborted (not saved): {e.Message} — the file on disk is unchanged", e.Hint ?? "Fix the cause, then retry");
            return o;
        }

        // 6. 컴파일 게이트(항상)
        ui?.Report("compiling");
        var compile = VbeCompile.Run(app, proj, launch.Pid, log, "push");
        object compilePayload = new { result = compile.Label, detail = compile.Detail, location = compile.Location?.ToPayload() };
        if (!compile.Publishable || (ui != null && !PackageUiSession.CanSave(compile.Kind)))
        {
            o.ErrorCode = "compile_failed";
            record.Result = "compileError";
            o.Payload = Payload(slot.Id, compilePayload, false);
            o.Fail(
                (compile.Kind == DistCompileVerdict.Kind.CompileError
                    ? $"compile error — not saved: {compile.Detail}" : $"compile check unavailable — not saved: {compile.Detail}")
                + (compile.Location is { } loc ? $" — at {loc.Describe()}" : "")
                + " (the file on disk is unchanged)",
                compile.Kind == DistCompileVerdict.Kind.CompileError
                    ? "The package does not compile in this workbook (a workbook-only sheet module may reference removed code); fix it on the development PC"
                    : "Check VBA project access (AccessVBOM) on this PC with vbaloop doctor, then retry");
            return o;
        }
        if (compile.Kind == DistCompileVerdict.Kind.Unverified)
            warnings.Add($"compile check incomplete: {compile.Detail}");

        // 7. 문서 속성 → 8. 저장(원래 형식 유지)
        WriteDocProp(wb, "vbaloop_Target", pkg.Meta.Target);
        WriteDocProp(wb, "vbaloop_BuildVersion", pkg.Meta.Version);
        WriteDocProp(wb, "vbaloop_BuildRev", pkg.Meta.Rev ?? "");
        WriteDocProp(wb, "vbaloop_BuiltAt", pkg.Meta.BuiltAt ?? "");
        WriteDocProp(wb, "vbaloop_PackageSha256", pkg.Sha256);
        WriteDocProp(wb, "vbaloop_AppliedAt", DateTime.UtcNow.ToString("o"));
        string savePath = ui?.OutputWorkbook ?? wbPath;
        if (ui != null && !PathDrift.SamePath(Convert.ToString(wb.FullName) ?? "", savePath))
            throw new CliError("Excel is no longer editing the allocated output copy; save aborted", code: "save_uncertain");
        FileUseDiagnostics.Capture(ui?.Diagnostics, openedPath, "before_save", launch.Pid);
        ui?.Report("saving");
        ui?.Diagnostics?.FileState("workbook_before_save", wbPath);
        ui?.Diagnostics?.Write("save_enter");
        try
        {
            ui?.Diagnostics?.Write("save_in_place");
            ComRetry.Invoke(() => { wb.Save(); });
            if (ui != null && !Convert.ToBoolean(wb.Saved))
                throw new CliError("Excel has not confirmed the copy was saved", code: "save_uncertain");
        }
        catch (COMException ex)
        {
            ui?.Diagnostics?.Error("save_error", ex, includeMessage: true);
            FileUseDiagnostics.Capture(ui?.Diagnostics, savePath, "save_error_before_cleanup", launch.Pid);
            throw;
        }
        ui?.Diagnostics?.Write("save_returned");
        string actual = Convert.ToString(wb.FullName) ?? "";
        ui?.Diagnostics?.Write("save_path_checked", new { samePath = PathDrift.SamePath(actual, savePath), extension = Path.GetExtension(actual) });
        if (!PathDrift.SamePath(actual, savePath))
        {
            o.ErrorCode = "save_uncertain";
            o.Payload = Payload(slot.Id, compilePayload, false);
            o.Fail($"save drift: Excel saved to {actual} instead of {savePath}; check the output before using it",
                "Check the file lock and permissions, then retry", environment: true);
            return o;
        }
        ui?.Diagnostics?.Write("close_enter");
        ComRetry.Invoke(() => { wb.Close(false); });
        ui?.Diagnostics?.Write("close_returned");
        if (ui != null)
        {
            string outputHash = GuiWorkbookCopy.VerifyOutput(savePath);
            ui.Diagnostics?.Write("output_verified", new { sha256 = outputHash });
            ui.Report("saved");
        }
        try { ui?.Diagnostics?.Write("quit_enter"); app.Quit(); ui?.Diagnostics?.Write("quit_returned"); }
        catch (Exception ex) { ui?.Diagnostics?.Error("quit_error", ex); /* finally에서 kill */ }

        record.Result = verdict.Decision == PackageVerdict.Decision.Reapply ? "reapplied" : "applied";
        record.ToVersion = pkg.Meta.Version;
        log.Write("info", "package_applied", new { workbook = wbPath, version = pkg.Meta.Version, backup = slot.Id });
        o.Payload = Payload(slot.Id, compilePayload, true);
        o.HumanLines.Add($"push --package done (saved): {DecisionLabel(verdict.Decision)} {pkg.Meta.Version}, compile {compile.Label}");
        o.HumanLines.Add($"  restore by copying {slot.CopyPath} back over the workbook");
        return o;
    }

    // ---- 주입 ----

    private static void Inject(dynamic proj, PackagePlan.Result plan, int codePage, List<string> added, List<string> removed, ToolLog log)
    {
        List<VbaProjectOps.ComponentInfo> comps = VbaProjectOps.ListComponents(proj);
        var byName = comps.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase);
        foreach (var s in plan.Steps)
        {
            switch (s.Action)
            {
                case PackagePlan.StepAction.Replace or PackagePlan.StepAction.Inject:
                    VbaProjectOps.ReplaceCode(byName[s.Target].Comp.CodeModule, s.Body);
                    break;
                case PackagePlan.StepAction.Add:
                    int type = s.Module!.Kind == PackageModuleKind.Std ? VbaProjectOps.TypeStdModule : VbaProjectOps.TypeClassModule;
                    dynamic comp = ComRetry.Invoke(() => proj.VBComponents.Add(type));
                    ComRetry.Invoke(() => { comp.Name = s.Target; });
                    string actualName = Convert.ToString(comp.Name) ?? "";
                    if (!actualName.Equals(s.Target, StringComparison.OrdinalIgnoreCase))
                        throw new CliError($"{s.Target}: failed to name the new module (actual: {actualName})");
                    added.Add(s.Target);
                    VbaProjectOps.ReplaceCode(comp.CodeModule, s.Body);
                    break;
                case PackagePlan.StepAction.ReimportForm:
                    VbaProjectOps.VerifiedRemove(proj, byName[s.Target].Comp, s.Target);
                    removed.Add(s.Target);
                    FormImport.Import(proj, s.Target, s.Header, s.Body, s.Module!.Frx!, codePage);
                    added.Add(s.Target);
                    break;
                case PackagePlan.StepAction.ImportForm:
                    FormImport.Import(proj, s.Target, s.Header, s.Body, s.Module!.Frx!, codePage);
                    added.Add(s.Target);
                    break;
                default:
                    continue; // Remove는 아래, IgnoreDocument는 그대로 둔다
            }
            log.Write("info", "package_module", new { module = s.Target, action = StepLabel(s.Action) });
        }
        foreach (var s in plan.Steps.Where(s => s.Action == PackagePlan.StepAction.Remove))
        {
            VbaProjectOps.VerifiedRemove(proj, byName[s.Target].Comp, s.Target);
            removed.Add(s.Target);
            log.Write("info", "package_module", new { module = s.Target, action = "remove" });
        }
    }

    /// <summary>push 4단계와 같은 무결성 검사: 예상 밖 이름(&lt;이름&gt;1 등)·누락, 본문이 있는 모듈의 빈 주입.</summary>
    private static void CheckIntegrity(dynamic proj, PackagePlan.Result plan, List<string> before, List<string> added, List<string> removed)
    {
        HashSet<string> after = VbaProjectOps.ComponentNames(proj);
        var diff = NameSetDiff.Check(before, added, removed, after);
        if (diff.Unexpected.Count > 0 || diff.Missing.Count > 0)
            throw new CliError("integrity check failed: "
                + (diff.Unexpected.Count > 0 ? $"unexpected names {string.Join(", ", diff.Unexpected)} " : "")
                + (diff.Missing.Count > 0 ? $"missing names {string.Join(", ", diff.Missing)}" : ""));

        List<VbaProjectOps.ComponentInfo> comps = VbaProjectOps.ListComponents(proj);
        var byName = comps.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase);
        foreach (var s in plan.Steps.Where(s => s.Module != null && CanonicalText.TrimTrailingEmpty(s.Body).Length > 0))
            if (Convert.ToInt32(byName[s.Target].Comp.CodeModule.CountOfLines) == 0)
                throw new CliError($"integrity check failed: {s.Target} body is empty after injection");
    }

    private static List<PackagePlan.WorkbookComponent> ReadComponents(dynamic proj, string workbookCodeName)
    {
        var list = new List<PackagePlan.WorkbookComponent>();
        List<VbaProjectOps.ComponentInfo> comps = VbaProjectOps.ListComponents(proj);
        foreach (var c in comps)
        {
            VbaComponentKind? kind = c.Type switch
            {
                VbaProjectOps.TypeStdModule => VbaComponentKind.StdModule,
                VbaProjectOps.TypeClassModule => VbaComponentKind.ClassModule,
                VbaProjectOps.TypeMsForm => VbaComponentKind.MsForm,
                VbaProjectOps.TypeDocument => VbaComponentKind.Document,
                _ => null,
            };
            string body = kind == null ? "" : VbaProjectOps.ReadModuleBody(c.Comp.CodeModule);
            bool isWorkbook = kind == VbaComponentKind.Document && c.Name.Equals(workbookCodeName, StringComparison.OrdinalIgnoreCase);
            list.Add(new PackagePlan.WorkbookComponent(c.Name, kind, isWorkbook, body));
        }
        return list;
    }

    // ---- 사전 검사 ----

    private static void GuardEncoding(PackageFormat.Package pkg, int codePage)
    {
        var violations = new List<string>();
        foreach (var m in pkg.Modules)
            foreach (var v in Cp949Guard.Check(m.Lines, codePage).Take(5))
                violations.Add($"{m.FileName}:{v.Line}:{v.Column} '{v.Char}' ({v.CodePoint})");
        if (violations.Count > 0)
            throw new CliError(
                $"package has characters this PC cannot store in VBA (CP{codePage}): {string.Join(" | ", violations.Take(5))}",
                "The system code page differs from the development PC; remove those characters from src and build the package again", code: "encoding_mismatch");
    }

    /// <summary>
    /// 쓰기 가능 여부를 Excel을 띄우기 전에 판정한다(설계 §13.3). 읽기 전용 속성은 닫아도 풀리지 않아 exit 2,
    /// 다른 프로세스의 잠금은 닫으면 풀리는 환경 상태라 exit 3.
    /// </summary>
    internal static void CheckWritable(string path)
    {
        if ((File.GetAttributes(path) & FileAttributes.ReadOnly) != 0)
            throw new CliError($"workbook file has the read-only attribute: {path}",
                "Clear the read-only attribute (file Properties), then run again", code: "read_only");
        try
        {
            // 읽기만 하는 관찰자는 허용한다. 기존 쓰기 핸들이나 공유 쓰기를 거부하는
            // 핸들은 여전히 충돌하므로 사용자 Excel과 동시에 수정하지 않는다.
            using var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        }
        catch (UnauthorizedAccessException)
        {
            throw new CliError($"no write permission for the workbook: {path}",
                "Copy the file to a folder you can write to, or ask for write access", code: "no_write_access");
        }
        catch (IOException ex) when (ex.HResult is unchecked((int)0x80070020) or unchecked((int)0x80070021))
        {
            string? owner = ReadLockOwner(path);
            throw new CliError($"workbook write access conflicts with a file lock (0x{ex.HResult:X8}){(owner != null ? $" (by {owner})" : "")}: {path}",
                "Close Excel/preview windows; if the file is closed, check shared-drive access or ask IT about document-security write restrictions",
                environment: true, code: "file_locked");
        }
        catch (IOException ex)
        {
            throw new CliError($"workbook file access failed (0x{ex.HResult:X8}): {path}: {ex.Message}",
                "Check the file path, storage device and network/share availability, then inspect again",
                environment: true, code: "file_io");
        }
    }

    private static string? ReadLockOwner(string path)
    {
        try
        {
            string lockFile = Path.Combine(Path.GetDirectoryName(path)!, "~$" + Path.GetFileName(path));
            if (!File.Exists(lockFile)) return null;
            byte[] bytes;
            using (var fs = new FileStream(lockFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                bytes = new byte[Math.Min(fs.Length, 512)];
                fs.ReadExactly(bytes);
            }
            return PackageApplyStore.ParseLockOwner(bytes, Encoding.GetEncoding(NativeMethods.GetACP()));
        }
        catch { return null; }
    }

    // ---- 문서 속성 ----

    private static string? ReadDocProp(dynamic wb, string name)
    {
        try
        {
            string? v = Convert.ToString(wb.CustomDocumentProperties.Item(name).Value);
            return string.IsNullOrWhiteSpace(v) ? null : v;
        }
        catch { return null; }
    }

    /// <summary>판정 근거라 실패하면 저장하지 않는다(다음 적용이 잘못 판정된다).</summary>
    private static void WriteDocProp(dynamic wb, string name, string value)
    {
        dynamic props = wb.CustomDocumentProperties;
        // 없는 속성의 Item은 E_INVALIDARG — 바인더가 COMException이 아니라 ArgumentException으로 던진다
        try { props.Item(name).Value = value; }
        catch (Exception ex) when (!IsFatalCom(ex))
        {
            props.Add(name, false, 4 /* msoPropertyTypeString */, value);
        }
        string? back = Convert.ToString(props.Item(name).Value);
        if (!string.Equals(back ?? "", value, StringComparison.Ordinal))
            throw new CliError($"document property {name} did not take the value (read back '{back}') — not saved",
                "The workbook may be protected; check it opens and saves in Excel by hand");
    }

    // ---- 표시 ----

    private static string StepLabel(PackagePlan.StepAction a) => a switch
    {
        PackagePlan.StepAction.Add => "add",
        PackagePlan.StepAction.Replace => "replace",
        PackagePlan.StepAction.Inject => "replace document",
        PackagePlan.StepAction.ImportForm => "add form",
        PackagePlan.StepAction.ReimportForm => "reimport form",
        PackagePlan.StepAction.Remove => "remove",
        _ => "leave sheet",
    };

    private static string DecisionLabel(PackageVerdict.Decision d) => d switch
    {
        PackageVerdict.Decision.Apply => "applied",
        PackageVerdict.Decision.AlreadyApplied => "alreadyApplied",
        PackageVerdict.Decision.Reapply => "reapplied",
        _ => "refused",
    };

    private static bool IsFatalCom(Exception ex)
        => ex is InvalidComObjectException || (ex is COMException ce && ComRetry.IsPoison(ce.HResult));

    /// <summary>Excel이 열고 있는 파일 복사 — Excel은 쓰기 공유를 거부한 채 열고 있어 읽기 쪽이 ReadWrite 공유를 허용해야 열린다.</summary>
    private static void CopyShared(string source, string dest)
    {
        using (var src = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (var dst = new FileStream(dest, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            src.CopyTo(dst);
        File.SetLastWriteTimeUtc(dest, File.GetLastWriteTimeUtc(source));
    }

    private static void KillIfAlive(ExcelLauncher.LaunchResult launch, ToolLog log, UpdaterDiagnostics? diagnostics = null)
    {
        // 도구 소유 인스턴스: Quit 후에도 RCW가 남아 종료되지 않으므로 신원이 맞으면 kill(dist 임시 Excel과 같은 처분)
        var identity = ProcessIdentity.Identify(launch.Pid, launch.StartTimeIso);
        diagnostics?.Write("cleanup_identity", new { launch.Pid, identity = identity.ToString() });
        if (identity != IdentityState.Match) return;
        try {
            diagnostics?.Write("cleanup_kill_enter", new { launch.Pid });
            Process.GetProcessById(launch.Pid).Kill(entireProcessTree: true);
            diagnostics?.Write("cleanup_kill_returned");
        }
        catch (Exception ex) { diagnostics?.Error("cleanup_kill_error", ex); log.Write("warn", "package_excel_kill_failed", new { pid = launch.Pid, error = ex.Message }); }
    }

    /// <summary>apply.jsonl 한 줄(설계 §13.5): 시각, 원래 경로, 버전, 결과, 백업 id.</summary>
    private sealed class ApplyRecord(string workbook, string package, PackageFormat.Package pkg)
    {
        public string? Decision { get; set; }
        public string? Result { get; set; }
        public string? FromVersion { get; set; }
        public string? ToVersion { get; set; }
        public string? Backup { get; set; }
        public string? Error { get; set; }

        public object ToJson() => new
        {
            ts = DateTime.UtcNow.ToString("o"),
            workbook, package, target = pkg.Meta.Target, packageVersion = pkg.Meta.Version, packageSha256 = pkg.Sha256,
            fromVersion = FromVersion, decision = Decision, result = Result, backup = Backup,
            error = Error,
        };
    }
}

/// <summary>같은 대상 경로의 패키지 적용 직렬화 — 프로젝트 mutex 대신 경로 기반 커널 뮤텍스(설계 §13.3).</summary>
internal sealed class PathMutex : IDisposable
{
    private readonly Mutex _mutex;
    private PathMutex(Mutex m) => _mutex = m;

    public static PathMutex Acquire(string fullPath)
    {
        var mutex = new Mutex(false, @"Local\vbaloop-pkg-" + PackageApplyStore.PathHash(fullPath)[..16]);
        bool held;
        try { held = mutex.WaitOne(0); }
        catch (AbandonedMutexException) { held = true; }
        if (!held)
        {
            mutex.Dispose();
            throw new CliError($"another vbaloop is already applying a package to this workbook: {fullPath}",
                "Wait for it to finish, then run again", environment: true);
        }
        return new PathMutex(mutex);
    }

    public void Dispose()
    {
        try { _mutex.ReleaseMutex(); } catch (ApplicationException) { }
        _mutex.Dispose();
    }
}
