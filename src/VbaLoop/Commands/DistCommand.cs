using System.Diagnostics;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop dist (설계 §9): 클린 사본 배포 아티팩트 빌더. base(세션과 무관한 on-disk 워크북)를
/// --out으로 복사 → 자체 임시 Excel 인스턴스에서 (필터 → 스탬프 → 무해화 → VBE 컴파일 검증 →
/// 선택적 부트스트랩)을 적용 → 저장 → 원자 게시. 컴파일 오류면 게시하지 않는다. 세션 워크북은 절대 건드리지 않으며, 부트스트랩 행(모달/MsgBox)도 임시 인스턴스
/// kill로 격리되어 세션에 무해하다. 순수 파이프라인 판정은 Core.Dist*가 담당(§9 이원화 금지).
/// </summary>
public static class DistCommand
{
    private const int BootstrapTimeoutSec = 120;

    /// <summary>git repo가 아니거나 git 부재로 프로젝트 리비전을 못 구할 때의 sentinel.
    /// 도구(vbaloop) 자신의 리비전을 배포본에 스탬프하지 않기 위함(외부 감사 필드 오염 방지, §9.4).</summary>
    private const string NoRevSentinel = "nogit";

    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "dist" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "dist");
        o.LogRecords = log.Records;

        string outPath = opts.Out != null
            ? Path.GetFullPath(opts.Out)
            : DistFormat.DefaultOut(cfg.RootDir, cfg.WorkbookPath);
        DistOutputKind format = DistFormat.FromOutPath(outPath); // 확장자로 형식 판정(미지원 조기 거부)
        string formatLabel = format == DistOutputKind.AddIn ? "xlam" : "xlsm";
        bool strip = opts.StripComments || cfg.DistStripComments;
        string? bootstrap = opts.Bootstrap ?? cfg.DistBootstrap;
        if (bootstrap != null && !MacroName.TryParse(bootstrap, out _, out _, out var bootErr))
            throw new CliError($"--bootstrap format error: {bootErr}", "Format: Module.Proc");
        if (PathDrift.SamePath(outPath, cfg.WorkbookPath))
            throw new CliError("--out is the same as the source workbook — the artifact must be a separate file");

        using var cmdLock = CommandMutex.Acquire(cfg);

        // 1. 사전 검증(Excel 불필요) — 사본을 만들기 전에 전부 끝낸다
        int codePage = NativeMethods.GetACP();
        var modules = LoadSrcModules(cfg, codePage);
        var classified = DistFilter.Classify(modules, cfg.DistExclude);
        var dangling = DistCompileScan.Scan(classified.Included, classified.Excluded);
        if (dangling.Count > 0)
        {
            log.Write("error", "dist_dangling_refs", new { count = dangling.Count });
            o.Payload = new { danglingRefs = dangling.Take(20).Select(d => new { d.InModule, d.Line, d.Symbol, d.Kind }) };
            o.Fail(
                $"Compile safety failure: remaining modules reference symbols of excluded modules — " +
                string.Join(", ", dangling.Take(5).Select(d => $"{d.InModule}:{d.Line} {d.Symbol}({d.Kind})")) +
                (dangling.Count > 5 ? $" and {dangling.Count - 5} more" : ""),
                "Narrow the exclusions or remove those references (Test modules are safe to drop wholesale only if they reference each other alone)");
            return o;
        }

        // 2. 동기화 가드: on-disk 워크북 == 마지막 push == src (사본의 코드 신뢰 근거)
        var sync = SyncCache.Load(SyncCache.PathFor(cfg));
        string fingerprint = Fingerprint.Compute(cfg.SrcRoot);
        if (sync.LastPushFingerprint != fingerprint)
        {
            o.Fail("src and workbook are out of sync — dist uses the copy from the last push as its base",
                "Apply src to the workbook with vbaloop push, then retry");
            return o;
        }
        if (!File.Exists(cfg.WorkbookPath))
            throw new CliError($"base workbook file not found: {cfg.WorkbookPath}", environment: true);

        // 3. 버전 스탬프 재료
        string version = cfg.DistVersion ?? "0.0.0";
        string rev = ProjectRev(cfg.RootDir);
        string builtAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        // 4. base 사본 → 스테이징 파일(원자 이동용). outPath는 전 파이프라인 성공 후에만 교체한다
        //    → 실패해도 기존 배포본(outPath)이 파괴되지 않는다.
        Directory.CreateDirectory(Path.GetDirectoryName(outPath)!);
        string stagePath = outPath + ".tmp-" + Environment.ProcessId;
        // 저장 대상: xlsm은 stage에 in-place Save. 애드인은 SaveAs 대상 확장자가 반드시 .xlam이어야
        // 한다 — 형식과 안 맞는 확장자(.tmp-*)로 SaveAs(55)하면 Excel이 스크래치 파일에 쓰고 최종
        // 리네임에 실패해 stage엔 미변형 사본만 남는다(런타임 실측 2026-07-10). 별도 .xlam 대상에
        // 저장한 뒤 원자 이동한다.
        string savePath = format == DistOutputKind.AddIn ? stagePath + ".xlam" : stagePath;
        TryDelete(stagePath); // 동일 PID의 잔재(하드킬 등) 제거 후 충돌 없이 복사
        TryDelete(savePath);
        File.Copy(cfg.WorkbookPath, stagePath, overwrite: false);

        // 5. 임시 인스턴스에서 사본(stage) 변형
        var warnings = new List<string>();
        if (rev == NoRevSentinel)
            warnings.Add($"git rev unavailable (not a repo or git missing) — BuildRev='{NoRevSentinel}'");
        bool moved = false;
        var filter = OleMessageFilter.Register();
        ExcelLauncher.LaunchResult? launch = null;
        try
        {
            var l = ExcelLauncher.CreateNew(visible: false);
            launch = l;
            dynamic app = l.App;
            try
            {
                app.DisplayAlerts = false;
                app.EnableEvents = false;
                app.AskToUpdateLinks = false;
                int prevSec = Convert.ToInt32(app.AutomationSecurity);
                app.AutomationSecurity = 1; // msoAutomationSecurityLow — Workbook_Open 억제
                dynamic wb;
                try { wb = app.Workbooks.Open(stagePath, 0, false); }
                finally { app.AutomationSecurity = prevSec; }
                if (Convert.ToBoolean(wb.ReadOnly))
                    throw new CliError($"copy opened read-only: {stagePath}",
                        "Check locks/permissions on the output path", environment: true);

                dynamic proj = SessionService.GetVbProject(wb);
                var mut = MutateProject(proj, classified, strip, version, rev, builtAt, warnings, log);

                // 5.5 게시 전 VBE 컴파일 검증(§9.8) — 변형이 끝난 사본 그대로를 컴파일한다.
                //     부트스트랩보다 먼저: 컴파일 안 되는 사본에서 매크로를 돌리면 모달 행부터 만난다.
                DistCompileVerdict.Verdict? compile = null;
                if (opts.SkipCompile)
                {
                    warnings.Add("--skip-compile: VBE compile check skipped before publishing");
                    log.Write("warn", "dist_compile_skipped", null);
                }
                else
                {
                    compile = VbeCompile.Run(app, proj, l.Pid, log, "dist");
                    if (!compile.Publishable)
                    {
                        // 게시 금지: outPath는 손대지 않고, stage/save 파일은 finally가 정리한다
                        try { app.Quit(); } catch { /* finally에서 kill 보증 */ }
                        o.Payload = new
                        {
                            outPath,
                            published = false,
                            compile = new { result = compile.Label, detail = compile.Detail },
                        };
                        o.Fail(
                            compile.Kind == DistCompileVerdict.Kind.CompileError
                                ? $"Compile error — not published: {compile.Detail}"
                                : $"Compile check unavailable — not published: {compile.Detail}",
                            compile.Kind == DistCompileVerdict.Kind.CompileError
                                ? "Fix the error in src, push, then retry (if caused by references to excluded modules, check [dist] exclude/@dist)"
                                : "Check VBE access (AccessVBOM) and modal state with vbaloop doctor, then retry. Publish without the check (--skip-compile) only after the user confirms");
                        return o;
                    }
                    if (compile.Kind == DistCompileVerdict.Kind.Unverified)
                        warnings.Add($"Compile check incomplete: {compile.Detail}");
                }

                object? bootResult = null;
                if (bootstrap != null)
                    bootResult = RunBootstrap(app, wb, bootstrap, l.Pid, log);

                SetDocProp(wb, "vbaloop_BuildVersion", version, warnings);
                SetDocProp(wb, "vbaloop_BuildRev", rev, warnings);
                SetDocProp(wb, "vbaloop_BuiltAt", builtAt, warnings);

                if (format == DistOutputKind.AddIn)
                {
                    // 애드인: 시트를 숨기고(IsAddin) .xlam 대상에 xlOpenXMLAddIn으로 SaveAs — base
                    // .xlsm를 .xlam로 변환. 주입 코드·버전 상수·문서속성은 포맷 무관하게 그대로
                    // 이월된다(§9.7). savePath는 확장자가 .xlam이라 최종 리네임까지 정상 완료된다.
                    ComRetry.Invoke(() => { wb.IsAddin = true; });
                    ComRetry.Invoke(() => { wb.SaveAs(savePath, DistFormat.XlOpenXMLAddIn); });
                }
                else
                {
                    ComRetry.Invoke(() => { wb.Save(); }); // savePath == stagePath, xlsm 포맷 유지
                }
                ComRetry.Invoke(() => { wb.Close(false); });
                try { app.Quit(); } catch { /* finally에서 kill 보증 */ }

                // 원자 이동: wb.Close로 저장 파일 핸들이 풀린 뒤에만 최종 경로로 교체.
                // 이 지점 이전의 어떤 실패도 outPath(기존 배포본)를 건드리지 않는다.
                MoveIntoPlace(savePath, outPath);
                moved = true;

                log.Write("info", "dist_done", new
                {
                    outPath, format = formatLabel, version, rev,
                    included = classified.Included.Count,
                    excluded = classified.ExcludedNames.Count,
                    removed = mut.Removed, stripped = mut.Stripped, logShimFlipped = mut.LogShimFlipped,
                    bootstrap = bootstrap != null,
                    compile = compile?.Label ?? "skipped",
                });

                o.Payload = new
                {
                    outPath, format = formatLabel, version, rev, builtAt,
                    included = classified.Included.Select(m => m.Name),
                    excluded = classified.ExcludedNames,
                    removed = mut.Removed,
                    strippedModules = mut.Stripped,
                    logShimFlipped = mut.LogShimFlipped,
                    stripComments = strip,
                    bootstrap = bootResult,
                    compile = compile == null ? new { result = "skipped", detail = (string?)null }
                        : new { result = compile.Label, detail = compile.Detail },
                    published = true,
                    warnings,
                };
                o.HumanLines.Add($"dist done: {outPath} [{formatLabel}]");
                o.HumanLines.Add($"  version {version} (rev {rev}, built {builtAt})");
                o.HumanLines.Add($"  included {classified.Included.Count} / excluded {classified.ExcludedNames.Count}" +
                    (mut.Removed > 0 ? $" (removed {mut.Removed})" : "") +
                    (strip ? $", comments stripped in {mut.Stripped} modules" : "") +
                    (mut.LogShimFlipped ? ", LogShim sanitized" : ""));
                if (bootstrap != null) o.HumanLines.Add($"  bootstrap run: {bootstrap}");
                o.HumanLines.Add(compile == null ? "  compile check: skipped (--skip-compile)"
                    : $"  compile check: {compile.Label}");
                foreach (var w in warnings) o.HumanLines.Add($"  warning: {w}");
                return o;
            }
            catch
            {
                try { app.Quit(); } catch { /* finally kill */ }
                throw;
            }
        }
        catch (Exception ex)
        {
            // stage 정리는 finally에서(임시 Excel kill 후 핸들 해제까지 재시도) — 여기서
            // 지우면 아직 열려있는 임시 인스턴스가 파일을 물고 있어 조용히 실패한다.
            if (ex is CliError) throw;
            o.Fail($"dist failed: {ex.GetType().Name}: {ex.Message}");
            return o;
        }
        finally
        {
            filter.Revoke();
            if (launch != null) KillIfAlive(launch.Pid, launch.StartTimeIso, log);
            // 임시 Excel이 파일을 물고 있다가 kill 후 늦게 놓으므로 재시도 삭제. 이동 성공(moved)
            // 시엔 savePath가 이미 없어 no-op. outPath는 절대 손대지 않는다.
            if (!moved) TryDeleteWithRetry(savePath);
            // 애드인은 열었던 원본 사본(stagePath)이 savePath와 별개로 남으므로 항상 정리.
            if (savePath != stagePath) TryDeleteWithRetry(stagePath);
        }
    }

    private readonly record struct MutationResult(int Removed, int Stripped, bool LogShimFlipped);

    /// <summary>
    /// 사본 VBProject에 dist 변형 적용: 제외 모듈 제거 → (선택)주석 제거 → LogShim 무해화 →
    /// modBuildInfo 합성 추가. 순수 판정은 Core.Dist*/VbaCommentStripper/LogShimRelease가,
    /// COM 주입은 VbaProjectOps 프리미티브가 담당.
    /// </summary>
    private static MutationResult MutateProject(
        dynamic proj, DistFilter.Result classified, bool strip,
        string version, string rev, string builtAt, List<string> warnings, ToolLog log)
    {
        var excludedSet = new HashSet<string>(classified.ExcludedNames, StringComparer.OrdinalIgnoreCase);
        int removed = 0, stripped = 0;
        bool logShimFlipped = false;

        // 제외 모듈 제거 (문서/시트 모듈은 제거 불가 → 경고만)
        List<VbaProjectOps.ComponentInfo> toRemove = VbaProjectOps.ListComponents(proj);
        foreach (var comp in toRemove)
        {
            if (!excludedSet.Contains(comp.Name)) continue;
            if (comp.Type == VbaProjectOps.TypeDocument)
            {
                warnings.Add($"{comp.Name}: document/sheet modules cannot be removed — exclusion ignored");
                continue;
            }
            VbaProjectOps.VerifiedRemove(proj, comp.Comp, comp.Name);
            log.Write("info", "dist_module_removed", new { module = comp.Name });
            removed++;
        }

        // 잔존 모듈: 주석 제거 + LogShim 무해화
        bool sawLogShim = false;
        List<VbaProjectOps.ComponentInfo> remaining = VbaProjectOps.ListComponents(proj);
        foreach (var comp in remaining)
        {
            if (comp.Name.Equals(VersionStamp.ModuleName, StringComparison.OrdinalIgnoreCase))
                continue; // 아래에서 새로 합성/교체
            // comp.Comp는 dynamic이므로 본문을 string으로 고정해야 확장메서드(LINQ)가 정적 바인딩됨
            string raw = VbaProjectOps.ReadModuleBody(comp.Comp.CodeModule);
            List<string> body = CanonicalText.SplitLines(raw).ToList();
            bool changed = false;

            if (strip)
            {
                var s = VbaCommentStripper.Strip(body);
                if (!s.SequenceEqual(body, StringComparer.Ordinal)) { body = s.ToList(); stripped++; changed = true; }
            }
            if (comp.Name.Equals("LogShim", StringComparison.OrdinalIgnoreCase))
            {
                sawLogShim = true;
                body = LogShimRelease.FlipToRelease(body).ToList(); // 상수 부재 시 CliError
                logShimFlipped = true;
                changed = true;
            }
            if (changed) VbaProjectOps.ReplaceCode(comp.Comp.CodeModule, body);
        }
        if (!sawLogShim)
            warnings.Add("LogShim component missing from the copy — sanitize skipped (if logging is depended on, check re-vendoring)");

        // modBuildInfo 합성(§9.4): 있으면 교체, 없으면 추가
        var buildBody = VersionStamp.BuildInfoBody(version, rev, builtAt);
        List<VbaProjectOps.ComponentInfo> afterMutation = VbaProjectOps.ListComponents(proj);
        var existing = afterMutation
            .FirstOrDefault(c => c.Name.Equals(VersionStamp.ModuleName, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
            VbaProjectOps.ReplaceCode(existing.Comp.CodeModule, buildBody);
        else
            VbaProjectOps.AddStdModule(proj, VersionStamp.ModuleName, buildBody);

        return new MutationResult(removed, stripped, logShimFlipped);
    }

    /// <summary>
    /// 부트스트랩(§9.0): 사본에서 워크시트 초기설정 프로시저를 in-process로 실행하되, 임시
    /// 인스턴스 kill 워치독으로 모달/MsgBox 행을 세션과 무관하게 격리(MsgBox는 test와 동일 non-goal).
    /// </summary>
    private static object RunBootstrap(dynamic app, dynamic wb, string macro, int pid, ToolLog log)
    {
        string wbName = Convert.ToString(wb.Name) ?? "";
        string macroRef = $"'{wbName}'!{macro}";
        using var done = new ManualResetEventSlim(false);
        bool killed = false;
        var watchdog = new Thread(() =>
        {
            if (!done.Wait(BootstrapTimeoutSec * 1000))
            {
                killed = true;
                try { Process.GetProcessById(pid).Kill(entireProcessTree: true); } catch { /* 이미 죽음 */ }
            }
        }) { IsBackground = true };
        watchdog.Start();
        try
        {
            ComRetry.Invoke(() => { app.Run(macroRef); });
        }
        catch (Exception ex)
        {
            if (killed)
                throw new CliError(
                    $"Bootstrap timeout ({BootstrapTimeoutSec}s): {macro} — temporary Excel killed (suspect modal/MsgBox)",
                    "Check that the bootstrap procedure is non-interactive and compiles clean");
            throw new CliError($"Bootstrap run failed: {macro}: {ex.Message}",
                "Check that the procedure compiles and runs properly in the copy");
        }
        finally { done.Set(); }
        log.Write("info", "dist_bootstrap_done", new { macro });
        return new { macro, ran = true };
    }

    private static void SetDocProp(dynamic wb, string name, string value, List<string> warnings)
    {
        try
        {
            dynamic props = wb.CustomDocumentProperties;
            try { props.Add(name, false, 4 /* msoPropertyTypeString */, value); }
            catch { props.Item(name).Value = value; }
        }
        catch (Exception ex)
        {
            warnings.Add($"Failed to write document property {name} (non-fatal): {ex.Message}");
        }
    }

    /// <summary>src 모듈을 헤더/본문 분리해 로드(Excel 불필요). CP949 가드 재적용(방어).</summary>
    private static List<DistModule> LoadSrcModules(ProjectConfig cfg, int codePage)
    {
        if (!Directory.Exists(cfg.SrcRoot))
            throw new CliError($"src root not found: {cfg.SrcRoot}", "Create it initially with vbaloop pull");

        var files = Directory.EnumerateFiles(cfg.SrcRoot)
            .Where(f => Path.GetExtension(f).ToLowerInvariant() is ".bas" or ".cls" or ".frm")
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (files.Count == 0)
            throw new CliError($"no source files in src: {cfg.SrcRoot}");

        var violations = new List<string>();
        var modules = new List<DistModule>();
        foreach (var file in files)
        {
            var lines = CanonicalText.SplitLines(CanonicalText.ReadFileStripBom(file));
            foreach (var v in Cp949Guard.Check(lines, codePage).Take(5))
                violations.Add($"{Path.GetFileName(file)}:{v.Line}:{v.Column} '{v.Char}' ({v.CodePoint})");
            var split = VbaHeader.Split(lines, file, strict: false);
            modules.Add(new DistModule(
                Path.GetFileNameWithoutExtension(file),
                Path.GetExtension(file).ToLowerInvariant(),
                split.Header, split.Body));
        }
        if (violations.Count > 0)
            throw new CliError(
                $"Encoding guard failed — characters outside CP{codePage}: {string.Join(" | ", violations.Take(5))}",
                "Remove/replace those characters, then retry");
        return modules;
    }

    /// <summary>프로젝트 repo의 git 리비전(+dirty). git이 없거나 repo 아니면 <see cref="NoRevSentinel"/>
    /// 폴백 — 툴 자신의 리비전을 배포본에 스탬프하면 감사 필드를 조용히 오도하므로 금지(§9.4).</summary>
    private static string ProjectRev(string root)
    {
        try
        {
            string rev = GitCapture(root, "rev-parse --short HEAD");
            if (string.IsNullOrWhiteSpace(rev)) return NoRevSentinel;
            string status = GitCapture(root, "status --porcelain");
            return status.Trim().Length > 0 ? rev.Trim() + "-dirty" : rev.Trim();
        }
        catch { return NoRevSentinel; }
    }

    private static string GitCapture(string root, string args)
    {
        var psi = new ProcessStartInfo("git", $"-C \"{root}\" {args}")
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        using var p = Process.Start(psi) ?? throw new InvalidOperationException("failed to start git");
        string outp = p.StandardOutput.ReadToEnd();
        p.WaitForExit(5000);
        return p.ExitCode == 0 ? outp : "";
    }

    private static void KillIfAlive(int pid, string startTimeIso, ToolLog log)
    {
        if (!ProcessIdentity.IsAlive(pid, startTimeIso)) return;
        try { Process.GetProcessById(pid).Kill(entireProcessTree: true); }
        catch (Exception ex) { log.Write("warn", "dist_excel_kill_failed", new { pid, error = ex.Message }); }
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* 무시 */ }
    }

    /// <summary>stage 파일 삭제 — 임시 Excel이 kill 직후 파일 핸들을 늦게 놓는 경우를 대비해
    /// 짧게 재시도. 끝내 못 지워도 조용히 포기(outPath는 안전하므로 데이터 손실 아님, 잔재만).</summary>
    private static void TryDeleteWithRetry(string path)
    {
        for (int i = 0; i < 5; i++)
        {
            try { if (File.Exists(path)) File.Delete(path); return; }
            catch (IOException) { Thread.Sleep(200); }
            catch (UnauthorizedAccessException) { Thread.Sleep(200); }
        }
        TryDelete(path); // 마지막 시도(그래도 실패면 무시)
    }

    /// <summary>stage 파일을 최종 경로로 원자 이동. Excel 프로세스가 파일 핸들을 늦게 놓는
    /// 경우(close 직후 짧은 잔존 락)를 대비해 짧게 재시도. 끝내 실패하면 예외 전파 —
    /// 호출부가 stage를 정리하고 outPath(기존 배포본)는 보존한다.</summary>
    private static void MoveIntoPlace(string stagePath, string outPath)
    {
        IOException? last = null;
        for (int i = 0; i < 5; i++)
        {
            try { File.Move(stagePath, outPath, overwrite: true); return; }
            catch (IOException ex) { last = ex; Thread.Sleep(200); }
        }
        throw last ?? new IOException($"atomic move failed: {stagePath} -> {outPath}");
    }
}
