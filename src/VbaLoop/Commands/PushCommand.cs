using System.Text;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop push (설계 §4): 헤더 스트립 → push 전 백업(SaveCopyAs) → DeleteLines+AddFromString
/// in-place 주입 → 이름 집합 무결성 체크 → save. Remove+Import는 코드 교체 경로에서 완전히 제거.
/// 폼만 폼 예외 절차(검증형 Remove → Import → 중복 이름 검사).
/// </summary>
public static class PushCommand
{
    private sealed class SrcModule
    {
        public required string Name { get; init; }
        public required string Ext { get; init; }
        public required string FilePath { get; init; }
        public required IReadOnlyList<string> Header { get; init; }
        public required IReadOnlyList<string> Body { get; init; }
        public string? FrxPath { get; init; }
        public bool IsDocumentHeader { get; init; }
    }

    private sealed class ModuleResult
    {
        public required string Name { get; init; }
        public required string Action { get; set; } // added|replaced|skipped|failed|pruned
        public string? Reason { get; set; }
        public int BodyLineCount { get; init; }
    }

    /// <summary>
    /// 이번 push가 메모리 워크북에 가한 변경 기록 — 실패 시 PushRecovery가 이걸로 원복/poison을 정한다.
    /// 기록은 변경 시도 "직전"에 남긴다: 주입 도중 실패한 모듈도 반쯤 바뀐 상태일 수 있기 때문.
    /// </summary>
    private sealed class Journal
    {
        public List<PushRecovery.Touch> Touches { get; } = new();
        public Dictionary<string, string> Snapshots { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, dynamic> Added { get; } = new(StringComparer.OrdinalIgnoreCase);

        public void CodeReplacing(string name, string snapshot)
        {
            Snapshots.TryAdd(name, snapshot);
            Touches.Add(new PushRecovery.Touch(name, PushRecovery.TouchKind.CodeReplaced));
        }

        public void ComponentAdded(string name, dynamic comp)
        {
            Added[name] = comp;
            Touches.Add(new PushRecovery.Touch(name, PushRecovery.TouchKind.ComponentAdded));
        }

        public void FormReimporting(string name)
            => Touches.Add(new PushRecovery.Touch(name, PushRecovery.TouchKind.FormReimported));

        public void Pruning(string name)
            => Touches.Add(new PushRecovery.Touch(name, PushRecovery.TouchKind.ComponentPruned));
    }

    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "push" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "push");
        o.LogRecords = log.Records;

        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "push", log);

        var ses = SessionEnsure.ForCommand(cfg, log, o, "push");
        var (app, meta) = (ses.App, ses.Meta);
        var filter = OleMessageFilter.Register();
        try
        {
            return ComFailureGuard.Run(cfg, meta, log, () => Execute(o, opts, cfg, app, meta, log));
        }
        finally
        {
            filter.Revoke();
        }
    }

    private static Outcome Execute(Outcome o, CliArgs opts, ProjectConfig cfg, dynamic app, SessionMeta meta, ToolLog log)
    {
        // 0. 해시 캐시(changed-only): 지문 일치면 no-op 성공 (--prune/--force 시에는 우회)
        // 캐시는 세션 메타가 아니라 영속 sync.json — session stop 후에도 살아남는다.
        // --force는 캐시 불신 시 전량 재주입 수단(워크북과 src 지문이 어긋난 의심 상황).
        var sync = SyncCache.Load(SyncCache.PathFor(cfg));
        string fingerprint = Fingerprint.Compute(cfg.SrcRoot);
        if (sync.LastPushFingerprint == fingerprint)
        {
            if (!opts.Force && !opts.Prune)
            {
                o.HumanLines.Add("push: 변경 없음 (소스 지문 일치, no-op)");
                o.Payload = new { cached = true, fingerprint, modules = Array.Empty<object>() };
                return o;
            }
            if (opts.Force)
                log.Write("info", "push_cache_bypassed", new { fingerprint });
        }

        // 1. 파일 읽기 + 사전 검증 — 워크북을 건드리기 전에 전부 끝낸다
        int codePage = NativeMethods.GetACP(); // 상수 하드코딩 금지
        var srcModules = LoadSrc(cfg, codePage, log);

        dynamic wb = SessionService.FindWorkbook(app, meta.WorkbookFullName);

        // read-only 워크북에 주입하면 Save가 조용히 SaveAs 오탈출한다.
        // session start가 이미 fail-fast하지만, 구 바이너리로 만든 채택 세션·세션 중 잠금 발생을
        // 대비해 주입 시작 전에 한 번 더 차단한다.
        if (Convert.ToBoolean(wb.ReadOnly))
        {
            log.Write("error", "push_readonly_refused", new { workbook = meta.WorkbookFullName });
            throw new CliError(
                "워크북이 read-only 상태 — 저장 불가이므로 주입 전에 push 중단",
                "session stop 후 vbaloop doctor로 잠금 보유 프로세스 확인, 정리 후 재시도",
                environment: true);
        }

        dynamic proj = SessionService.GetVbProject(wb);

        List<VbaProjectOps.ComponentInfo> components = VbaProjectOps.ListComponents(proj);
        var byName = components.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase);
        var before = components.Select(c => c.Name).ToList();

        // 4(사전). 삭제/rename 판정 — 주입 전에 실패시켜 부분 변경을 피한다.
        // 문서/시트 모듈(TypeDocument)은 생성·삭제 불가 → 워크북 전용 존재가 정상 상태이므로
        // 프루닝/차단에서 제외한다. 표준/클래스/폼만 프루너블.
        var srcNames = new HashSet<string>(srcModules.Select(m => m.Name), StringComparer.OrdinalIgnoreCase);
        var plan = ExtraModulePlan.Classify(components.Select(c => (c.Name, c.Type)), srcNames);
        var extras = components
            .Where(c => plan.Prunable.Contains(c.Name, StringComparer.OrdinalIgnoreCase))
            .ToList();
        if (plan.Prunable.Count > 0 && !opts.Prune)
            throw new CliError(
                $"src에 없는 워크북 모듈 존재: {string.Join(", ", plan.Prunable)} (rename은 삭제+신규로 취급됨)",
                "--prune으로 명시 삭제");

        // 0.5. push 전 백업 — 첫 워크북 변형 직전. 캐시 히트(no-op)는 위에서 이미 반환했다.
        // 실패하면 워크북을 건드리기 전에 중단(fail closed): 백업 없는 주입은 rollback 근거가 없다.
        string? backupId = BackupBeforeMutation(cfg, app, wb, fingerprint, sync, log);
        if (backupId != null)
            o.HumanLines.Add($"백업: _dbg/backups/{backupId} (되돌리기: session stop 후 vbaloop rollback --latest)");

        var results = new List<ModuleResult>();

        // 무시된 문서/시트 모듈은 차단하지 않고 결과에 기록만 한다(매크로가 만든 Sheet2~ 등)
        foreach (var docName in plan.IgnoredDocuments)
        {
            log.Write("info", "document_extra_ignored", new { module = docName });
            results.Add(new ModuleResult
            {
                Name = docName,
                Action = "skipped",
                Reason = "워크북 전용 문서/시트 모듈 — push 관리 대상 아님(생성/삭제 불가). src로 회수하려면 vbaloop pull",
            });
        }

        var expectedAdded = new List<string>();
        var expectedRemoved = new List<string>();
        var journal = new Journal();
        string? abortReason = null;
        bool fatalCom = false;

        // 2~5. 모듈별 주입 (실패 시 전체 중단, 저장 안 함 — 이번 push가 건드린 모듈 전부 원복 또는 poison)
        foreach (var sm in srcModules)
        {
            if (abortReason != null)
            {
                results.Add(new ModuleResult { Name = sm.Name, Action = "skipped", Reason = "선행 실패로 중단" });
                continue;
            }
            try
            {
                var r = sm.Ext == ".frm"
                    ? PushForm(sm, proj, byName, codePage, expectedAdded, expectedRemoved, log, sync, opts.ForceFormImport, journal)
                    : PushCode(sm, proj, byName, expectedAdded, log, journal);
                results.Add(r);
            }
            catch (Exception e)
            {
                // CliError만이 아니라 전부 — 비-CliError COM 예외가 루프를 탈출하면 앞선 주입이 남은 채
                // 복구 없이 끝난다. 치명 COM(연결 끊김)은 여기서 poison 기록 후 원복을 건너뛴다.
                fatalCom = SessionService.PoisonIfFatal(cfg, meta, e, log);
                results.Add(new ModuleResult { Name = sm.Name, Action = "failed", Reason = e.Message });
                abortReason = $"{sm.Name}: {e.Message}";
                log.Write("error", "push_module_failed", new { module = sm.Name, error = e.Message, fatalCom });
            }
        }

        if (abortReason != null)
            return FailWithRecovery(o, PushRecovery.Phase.Inject, abortReason, journal, proj, cfg, meta, log,
                results, fingerprint, backupId, fatalCom);

        // 4. prune 실행 (검증형 Remove: 제거 후 이름 부재 확인, 실패 시 save 전 중단)
        if (opts.Prune)
        {
            foreach (var extra in extras)
            {
                try
                {
                    journal.Pruning(extra.Name);
                    VerifiedRemove(proj, extra, log);
                }
                catch (Exception e)
                {
                    fatalCom = SessionService.PoisonIfFatal(cfg, meta, e, log);
                    results.Add(new ModuleResult { Name = extra.Name, Action = "failed", Reason = e.Message });
                    return FailWithRecovery(o, PushRecovery.Phase.Prune, $"{extra.Name}: {e.Message}", journal, proj,
                        cfg, meta, log, results, fingerprint, backupId, fatalCom);
                }
                expectedRemoved.Add(extra.Name);
                results.Add(new ModuleResult { Name = extra.Name, Action = "pruned" });
            }
        }

        // 6. 무결성 체크 (save 이전): 이름 집합 diff — <기존명>+숫자 패턴은 중복 모듈 시그니처
        HashSet<string> after = VbaProjectOps.ComponentNames(proj);
        var diff = NameSetDiff.Check(before, expectedAdded, expectedRemoved, after);
        if (diff.Unexpected.Count > 0 || diff.Missing.Count > 0)
        {
            var parts = new List<string>();
            if (diff.DuplicateSignatures.Count > 0)
                parts.Add($"중복 모듈 시그니처: {string.Join(", ", diff.DuplicateSignatures)}");
            if (diff.Unexpected.Count > 0)
                parts.Add($"기대하지 않은 이름: {string.Join(", ", diff.Unexpected)}");
            if (diff.Missing.Count > 0)
                parts.Add($"사라진 이름: {string.Join(", ", diff.Missing)}");
            return FailWithRecovery(o, PushRecovery.Phase.Integrity, $"무결성 체크 실패: {string.Join(" / ", parts)}",
                journal, proj, cfg, meta, log, results, fingerprint, backupId, fatalCom: false);
        }

        // 빈 주입 방지: 본문이 있는 모듈은 주입 후 CountOfLines > 0
        List<VbaProjectOps.ComponentInfo> compsAfter = VbaProjectOps.ListComponents(proj);
        var byNameAfter = compsAfter.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase);
        foreach (var r in results.Where(r => r.Action is "added" or "replaced" && r.BodyLineCount > 0))
        {
            int lines = Convert.ToInt32(byNameAfter[r.Name].Comp.CodeModule.CountOfLines);
            if (lines == 0)
            {
                return FailWithRecovery(o, PushRecovery.Phase.Integrity, $"무결성 체크 실패: {r.Name} 주입 후 본문이 비어 있음",
                    journal, proj, cfg, meta, log, results, fingerprint, backupId, fatalCom: false);
            }
        }

        // 7. 종결: 무결성 통과 직후 save — 저장 없이 성공을 보고하지 않는다(poison-kill 유실 방지)
        app.DisplayAlerts = false;
        DateTime mtimeBefore = File.GetLastWriteTimeUtc(cfg.WorkbookPath);
        ComRetry.Invoke(() => { wb.Save(); });

        // Save 신원 검증: read-only 워크북에서 Excel은 오류 대신 기본 저장 위치로 조용히
        // SaveAs할 수 있다. 드리프트면 캐시를 갱신하지 않아야
        // 이후 push가 no-op으로 고착되는 split-brain을 막는다.
        string actualFull = Convert.ToString(wb.FullName) ?? "";
        if (!PathDrift.SamePath(actualFull, cfg.WorkbookPath))
        {
            log.Write("error", "push_saveas_drift", new { expected = cfg.WorkbookPath, actual = actualFull });
            // 세션 워크북의 FullName이 오탈출 사본을 가리켜 세션 자체가 못 쓰는 상태 —
            // poison해서 다음 명령의 auto-ensure가 kill/재기동하게 한다(잠금 잔존 시 fail-fast).
            meta.MarkPoisoned(cfg.SessionMetaPath,
                $"push: Save가 정본 대신 {actualFull}로 SaveAs됨(read-only 잠금 의심)", null);
            o.Payload = BuildPayload(results, fingerprint, saved: false, backupId);
            o.Fail(
                $"저장 드리프트: Save가 정본({cfg.WorkbookPath})이 아니라 {actualFull}에 기록됨 — 캐시 미갱신, 오탈출 사본은 수동 삭제 필요",
                "vbaloop doctor로 정본 파일 잠금 확인 후 재시도",
                environment: true);
            return o;
        }
        if (File.GetLastWriteTimeUtc(cfg.WorkbookPath) == mtimeBefore)
            log.Write("warn", "push_save_mtime_unchanged", new { path = cfg.WorkbookPath });

        sync.LastPushAt = DateTime.UtcNow.ToString("o");
        sync.LastPushFingerprint = fingerprint;
        sync.Save(SyncCache.PathFor(cfg));
        log.Write("info", "push_done", new
        {
            counts = results.GroupBy(r => r.Action).ToDictionary(g => g.Key, g => g.Count()),
            fingerprint,
        });

        var summary = results.GroupBy(r => r.Action).Select(g => $"{g.Key} {g.Count()}");
        o.HumanLines.Add($"push 완료(저장됨): {string.Join(", ", summary)}");
        o.Payload = BuildPayload(results, fingerprint, saved: true, backupId);
        return o;
    }

    private static object BuildPayload(List<ModuleResult> results, string fingerprint, bool saved, string? backupId,
        object? recovery = null) => new
    {
        saved,
        fingerprint,
        backup = backupId,
        recovery,
        modules = results.Select(m => new { name = m.Name, action = m.Action, reason = m.Reason }),
    };

    // ---- 실패 복구: 건드린 모듈 전부 원복, 아니면 poison (PushRecovery) ----

    /// <summary>
    /// 실패 종결 공통부. PushRecovery 계획대로 역순 원복 → 하나라도 실패/불가면 poison(저장 금지, 다음
    /// 명령이 Excel kill 후 디스크 정본 재오픈). 에러 문구와 hint는 워크북의 실제 상태만 말한다.
    /// </summary>
    private static Outcome FailWithRecovery(
        Outcome o, PushRecovery.Phase phase, string reason, Journal journal, dynamic proj,
        ProjectConfig cfg, SessionMeta meta, ToolLog log, List<ModuleResult> results,
        string fingerprint, string? backupId, bool fatalCom)
    {
        var plan = PushRecovery.Decide(journal.Touches, phase);
        var restored = new List<string>();
        var failures = new List<string>();
        if (!plan.Poison && !fatalCom)
        {
            foreach (var t in plan.RestoreOrder)
            {
                try
                {
                    if (RestoreOne(t, journal, proj, log)) restored.Add(t.Module);
                    else failures.Add($"{t.Module}: 원복 후 본문이 스냅샷과 다름");
                }
                catch (Exception ex)
                {
                    failures.Add($"{t.Module}: {ex.Message}");
                    if (SessionService.PoisonIfFatal(cfg, meta, ex, log)) { fatalCom = true; break; }
                }
            }
        }

        var state = PushRecovery.Resolve(plan, journal.Touches.Count, failures.Count, fatalCom);
        string? poisonReason = plan.PoisonReason
            ?? (fatalCom ? "치명적 COM 오류(세션 연결 끊김)"
                : failures.Count > 0 ? "원복 실패: " + string.Join("; ", failures) : null);
        if (state == PushRecovery.WorkbookState.Poisoned && !meta.Poisoned)
            meta.MarkPoisoned(cfg.SessionMetaPath, $"push 부분 실패 — {poisonReason}", null);
        // 실패 모듈은 failed(사유 유지) 그대로 두고, 성공 주입됐다가 되돌린 모듈만 restored로 표기
        foreach (var r in results.Where(r => r.Action != "failed" && restored.Contains(r.Name, StringComparer.OrdinalIgnoreCase)))
            r.Action = "restored";

        log.Write(state == PushRecovery.WorkbookState.Poisoned ? "error" : "warn", "push_recovery", new
        {
            phase = phase.ToString(), state = state.ToString(), restored, failures, poisonReason, backup = backupId,
        });

        var (summary, hint) = PushRecovery.Describe(state, restored.Count, backupId, poisonReason);
        o.Payload = BuildPayload(results, fingerprint, saved: false, backupId, new
        {
            phase = phase.ToString().ToLowerInvariant(),
            workbookState = state.ToString().ToLowerInvariant(),
            restored,
            failures,
            poisonReason,
        });
        o.Fail($"push 중단(저장 안 함): {reason} — {summary}", hint, environment: fatalCom);
        return o;
    }

    /// <summary>변경 하나 원복: 본문 교체는 스냅샷 재주입+검증, 신규 생성은 검증형 제거.</summary>
    private static bool RestoreOne(PushRecovery.Touch t, Journal journal, dynamic proj, ToolLog log)
    {
        if (t.Kind == PushRecovery.TouchKind.ComponentAdded)
        {
            dynamic comp = journal.Added[t.Module];
            string actualName = Convert.ToString(comp.Name) ?? t.Module; // 이름 지정 실패 시 Module1 등
            VbaProjectOps.VerifiedRemove(proj, comp, actualName);
            log.Write("info", "push_restore_removed", new { module = t.Module, actualName });
            return true;
        }

        // CodeReplaced: 원래 컴포넌트를 이름으로 다시 찾는다(RCW 재사용 대신 현재 상태 기준)
        List<VbaProjectOps.ComponentInfo> comps = VbaProjectOps.ListComponents(proj);
        var ci = comps.FirstOrDefault(c => c.Name.Equals(t.Module, StringComparison.OrdinalIgnoreCase));
        if (ci == null) return false;
        bool ok = VbaProjectOps.RestoreSnapshot(ci.Comp.CodeModule, journal.Snapshots[t.Module]);
        log.Write(ok ? "info" : "error", "push_restore_snapshot", new { module = t.Module, verified = ok });
        return ok;
    }

    // ---- 0.5단계: push 전 백업 ----

    /// <summary>
    /// 세션 워크북의 현재(메모리) 상태를 SaveCopyAs로 _dbg/backups/&lt;id&gt;-push/에 복제하고 보존 정책을
    /// 적용한다. SaveCopyAs는 FullName/Saved 플래그를 바꾸지 않는다. keep = 0이면 null(비활성).
    /// 치명 COM 오류는 그대로 전파해 ComFailureGuard가 poison하게 하고, 그 외 실패는 CliError로 중단.
    /// </summary>
    private static string? BackupBeforeMutation(
        ProjectConfig cfg, dynamic app, dynamic wb, string fingerprint, SyncCache sync, ToolLog log)
    {
        if (cfg.BackupKeep <= 0)
        {
            log.Write("info", "push_backup_disabled", null);
            return null;
        }

        var slot = BackupStore.Allocate(cfg.BackupsDir, BackupStore.KindPush, cfg.WorkbookPath, DateTime.UtcNow);
        try
        {
            app.DisplayAlerts = false;
            ComRetry.Invoke(() => { wb.SaveCopyAs(slot.CopyPath); });
            BackupStore.Complete(slot, new BackupStore.Metadata
            {
                Command = "push",
                WorkbookPath = cfg.WorkbookPath,
                SrcFingerprint = fingerprint,
                LastPushFingerprint = sync.LastPushFingerprint,
                ExeVersion = VersionCommand.GetBuildInfo(),
            });
        }
        catch (Exception ex)
        {
            BackupStore.Discard(slot);
            log.Write("error", "push_backup_failed", new { id = slot.Id, error = ex.Message });
            if (IsFatalCom(ex)) throw;
            throw new CliError(
                $"push 전 백업 실패 — 워크북을 건드리기 전에 중단: {ex.Message}",
                "_dbg 쓰기 권한·디스크 여유·경로 길이(Excel 218자 한계) 확인. 백업 없이 진행하려면 vbaloop.toml에 [backup] keep = 0");
        }
        log.Write("info", "push_backup_done", new { id = slot.Id, path = slot.CopyPath });

        var (deleted, failed) = BackupStore.Prune(cfg.BackupsDir, cfg.BackupKeep, new[] { slot.Id }, countProtected: true);
        if (deleted.Count > 0) log.Write("info", "backup_pruned", new { deleted, keep = cfg.BackupKeep });
        if (failed.Count > 0) log.Write("warn", "backup_prune_failed", new { failed });
        return slot.Id;
    }

    private static bool IsFatalCom(Exception ex)
        => ex is System.Runtime.InteropServices.InvalidComObjectException
           || (ex is System.Runtime.InteropServices.COMException ce && ComRetry.IsPoison(ce.HResult));

    // ---- 1단계: 파일 읽기/검증 ----

    private static List<SrcModule> LoadSrc(ProjectConfig cfg, int codePage, ToolLog log)
    {
        if (!Directory.Exists(cfg.SrcRoot))
            throw new CliError($"src 루트 없음: {cfg.SrcRoot}", "vbaloop pull로 초기 생성");

        var files = Directory.EnumerateFiles(cfg.SrcRoot)
            .Where(f => Path.GetExtension(f).ToLowerInvariant() is ".bas" or ".cls" or ".frm")
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (files.Count == 0)
            throw new CliError($"src에 소스 파일이 없음: {cfg.SrcRoot}", "vbaloop pull로 초기 생성");

        // 인코딩 가드: 인코딩 불가 문자를 파일/라인 위치로 리포트하고 save 전에 실패
        var allViolations = new List<string>();
        var modules = new List<SrcModule>();

        foreach (var file in files)
        {
            var lines = CanonicalText.SplitLines(CanonicalText.ReadFileStripBom(file));
            var violations = Cp949Guard.Check(lines, codePage);
            foreach (var v in violations.Take(20))
                allViolations.Add($"{Path.GetFileName(file)}:{v.Line}:{v.Column}: '{v.Char}' ({v.CodePoint}) — CP{codePage} 표현 불가");
            if (violations.Count > 20)
                allViolations.Add($"{Path.GetFileName(file)}: 외 {violations.Count - 20}건");
            if (violations.Count > 0) continue;

            var split = VbaHeader.Split(lines, file, strict: true); // Attribute-after-code는 여기서 CliError
            string name = Path.GetFileNameWithoutExtension(file);
            string ext = Path.GetExtension(file).ToLowerInvariant();
            string? frx = null;
            if (ext == ".frm")
            {
                var frxCandidate = Path.ChangeExtension(file, ".frx");
                frx = File.Exists(frxCandidate) ? frxCandidate : null;
            }
            modules.Add(new SrcModule
            {
                Name = name,
                Ext = ext,
                FilePath = file,
                Header = split.Header,
                Body = split.Body,
                FrxPath = frx,
                IsDocumentHeader =
                    split.Header.Any(l => l.Contains("VB_PredeclaredId = True", StringComparison.OrdinalIgnoreCase)) &&
                    split.Header.Any(l => l.Contains("VB_Exposed = True", StringComparison.OrdinalIgnoreCase)),
            });
        }

        if (allViolations.Count > 0)
        {
            log.Write("error", "encoding_guard_failed", new { violations = allViolations });
            throw new CliError(
                $"인코딩 가드 실패 — CP{codePage} 밖 문자는 push 시 확정 손실: " +
                string.Join(" | ", allViolations.Take(5)) +
                (allViolations.Count > 5 ? $" 외 {allViolations.Count - 5}건" : ""),
                "해당 문자를 제거/치환 (em dash는 ChrW(&H2014) 등 코드 생성으로)");
        }

        return modules;
    }

    // ---- 2~3단계: 코드 모듈 in-place 주입 ----

    private static ModuleResult PushCode(
        SrcModule sm, dynamic proj,
        Dictionary<string, VbaProjectOps.ComponentInfo> byName,
        List<string> expectedAdded, ToolLog log, Journal journal)
    {
        if (byName.TryGetValue(sm.Name, out var ci))
        {
            if (ci.Type == VbaProjectOps.TypeMsForm)
                throw new CliError($"{sm.Name}: 워크북 컴포넌트는 폼인데 src는 {sm.Ext} — 타입 불일치");
            journal.CodeReplacing(sm.Name, VbaProjectOps.ReadModuleBody(ci.Comp.CodeModule));
            ReplaceModuleCode(ci.Comp.CodeModule, sm.Body);
            log.Write("info", "module_replaced", new { module = sm.Name });
            return new ModuleResult { Name = sm.Name, Action = "replaced", BodyLineCount = sm.Body.Count };
        }

        // 신규 모듈: 문서모듈은 VBComponents.Add로 생성 불가 → 경고+스킵(설계 §4 push 3단계)
        if (sm.Ext == ".cls" && sm.IsDocumentHeader)
        {
            log.Write("warn", "document_module_skipped", new { module = sm.Name });
            return new ModuleResult
            {
                Name = sm.Name,
                Action = "skipped",
                Reason = "문서모듈은 생성 불가 — 워크북에 대응 컴포넌트 없음",
            };
        }

        int addType = sm.Ext == ".bas" ? VbaProjectOps.TypeStdModule : VbaProjectOps.TypeClassModule;
        dynamic newComp = ComRetry.Invoke(() => proj.VBComponents.Add(addType));
        journal.ComponentAdded(sm.Name, newComp); // 이름 지정 전에 기록 — 실패해도 Module1 잔재를 회수
        ComRetry.Invoke(() => { newComp.Name = sm.Name; });
        string actualName = Convert.ToString(newComp.Name) ?? "";
        if (!actualName.Equals(sm.Name, StringComparison.OrdinalIgnoreCase))
            throw new CliError($"{sm.Name}: 신규 모듈 이름 지정 실패 (실제: {actualName})");
        expectedAdded.Add(sm.Name);
        ReplaceModuleCode(newComp.CodeModule, sm.Body);
        log.Write("info", "module_added", new { module = sm.Name });
        return new ModuleResult { Name = sm.Name, Action = "added", BodyLineCount = sm.Body.Count };
    }

    /// <summary>
    /// 원자성(설계 §4 push 2단계): 교체 전 스냅샷 → DeleteLines+AddFromString →
    /// 실패 시 스냅샷 재주입으로 원복. "DeleteLines만 성공하고 빈 채 남는" 상태를 차단.
    /// </summary>
    private static void ReplaceModuleCode(dynamic codeModule, IReadOnlyList<string> bodyLines)
        // 스냅샷 원복 + 주입 후 팬텀 () 스윕 + 논리행수 정수 가드는 공유 프리미티브가 담당
        // (트레일링 CRLF 미부착으로 왕복 끝 빈 줄 증식 방지).
        => VbaProjectOps.ReplaceCode(codeModule, bodyLines);

    // ---- 5단계: 폼 (폼 예외 절차) ----

    private static ModuleResult PushForm(
        SrcModule sm, dynamic proj,
        Dictionary<string, VbaProjectOps.ComponentInfo> byName,
        int codePage,
        List<string> expectedAdded, List<string> expectedRemoved, ToolLog log,
        SyncCache sync, bool forceImport, Journal journal)
    {
        string? srcFrxHash = sm.FrxPath != null ? HashFile(sm.FrxPath) : null;

        if (byName.TryGetValue(sm.Name, out var ci))
        {
            if (ci.Type != VbaProjectOps.TypeMsForm)
                throw new CliError($"{sm.Name}: src는 폼인데 워크북 컴포넌트 타입이 {ci.Type} — 타입 불일치");

            // in-place(코드비하인드만) 조건 = 레이아웃 텍스트 동등 AND src frx가 마지막 동기화
            // 상태 그대로. frx는 비결정적이라 fresh Export 바이트 비교가 불가하므로(실측)
            // sync.forms 캐시로 "사용자가 src frx를 의도적으로 교체했는가"만 본다.
            if (!forceImport)
            {
                var curHeader = ExportCurrentFormHeader(ci, codePage);
                bool layoutSame = CanonicalText.TrimTrailingEmpty(sm.Header)
                    .SequenceEqual(CanonicalText.TrimTrailingEmpty(curHeader), StringComparer.Ordinal);
                sync.Forms.TryGetValue(sm.Name, out var formSync);
                bool frxUnchanged = formSync?.FrxHash != null && formSync.FrxHash == srcFrxHash;

                if (layoutSame && frxUnchanged)
                {
                    journal.CodeReplacing(sm.Name, VbaProjectOps.ReadModuleBody(ci.Comp.CodeModule));
                    ReplaceModuleCode(ci.Comp.CodeModule, sm.Body);
                    UpdateFormCache(sync, sm, srcFrxHash);
                    log.Write("info", "form_code_replaced", new { module = sm.Name });
                    return new ModuleResult
                    {
                        Name = sm.Name,
                        Action = "replaced",
                        Reason = "코드비하인드만 in-place",
                        BodyLineCount = sm.Body.Count,
                    };
                }
            }

            // 레이아웃 변경 / frx 교체 / --force-form-import → 검증형 Remove 후 Import
            if (sm.FrxPath == null)
                throw new CliError($"{sm.Name}: 레이아웃/frx 반영에 필요한 .frx 파일이 src에 없음");
            journal.FormReimporting(sm.Name); // 원본 폼 바이너리가 없어 원복 불가 → 이후 실패는 poison
            VerifiedRemove(proj, ci, log);
            expectedRemoved.Add(sm.Name);
            ImportForm(sm, proj, codePage, log, journal: null);
            expectedAdded.Add(sm.Name);
            UpdateFormCache(sync, sm, srcFrxHash);
            return new ModuleResult
            {
                Name = sm.Name,
                Action = "replaced",
                Reason = forceImport ? "강제 Import(--force-form-import)" : "레이아웃/frx 변경 — Export/Import 왕복",
                BodyLineCount = sm.Body.Count,
            };
        }

        // 신규 폼
        if (sm.FrxPath == null)
            throw new CliError($"{sm.Name}: 신규 폼 Import에 필요한 .frx 파일이 src에 없음");
        ImportForm(sm, proj, codePage, log, journal);
        expectedAdded.Add(sm.Name);
        UpdateFormCache(sync, sm, srcFrxHash);
        return new ModuleResult { Name = sm.Name, Action = "added", BodyLineCount = sm.Body.Count };
    }

    /// <summary>push 후 폼 해시 캐시 갱신 — 재pull이 frx 노이즈를 생략하는 근거.</summary>
    private static void UpdateFormCache(SyncCache sync, SrcModule sm, string? srcFrxHash)
    {
        sync.Forms[sm.Name] = new SyncCache.FormSync
        {
            FrmHash = HashFile(sm.FilePath),
            FrxHash = srcFrxHash,
        };
    }

    private static string HashFile(string path)
        => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path)));

    private static IReadOnlyList<string> ExportCurrentFormHeader(
        VbaProjectOps.ComponentInfo ci, int codePage)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "vbaloop", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            string frmTemp = Path.Combine(tempDir, ci.Name + ".frm");
            ComRetry.Invoke(() => { ci.Comp.Export(frmTemp); });
            var enc = Encoding.GetEncoding(codePage);
            var lines = CanonicalText.SplitLines(enc.GetString(File.ReadAllBytes(frmTemp)));
            // frx 바이트는 비교하지 않는다 — Export마다 비결정적(실측). 레이아웃은 frm 텍스트로만.
            return VbaHeader.Split(lines, frmTemp, strict: false).Header;
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { /* 무시 */ }
        }
    }

    /// <param name="journal">신규 폼이면 Import 직후 ComponentAdded로 기록(원복 = 제거). 재Import는 null —
    /// 호출자가 이미 FormReimported(원복 불가)로 기록했다.</param>
    private static void ImportForm(SrcModule sm, dynamic proj, int codePage, ToolLog log, Journal? journal)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "vbaloop", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            // 인코딩은 GetACP 동적 조회 (폼 예외 a) — 가드는 LoadSrc에서 이미 통과
            var enc = Cp949Guard.GetStrict(codePage);
            // 헤더 끝(Attribute 블록 뒤)과 본문 끝의 빈 줄을 정규화 — Export가 이 지점에
            // 빈 줄을 누적시키는 실측 동작 때문에 트림 없이는 왕복마다 파일이 자란다
            string content = CanonicalText.ToCanonical(
                CanonicalText.TrimTrailingEmpty(sm.Header)
                    .Concat(CanonicalText.TrimTrailingEmpty(sm.Body)).ToList());
            string frmTemp = Path.Combine(tempDir, sm.Name + ".frm");
            File.WriteAllBytes(frmTemp, enc.GetBytes(content));
            File.Copy(sm.FrxPath!, Path.Combine(tempDir, sm.Name + ".frx"));

            dynamic newComp = ComRetry.Invoke(() => proj.VBComponents.Import(frmTemp));
            journal?.ComponentAdded(sm.Name, newComp);
            string actualName = Convert.ToString(newComp.Name) ?? "";

            // Import 후 중복 이름 검사 (폼 예외 c) — <name>1 패턴이면 save 전에 실패
            if (!actualName.Equals(sm.Name, StringComparison.OrdinalIgnoreCase))
                throw new CliError(
                    $"{sm.Name}: Import 결과 이름이 {actualName} — 중복 사본 생성 의심(중복 모듈 시그니처), 저장 중단");
            log.Write("info", "form_imported", new { module = sm.Name });
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { /* 무시 */ }
        }
    }

    // ---- 4단계: 검증형 Remove (완료 검증 없이는 Import/save 진입 금지) ----

    private static void VerifiedRemove(dynamic proj, VbaProjectOps.ComponentInfo ci, ToolLog log)
    {
        ComRetry.Invoke(() => { proj.VBComponents.Remove(ci.Comp); });
        HashSet<string> namesNow = VbaProjectOps.ComponentNames(proj);
        if (namesNow.Contains(ci.Name))
            throw new CliError(
                $"{ci.Name}: Remove 후에도 컴포넌트가 남아 있음(비동기 지연) — 저장 전 중단",
                "session start로 세션 재시작 후 재시도");
        log.Write("info", "module_removed", new { module = ci.Name });
    }
}
