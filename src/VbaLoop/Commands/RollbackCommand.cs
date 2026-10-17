using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop rollback (설계 §4 rollback): push 전 백업으로 정본 워크북 파일을 되돌린다.
/// --list / --latest / --backup &lt;id&gt;. Excel을 띄우지 않는 순수 파일 작업이다 — 그래서 세션
/// Excel이 살아 있으면 exit 3으로 거부한다(열린 워크북이 복원을 덮어쓰거나 파일을 잠근다).
/// 순서: 차단 규칙 → 대상 선택 → 잠금 확인 → pre-rollback 안전 백업 → 원자 교체 → sync 무효화 → prune.
/// </summary>
public static class RollbackCommand
{
    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "rollback" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "rollback");
        o.LogRecords = log.Records;

        int forms = (opts.List ? 1 : 0) + (opts.Latest ? 1 : 0) + (opts.Backup != null ? 1 : 0);
        if (forms != 1)
            throw new CliError("rollback 형식: --list | --latest | --backup <id> 중 정확히 하나",
                "vbaloop rollback --list로 백업 목록부터 확인");

        if (opts.List) return List(o, cfg);

        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "rollback", log);

        string? block = WorkbookRollback.SessionBlockReason(
            SessionMeta.Load(cfg.SessionMetaPath), ProcessIdentity.IsAlive);
        if (block != null)
        {
            log.Write("warn", "rollback_refused_session", new { reason = block });
            throw new CliError($"rollback 거부: {block}",
                "vbaloop session stop 후 재시도 (push 실패 직후라면 --no-save 권장: session stop --no-save)",
                environment: true);
        }

        var target = BackupStore.Select(BackupStore.List(cfg.BackupsDir), opts.Latest, opts.Backup);
        var meta = target.Meta!;
        if (!PathDrift.SamePath(meta.WorkbookPath, cfg.WorkbookPath))
            throw new CliError(
                $"백업 {target.Id}는 다른 워크북의 것: {meta.WorkbookPath} (현재 정본: {cfg.WorkbookPath})",
                "vbaloop.toml의 [excel].workbook이 바뀌었는지 확인 — 다른 파일로의 복원은 지원하지 않음");

        if (WorkbookRollback.IsLocked(cfg.WorkbookPath))
            throw new CliError($"워크북이 다른 프로세스에 열려 있음: {cfg.WorkbookPath}",
                "그 Excel 창을 닫거나 vbaloop doctor로 잠금 프로세스 확인 후 재시도", environment: true);

        var warnings = new List<string>();
        var sync = SyncCache.Load(SyncCache.PathFor(cfg));

        // 안전 백업: rollback 자체를 되돌릴 근거. [backup] keep = 0이어도 만든다 — 명시적 복원 요청에서
        // 현재 파일을 근거 없이 덮어쓰는 것은 비활성 설정의 의도(push 오버헤드 제거) 밖이다.
        string? safetyId = null;
        if (File.Exists(cfg.WorkbookPath))
        {
            var slot = BackupStore.Allocate(cfg.BackupsDir, BackupStore.KindPreRollback, cfg.WorkbookPath, DateTime.UtcNow);
            try
            {
                File.Copy(cfg.WorkbookPath, slot.CopyPath);
                BackupStore.Complete(slot, new BackupStore.Metadata
                {
                    Command = "rollback",
                    WorkbookPath = cfg.WorkbookPath,
                    SrcFingerprint = Directory.Exists(cfg.SrcRoot) ? Fingerprint.Compute(cfg.SrcRoot) : null,
                    LastPushFingerprint = sync.LastPushFingerprint,
                    ExeVersion = VersionCommand.GetBuildInfo(),
                    RestoringFrom = target.Id,
                });
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or CliError)
            {
                BackupStore.Discard(slot);
                throw new CliError($"pre-rollback 안전 백업 실패 — 복원하지 않음: {ex.Message}",
                    "_dbg 쓰기 권한/디스크 여유 확인");
            }
            safetyId = slot.Id;
            log.Write("info", "rollback_safety_backup", new { id = safetyId });
        }
        else
        {
            warnings.Add($"정본 워크북이 없어 안전 백업 없이 복원함: {cfg.WorkbookPath}");
        }

        WorkbookRollback.ReplaceWorkbook(target.WorkbookCopyPath!, meta.Sha256, cfg.WorkbookPath);
        log.Write("info", "rollback_done", new { restored = target.Id, safety = safetyId });

        WorkbookRollback.InvalidateSync(sync, target.Id);
        sync.Save(SyncCache.PathFor(cfg));

        var protect = new List<string> { target.Id };
        if (safetyId != null) protect.Add(safetyId);
        var (deleted, failed) = BackupStore.Prune(cfg.BackupsDir, cfg.BackupKeep, protect);
        if (deleted.Count > 0) log.Write("info", "backup_pruned", new { deleted, keep = cfg.BackupKeep });
        if (failed.Count > 0) log.Write("warn", "backup_prune_failed", new { failed });

        const string srcWarning =
            "src가 이제 워크북과 다를 수 있음 — 다음 push는 src를 전량 재주입해 이 복원을 덮어쓴다. " +
            "비교: vbaloop diff / 복원된 워크북을 src로 채택: vbaloop pull --force";
        warnings.Add(srcWarning);

        o.HumanLines.Add($"rollback 완료: {target.Id} ({target.Kind}, {Local(target.CreatedUtc)}) -> {cfg.WorkbookPath}");
        if (safetyId != null)
            o.HumanLines.Add($"  안전 백업: {safetyId} (되돌리기: vbaloop rollback --backup {safetyId})");
        o.HumanLines.Add("  sync 캐시 무효화: 다음 push는 전량 재주입, plain pull은 차단(--force 필요)");
        foreach (var w in warnings) o.HumanLines.Add($"  주의: {w}");
        o.Hint = "vbaloop diff로 src와 비교 후 pull --force(복원 채택) 또는 push(src 재반영)";
        o.Payload = new
        {
            restored = target.Id,
            kind = target.Kind,
            createdAtUtc = meta.CreatedAtUtc,
            workbook = cfg.WorkbookPath,
            sha256 = meta.Sha256,
            safetyBackup = safetyId,
            syncInvalidated = true,
            pruned = deleted,
            warnings,
        };
        return o;
    }

    private static Outcome List(Outcome o, ProjectConfig cfg)
    {
        var entries = BackupStore.List(cfg.BackupsDir);
        o.HumanLines.Add($"백업 {entries.Count(e => e.Complete)}개 ({cfg.BackupsDir}, 보존 {(cfg.BackupKeep > 0 ? cfg.BackupKeep + "개" : "비활성")})");
        foreach (var e in entries)
        {
            string size = e.Meta != null ? $"{e.Meta.SizeBytes / 1024.0:0.#} KB" : "-";
            string extra = e.Meta?.RestoringFrom != null ? $" (복원 직전, -> {e.Meta.RestoringFrom})" : "";
            o.HumanLines.Add($"  {e.Id}  {Local(e.CreatedUtc)}  {size}{(e.Complete ? "" : "  [미완 — 복원 불가]")}{extra}");
        }
        if (entries.Count == 0)
            o.HumanLines.Add("  (없음 — push가 주입 직전에 만든다)");
        o.Payload = new
        {
            dir = cfg.BackupsDir,
            keep = cfg.BackupKeep,
            backups = entries.Select(e => new
            {
                id = e.Id,
                kind = e.Kind,
                complete = e.Complete,
                createdAtUtc = e.CreatedUtc.ToString("o"),
                sizeBytes = e.Meta?.SizeBytes,
                srcFingerprint = e.Meta?.SrcFingerprint,
                lastPushFingerprint = e.Meta?.LastPushFingerprint,
                restoringFrom = e.Meta?.RestoringFrom,
                exeVersion = e.Meta?.ExeVersion,
            }),
        };
        return o;
    }

    private static string Local(DateTime utc) => utc.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
}
