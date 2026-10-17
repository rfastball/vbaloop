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
            throw new CliError("rollback format: exactly one of --list | --latest | --backup <id>",
                "Check the backup list first with vbaloop rollback --list");

        if (opts.List) return List(o, cfg);

        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "rollback", log);

        string? block = WorkbookRollback.SessionBlockReason(
            SessionMeta.Load(cfg.SessionMetaPath), ProcessIdentity.IsAlive);
        if (block != null)
        {
            log.Write("warn", "rollback_refused_session", new { reason = block });
            throw new CliError($"rollback refused: {block}",
                "Run vbaloop session stop, then retry (right after a push failure, prefer --no-save: session stop --no-save)",
                environment: true);
        }

        var target = BackupStore.Select(BackupStore.List(cfg.BackupsDir), opts.Latest, opts.Backup);
        var meta = target.Meta!;
        if (!PathDrift.SamePath(meta.WorkbookPath, cfg.WorkbookPath))
            throw new CliError(
                $"Backup {target.Id} belongs to a different workbook: {meta.WorkbookPath} (current canonical workbook: {cfg.WorkbookPath})",
                "Check whether [excel].workbook in vbaloop.toml changed — restoring to a different file is not supported");

        if (WorkbookRollback.IsLocked(cfg.WorkbookPath))
            throw new CliError($"workbook is open in another process: {cfg.WorkbookPath}",
                "Ask the user to close that Excel window, or find the locking process with vbaloop doctor, then retry", environment: true);

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
                throw new CliError($"pre-rollback safety backup failed — not restoring: {ex.Message}",
                    "Check _dbg write permission/free disk space");
            }
            safetyId = slot.Id;
            log.Write("info", "rollback_safety_backup", new { id = safetyId });
        }
        else
        {
            warnings.Add($"canonical workbook missing, restored without a safety backup: {cfg.WorkbookPath}");
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
            "src may now differ from the workbook — the next push re-injects all of src and overwrites this restore. " +
            "Compare: vbaloop diff / adopt the restored workbook into src: vbaloop pull --force";
        warnings.Add(srcWarning);

        o.HumanLines.Add($"rollback done: {target.Id} ({target.Kind}, {Local(target.CreatedUtc)}) -> {cfg.WorkbookPath}");
        if (safetyId != null)
            o.HumanLines.Add($"  safety backup: {safetyId} (undo: vbaloop rollback --backup {safetyId})");
        o.HumanLines.Add("  sync cache invalidated: next push re-injects everything, plain pull is blocked (--force required)");
        foreach (var w in warnings) o.HumanLines.Add($"  warning: {w}");
        o.Hint = "Compare with src via vbaloop diff, then pull --force (adopt the restore) or push (re-apply src)";
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
        o.HumanLines.Add($"{entries.Count(e => e.Complete)} backups ({cfg.BackupsDir}, keep {(cfg.BackupKeep > 0 ? cfg.BackupKeep.ToString() : "disabled")})");
        foreach (var e in entries)
        {
            string size = e.Meta != null ? $"{e.Meta.SizeBytes / 1024.0:0.#} KB" : "-";
            string extra = e.Meta?.RestoringFrom != null ? $" (taken just before a restore, -> {e.Meta.RestoringFrom})" : "";
            o.HumanLines.Add($"  {e.Id}  {Local(e.CreatedUtc)}  {size}{(e.Complete ? "" : "  [incomplete — cannot restore]")}{extra}");
        }
        if (entries.Count == 0)
            o.HumanLines.Add("  (none — push creates one just before injection)");
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
