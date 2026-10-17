namespace VbaLoop.Core;

/// <summary>
/// rollback의 Excel 무관 부분(설계 §4 rollback): 세션 차단 규칙, 정본 파일의 원자 교체, sync 캐시 무효화.
/// COM을 전혀 쓰지 않는다 — 세션이 살아 있으면 애초에 거부하므로 워크북 파일만 다룬다.
/// </summary>
public static class WorkbookRollback
{
    /// <summary>
    /// 세션 차단 규칙(순수): 세션 메타가 있고 기록된 Excel 프로세스(pid+시작시각)가 살아 있으면 거부.
    /// active뿐 아니라 poisoned/stale이라도 프로세스가 살아 있으면 파일을 물고 있거나 나중에 Save로
    /// 복원을 덮어쓸 수 있다 — session stop(정리/kill)이 먼저다. 프로세스가 죽은 잔여 메타는 통과:
    /// 다음 명령의 auto-ensure가 메타를 치우고 복원된 파일을 새로 연다.
    /// 반환: 거부 사유(통과면 null).
    /// </summary>
    public static string? SessionBlockReason(SessionMeta? meta, Func<int, string?, bool> isAlive)
    {
        if (meta == null) return null;
        if (!isAlive(meta.Pid, meta.ExcelStartTime)) return null;
        string state = meta.Poisoned ? "poisoned" : "active";
        return $"Session Excel (PID {meta.Pid}, {state}) is alive — the open workbook would overwrite the restore or lock the file";
    }

    /// <summary>
    /// 정본 교체: 백업 사본 → 정본과 같은 폴더의 임시 파일(해시 대조) → File.Replace 원자 교체
    /// (정본이 없으면 File.Move) → 정본 해시 대조. delete-then-copy는 절대 쓰지 않는다 — 중간 실패가
    /// 정본 부재를 만들기 때문. File.Replace가 실패하고 임시/정본이 모두 남아 있으면 같은 볼륨
    /// rename(MoveFileEx REPLACE_EXISTING)으로 한 번 더 시도한다. 임시 파일은 어떤 경로든 정리.
    /// expectedSha256가 비어 있으면(구 메타) 사본 해시를 기준으로 삼는다.
    /// </summary>
    public static void ReplaceWorkbook(string backupCopyPath, string? expectedSha256, string workbookPath)
    {
        if (!File.Exists(backupCopyPath))
            throw new CliError($"backup copy not found: {backupCopyPath}");
        string actual = BackupStore.HashFile(backupCopyPath);
        if (!string.IsNullOrEmpty(expectedSha256) && !actual.Equals(expectedSha256, StringComparison.OrdinalIgnoreCase))
            throw new CliError($"Backup copy hash mismatch — a damaged backup is not restored: {backupCopyPath}",
                "Choose another backup (rollback --list)");

        string dir = Path.GetDirectoryName(Path.GetFullPath(workbookPath))!;
        string tmp = Path.Combine(dir, $"{Path.GetFileName(workbookPath)}.vbaloop-rollback.{Environment.ProcessId}.tmp");
        try
        {
            File.Copy(backupCopyPath, tmp, overwrite: true);
            if (!BackupStore.HashFile(tmp).Equals(actual, StringComparison.OrdinalIgnoreCase))
                throw new CliError($"temporary copy hash mismatch (suspect disk write error): {tmp}", environment: true);

            if (!File.Exists(workbookPath))
            {
                File.Move(tmp, workbookPath);
            }
            else
            {
                try
                {
                    File.Replace(tmp, workbookPath, destinationBackupFileName: null, ignoreMetadataErrors: true);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException)
                {
                    // ReplaceFile 미지원 파일시스템(일부 네트워크 공유) 대비. 잠금이면 이 rename도 실패한다.
                    if (!File.Exists(tmp) || !File.Exists(workbookPath)) throw;
                    File.Move(tmp, workbookPath, overwrite: true);
                }
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new CliError(
                $"Workbook replacement failed — canonical workbook unchanged: {ex.Message}",
                "If another Excel window has the workbook open, close it and retry (find the locking process with vbaloop doctor)",
                environment: true);
        }
        finally
        {
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch { /* 잔재만 — 정본과 무관 */ }
        }

        string after = BackupStore.HashFile(workbookPath);
        if (!after.Equals(actual, StringComparison.OrdinalIgnoreCase))
            throw new CliError($"Canonical workbook hash differs from the backup after restore: {workbookPath}",
                "Roll back to the safety backup (pre-rollback) or check the file manually", environment: true);
    }

    /// <summary>
    /// 정본이 다른 프로세스에 열려 있는지(배타 오픈 실패) — 안전 백업을 만들기 전에 조기 거부해
    /// 쓸모없는 pre-rollback 슬롯이 쌓이는 것을 막는다. 파일이 없으면 false.
    /// </summary>
    public static bool IsLocked(string path)
    {
        if (!File.Exists(path)) return false;
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }
    }

    /// <summary>pull 가드 무력화용 표식 접두 — 어떤 실제 지문(hex)과도 일치하지 않는다.</summary>
    public const string PullBaselinePrefix = "rollback:";

    /// <summary>
    /// 복원 후 sync 캐시 무효화: (1) push 지문/시각을 지워 다음 push가 전량 재주입이 되게 하고,
    /// (2) 폼 frx 캐시를 비워 레이아웃이 Export/Import 왕복으로 다시 반영되게 하며, (3) pull 기준선을
    /// 실제 지문과 절대 일치하지 않는 표식으로 바꿔 plain pull이 막히게 한다 — 복원으로 워크북에서
    /// 사라진 편집이 src에만 남아 있을 수 있으므로, src 덮어쓰기는 pull --force라는 명시 선택이어야 한다.
    /// </summary>
    public static void InvalidateSync(SyncCache sync, string restoredId)
    {
        sync.LastPushFingerprint = null;
        sync.LastPushAt = null;
        sync.LastPullFingerprint = PullBaselinePrefix + restoredId;
        sync.Forms.Clear();
    }
}
