namespace VbaLoop.Core;

/// <summary>
/// lease 오케스트레이션 — 자동 획득(acquire-or-renew) 모델(사용자 확정).
/// 명시 acquire 없이 게이트 명령이 자동 획득한다. 그 대가로
/// LeaseFile의 유휴 백스톱(process-bound도 만료+grace 회수)이 영구 점유를 방지한다.
/// 경쟁 정책: 살아있는 타 소유자 = 즉시 실패(대기 없음), 회수 가능 = 삭제 후
/// 재시도, 내 것 = 멱등 갱신.
/// </summary>
public static class LeaseService
{
    public enum AcquireOutcome { Claimed, Renewed }

    public sealed record LeaseState(
        LeaseOwner? Owner, bool FilePresent, bool Mine, bool Reclaimable, double IdleSeconds);

    public static LeaseState Inspect(ProjectConfig cfg)
    {
        string path = LeaseFile.PathFor(cfg);
        var owner = LeaseFile.Read(path);
        bool present = File.Exists(path);
        var id = AgentIdentity.Resolve(cfg.RootDir);
        var now = DateTime.UtcNow;

        bool mine = owner != null &&
                    LeaseFile.IsOwnedBy(owner, id.Thread, Environment.MachineName, Environment.UserName);
        bool reclaimable = present && LeaseFile.IsAutoReclaimable(
            owner, now, IsOwnerAlive, Environment.MachineName,
            SafeLastWriteUtc(path));
        double idle = owner != null ? LeaseFile.IdleSeconds(owner, now) : double.NaN;
        return new LeaseState(owner, present, mine, reclaimable, idle);
    }

    /// <summary>
    /// 게이트 공용 진입점. 자유 → claim / 내 것 → 갱신 / 회수 가능 → 회수 후 claim /
    /// 살아있는 타 소유자 → CliError(exit 2, 대기 없음 — 에이전트 루프가 재시도 주체).
    /// </summary>
    public static AcquireOutcome AcquireOrRenew(ProjectConfig cfg, string command, ToolLog? log = null)
    {
        string path = LeaseFile.PathFor(cfg);
        var id = AgentIdentity.Resolve(cfg.RootDir);
        var now = DateTime.UtcNow;

        for (int attempt = 0; attempt < 6; attempt++)
        {
            var owner = LeaseFile.Read(path);

            if (owner == null && !File.Exists(path))
            {
                var fresh = NewOwner(cfg, id, command, now);
                if (LeaseFile.TryClaim(path, fresh) == LeaseFile.ClaimResult.Claimed)
                {
                    log?.Write("info", "lease_claimed", new { ownerThread = fresh.OwnerThread, source = fresh.OwnerSource });
                    return AcquireOutcome.Claimed;
                }
                continue; // 경쟁 패배 → 재평가
            }

            if (owner != null && LeaseFile.IsOwnedBy(owner, id.Thread, Environment.MachineName, Environment.UserName))
            {
                // env 소유자는 같은 세션 id라도 에이전트 프로세스가 바뀔 수 있다(세션 resume) —
                // 옛 PID가 남으면 살아있는 세션의 lease가 사망 회수 대상이 된다
                owner.AgentPid = id.AgentPid;
                owner.AgentStartTime = id.AgentStartTime;
                LeaseFile.Touch(owner, command, now);
                LeaseFile.Save(path, owner);
                return AcquireOutcome.Renewed;
            }

            if (LeaseFile.IsAutoReclaimable(owner, now, IsOwnerAlive,
                    Environment.MachineName, SafeLastWriteUtc(path)))
            {
                LeaseFile.Delete(path);
                log?.Write("warn", "lease_reclaimed", new
                {
                    prevOwner = owner?.OwnerThread,
                    prevSource = owner?.OwnerSource,
                    idleSeconds = owner != null ? LeaseFile.IdleSeconds(owner, now) : (double?)null,
                });
                continue; // 회수 후 재claim
            }

            throw BusyError(owner!, now);
        }
        throw new CliError("lease acquisition race not resolved", "Check state with vbaloop lease status, then retry");
    }

    /// <summary>session stop 특례의 소유 판정용: 내 것이면 해제. 없으면 no-op.</summary>
    public static bool ReleaseIfMine(ProjectConfig cfg, ToolLog? log = null)
    {
        var state = Inspect(cfg);
        if (state.Owner == null || !state.Mine) return false;
        LeaseFile.Delete(LeaseFile.PathFor(cfg));
        log?.Write("info", "lease_released", new { ownerThread = state.Owner.OwnerThread });
        return true;
    }

    public static CliError BusyError(LeaseOwner owner, DateTime utcNow)
    {
        double idle = LeaseFile.IdleSeconds(owner, utcNow);
        string idleText = double.IsNaN(idle) ? "?" : $"{(int)idle}s";
        return new CliError(
            $"Another agent holds the lease: {owner.OwnerThread} (source={owner.OwnerSource}, idle {idleText}, last command {owner.LastCommand})",
            "Wait if the owner is still working; if the lease is clearly abandoned, run vbaloop lease break --force");
    }

    private static LeaseOwner NewOwner(ProjectConfig cfg, AgentIdentity.Identity id, string command, DateTime now)
    {
        var owner = new LeaseOwner
        {
            OwnerThread = id.Thread,
            OwnerSource = id.Source,
            AgentPid = id.AgentPid,
            AgentStartTime = id.AgentStartTime,
            Host = Environment.MachineName,
            User = Environment.UserName,
            ProjectPath = cfg.RootDir,
            WorkbookFullName = cfg.WorkbookPath,
            AcquiredAt = now.ToString("o"),
        };
        LeaseFile.Touch(owner, command, now);
        return owner;
    }

    /// <summary>회수 판정용 생존 — 경계 프로세스(sihost 등)는 vbaloop를 호출하는 에이전트일 수 없으므로
    /// 죽은 것으로 본다. 현 계보 탐색은 경계를 후보로 고르지 않으니 구버전이 남긴 lease에만 해당.</summary>
    public static bool IsOwnerAlive(int pid, string? startTimeIso)
        => IsOwnerAlive(pid, startTimeIso, ProcessIdentity.IsAlive, AgentIdentity.IsBoundaryProcess);

    /// <summary>의존성 주입 버전 (단위 테스트 대상).</summary>
    public static bool IsOwnerAlive(int pid, string? startTimeIso,
        Func<int, string?, bool> isProcessAlive, Func<int, bool> isBoundaryProcess)
        => isProcessAlive(pid, startTimeIso) && !isBoundaryProcess(pid);

    private static DateTime? SafeLastWriteUtc(string path)
    {
        try
        {
            return File.Exists(path) ? File.GetLastWriteTimeUtc(path) : null;
        }
        catch
        {
            return null;
        }
    }
}
