using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class LeaseFileTests
{
    private static readonly DateTime Now = new(2026, 7, 8, 12, 0, 0, DateTimeKind.Utc);
    private const string Host = "THIS-PC";

    private static LeaseOwner Owner(
        string source = "process", int? pid = 100, string host = Host,
        int expiresInSec = 600, string? startTime = "ST1")
        => new()
        {
            OwnerSource = source,
            AgentPid = pid,
            AgentStartTime = startTime,
            Host = host,
            OwnerThread = "proc:THIS-PC:100:ST1",
            User = "user1",
            HeartbeatAt = Now.AddSeconds(-60).ToString("o"),
            ExpiresAt = Now.AddSeconds(expiresInSec).ToString("o"),
        };

    private static bool Reclaimable(LeaseOwner? o, bool alive = true, DateTime? fileAge = null)
        => LeaseFile.IsAutoReclaimable(o, Now, (_, _) => alive, Host, fileAge ?? Now.AddMinutes(-5));

    // ---- 회수 결정 테이블 (vh 이식 + 유휴 백스톱 수정) ----

    [Fact]
    public void NullOwner_ReclaimableOnlyAfterMinAge()
    {
        Assert.True(Reclaimable(null, fileAge: Now.AddSeconds(-11)));
        Assert.False(Reclaimable(null, fileAge: Now.AddSeconds(-3))); // 방금 쓰인 파일 강탈 금지
        Assert.True(LeaseFile.IsAutoReclaimable(null, Now, (_, _) => true, Host, null));
    }

    [Fact]
    public void CrossHost_NeverReclaimable()
    {
        // 죽은 프로세스 + 만료 + grace 경과여도 다른 머신이면 절대 회수 금지
        var o = Owner(host: "OTHER-PC", expiresInSec: -99999);
        Assert.False(Reclaimable(o, alive: false));
    }

    [Fact]
    public void ProcessBound_DeadProcess_ReclaimedImmediately()
    {
        var o = Owner(expiresInSec: 600); // TTL 안 지났어도
        Assert.True(Reclaimable(o, alive: false));
    }

    [Fact]
    public void ProcessBound_AliveAndActive_NeverReclaimed()
    {
        // 활동 중(만료 전) 에이전트는 절대 안 뺏김
        Assert.False(Reclaimable(Owner(expiresInSec: 600), alive: true));
        // 만료됐지만 grace 이내 — 아직 유지
        Assert.False(Reclaimable(Owner(expiresInSec: -100), alive: true));
    }

    [Fact]
    public void ProcessBound_AliveButIdlePastGrace_Reclaimed()
    {
        // 유휴 백스톱(자동 획득 모델 정합 수정): 살아있어도 만료+grace(1800s) 경과면 회수
        var o = Owner(expiresInSec: -(LeaseFile.ReclaimGraceSeconds + 60));
        Assert.True(Reclaimable(o, alive: true));
    }

    [Fact]
    public void EnvOrWeak_ReclaimedOnlyPastExpiryPlusGrace()
    {
        Assert.False(Reclaimable(Owner(source: "env", pid: null), alive: false));
        Assert.False(Reclaimable(Owner(source: "weak", pid: null, expiresInSec: -100)));
        Assert.True(Reclaimable(Owner(source: "env", pid: null,
            expiresInSec: -(LeaseFile.ReclaimGraceSeconds + 60))));
    }

    [Fact]
    public void UnparsableExpiry_TreatedAsExpired()
    {
        var o = Owner(source: "weak", pid: null);
        o.ExpiresAt = "not-a-date";
        Assert.True(Reclaimable(o));
    }

    // ---- 원자 claim/save 시맨틱 ----

    [Fact]
    public void TryClaim_SecondClaimerLoses()
    {
        var path = Path.Combine(Path.GetTempPath(), "vbaloop-test", Path.GetRandomFileName(), "lease.json");
        try
        {
            Assert.Equal(LeaseFile.ClaimResult.Claimed, LeaseFile.TryClaim(path, Owner()));
            Assert.Equal(LeaseFile.ClaimResult.AlreadyHeld, LeaseFile.TryClaim(path, Owner()));

            var read = LeaseFile.Read(path);
            Assert.NotNull(read);
            Assert.Equal("proc:THIS-PC:100:ST1", read!.OwnerThread);

            // tmp 잔재가 남지 않아야 함
            Assert.Single(Directory.GetFiles(Path.GetDirectoryName(path)!));

            LeaseFile.Delete(path);
            Assert.Equal(LeaseFile.ClaimResult.Claimed, LeaseFile.TryClaim(path, Owner()));
        }
        finally
        {
            Directory.Delete(Path.GetDirectoryName(path)!, recursive: true);
        }
    }

    [Fact]
    public void Save_OverwritesHeldLease_AndTouchUpdatesFields()
    {
        var path = Path.Combine(Path.GetTempPath(), "vbaloop-test", Path.GetRandomFileName(), "lease.json");
        try
        {
            var owner = Owner();
            LeaseFile.TryClaim(path, owner);
            LeaseFile.Touch(owner, "push", Now);
            LeaseFile.Save(path, owner);

            var read = LeaseFile.Read(path)!;
            Assert.Equal("push", read.LastCommand);
            Assert.Equal(Environment.ProcessId, read.LastCommandPid);
            Assert.Equal(Now.AddSeconds(LeaseFile.TtlSeconds).ToString("o"), read.ExpiresAt);
        }
        finally
        {
            Directory.Delete(Path.GetDirectoryName(path)!, recursive: true);
        }
    }

    [Fact]
    public void IsOwnedBy_RequiresThreadHostUser()
    {
        var o = Owner();
        Assert.True(LeaseFile.IsOwnedBy(o, "proc:THIS-PC:100:ST1", "this-pc", "USER1"));
        Assert.False(LeaseFile.IsOwnedBy(o, "proc:THIS-PC:999:ST1", Host, "user1"));
        Assert.False(LeaseFile.IsOwnedBy(o, "proc:THIS-PC:100:ST1", "OTHER", "user1"));
        Assert.False(LeaseFile.IsOwnedBy(o, "proc:THIS-PC:100:ST1", Host, "other"));
    }
}
