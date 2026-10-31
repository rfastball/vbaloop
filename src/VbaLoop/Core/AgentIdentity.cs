using System.Diagnostics;
using System.Text;

namespace VbaLoop.Core;

/// <summary>
/// lease 소유자 신원 결정.
/// 우선순위: env id > 프로세스 계보 > weak. 계보는 allowlist가 아니라 denylist 방식 —
/// boundary(데스크톱/서비스 경계)에서 정지, shell은 통과, "경계 직전 가장 바깥의 non-shell
/// 조상"이 장수 에이전트 호스트다(claude/codex/node 등 이름 하드코딩 아님).
/// vh에 없던 보강 3건(검토 확정): hop 0(자기 자신, vbaloop.exe는 non-shell이라 자기에게
/// lease가 묶여 즉사하는 함정) 후보 제외, 부모 StartTime &gt; 자식 StartTime이면 PID 재사용으로
/// 체인 중단, StartTime 읽기 실패 조상도 중단(검증 불가 조상보다 weak 강등이 안전).
/// </summary>
public static class AgentIdentity
{
    public sealed record Identity(string Thread, string Source, int? AgentPid, string? AgentStartTime);

    public sealed record AgentHost(int Pid, string Name, string StartTimeIso);

    /// <summary>계보 탐색을 멈추는 데스크톱/서비스 경계 프로세스.
    /// sihost: 셸 활성화로 뜬 앱(Claude 데스크톱 등)의 부모 — 로그아웃까지 살아 있어 소유자가 되면
    /// 생존 회수가 영영 성립하지 않는다(2026-10-31 실측).</summary>
    public static readonly string[] BoundaryNames =
        { "explorer", "sihost", "services", "wininit", "winlogon", "svchost", "lsass", "csrss", "smss", "system", "idle" };

    public static readonly string[] ShellNames =
        { "cmd", "powershell", "pwsh", "conhost", "windowsterminal", "openconsole", "bash", "sh", "wsl", "wslhost" };

    /// <summary>세션 id 환경변수. PidEnv가 있으면 그 값(에이전트 프로세스 PID)을 생존 판정에 묶어
    /// env 소유자도 세션 종료 즉시 회수되게 한다.</summary>
    public sealed record EnvId(string Name, string? PidEnv = null);

    /// <summary>우선순위 순. CLAUDE_CODE_SESSION_ID/CLAUDE_PID는 Claude Code가 Bash 도구 env에 주는 실제 이름
    /// (2026-10-31 실측) — 부모가 먼저 죽은 백그라운드 셸도 env는 물려받아 계보 단절과 무관하다.
    /// CLAUDE_SESSION_ID는 구 이름 호환.</summary>
    public static readonly EnvId[] EnvIds =
    {
        new("VBALOOP_SESSION_ID"), new("CODEX_THREAD_ID"), new("CODEX_SESSION_ID"),
        new("CLAUDE_CODE_SESSION_ID", "CLAUDE_PID"), new("CLAUDE_SESSION_ID"), new("AGENT_THREAD_ID"),
    };

    public static Identity Resolve(string projectPath)
        => Resolve(projectPath, Environment.GetEnvironmentVariable,
            ProcessTree.Snapshot(), Environment.ProcessId, StartTimeUtcOf);

    /// <summary>의존성 주입 버전 (단위 테스트 대상).</summary>
    public static Identity Resolve(
        string projectPath,
        Func<string, string?> getEnv,
        IReadOnlyDictionary<int, ProcessTree.ProcEntry> table,
        int startPid,
        Func<int, DateTime?> startTimeUtcOf)
    {
        // 1순위: 호출자가 주는 유니크 세션 id (env). 프로세스 신원보다 우선(vh와 동일).
        foreach (var env in EnvIds)
        {
            var val = getEnv(env.Name);
            if (string.IsNullOrWhiteSpace(val)) continue;
            var (pid, start) = EnvAgentProcess(env.PidEnv, getEnv, startPid, startTimeUtcOf);
            return new Identity($"env:{env.Name}:{val.Trim()}", "env", pid, start);
        }

        // 2순위: 프로세스 계보 — 장수 에이전트 호스트의 (PID, StartTime)
        var host = FindAgentHost(table, startPid, startTimeUtcOf);
        if (host != null)
            return new Identity(
                $"proc:{Environment.MachineName}:{host.Pid}:{host.StartTimeIso}",
                "process", host.Pid, host.StartTimeIso);

        // 3순위: weak — 세션 구분 불가 토큰 (host:user:projectKey)
        return new Identity(
            $"local:{Environment.MachineName}:{Environment.UserName}:{ProjectKey(projectPath)}",
            "weak", null, null);
    }

    /// <summary>경계 직전 가장 바깥의 non-shell 조상 탐색 (순수 로직, 테이블 주입).</summary>
    public static AgentHost? FindAgentHost(
        IReadOnlyDictionary<int, ProcessTree.ProcEntry> table,
        int startPid,
        Func<int, DateTime?> startTimeUtcOf)
    {
        AgentHost? candidate = null;
        var seen = new HashSet<int>();
        int cur = startPid;
        DateTime? curStart = startTimeUtcOf(cur);
        if (curStart == null) return null;

        for (int hops = 0; hops < 24; hops++)
        {
            if (cur <= 0 || !seen.Add(cur)) break;
            if (!table.TryGetValue(cur, out var entry)) break;
            if (BoundaryNames.Contains(entry.Name)) break;

            // hop 0(자기 자신) 제외 + vbaloop 이름 방어(다른 vbaloop 프로세스가 조상일 일은
            // 없지만 재귀 호출 대비) — shell은 후보가 아니되 체인은 계속 탄다
            if (hops > 0 && !ShellNames.Contains(entry.Name) && entry.Name != "vbaloop")
                candidate = new AgentHost(cur, entry.Name, curStart.Value.ToString("o"));

            int parent = entry.ParentPid;
            if (parent <= 0) break;
            var parentStart = startTimeUtcOf(parent);
            if (parentStart == null) break;          // 검증 불가 조상 → 중단
            if (parentStart > curStart) break;       // 부모가 자식보다 최신 = PID 재사용(가짜 부모)

            cur = parent;
            curStart = parentStart;
        }
        return candidate;
    }

    /// <summary>env가 가리키는 에이전트 프로세스의 (PID, StartTime). 조상 확인은 하지 않는다 — 고아가 된
    /// 백그라운드 체인에선 살아 있는 에이전트에도 계보가 닿지 않기 때문. 대신 PID 재사용 방어로
    /// 에이전트가 자기보다 먼저 시작했어야 한다. 해석 불가면 (null, null) = 유휴 백스톱만.</summary>
    private static (int? Pid, string? StartTimeIso) EnvAgentProcess(
        string? pidEnv, Func<string, string?> getEnv, int startPid, Func<int, DateTime?> startTimeUtcOf)
    {
        if (pidEnv == null || !int.TryParse(getEnv(pidEnv)?.Trim(), out int pid) || pid <= 0 || pid == startPid)
            return (null, null);
        var agentStart = startTimeUtcOf(pid);
        var selfStart = startTimeUtcOf(startPid);
        if (agentStart == null || (selfStart != null && agentStart > selfStart)) return (null, null);
        return (pid, agentStart.Value.ToString("o"));
    }

    /// <summary>pid의 현재 이미지가 경계 프로세스인가 — 구버전이 경계 프로세스(sihost)에 귀속시킨
    /// lease를 회수하기 위한 판정. 읽기 실패는 false(보수적).</summary>
    public static bool IsBoundaryProcess(int pid)
    {
        try
        {
            using var p = Process.GetProcessById(pid);
            return BoundaryNames.Contains(p.ProcessName.ToLowerInvariant());
        }
        catch
        {
            return false;
        }
    }

    private static DateTime? StartTimeUtcOf(int pid)
    {
        try
        {
            return Process.GetProcessById(pid).StartTime.ToUniversalTime();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>base64(UTF8(경로))에서 비영숫자 제거 후 앞 12자.</summary>
    public static string ProjectKey(string projectPath)
    {
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(projectPath.ToLowerInvariant()));
        var clean = new string(b64.Where(char.IsLetterOrDigit).ToArray());
        return clean.Length <= 12 ? clean : clean[..12];
    }
}
