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

    /// <summary>계보 탐색을 멈추는 데스크톱/서비스 경계 프로세스.</summary>
    public static readonly string[] BoundaryNames =
        { "explorer", "services", "wininit", "winlogon", "svchost", "lsass", "csrss", "smss", "system", "idle" };

    public static readonly string[] ShellNames =
        { "cmd", "powershell", "pwsh", "conhost", "windowsterminal", "openconsole", "bash", "sh", "wsl", "wslhost" };

    public static readonly string[] EnvIdNames =
        { "VBALOOP_SESSION_ID", "CODEX_THREAD_ID", "CODEX_SESSION_ID", "CLAUDE_SESSION_ID", "AGENT_THREAD_ID" };

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
        foreach (var name in EnvIdNames)
        {
            var val = getEnv(name);
            if (!string.IsNullOrWhiteSpace(val))
                return new Identity($"env:{name}:{val.Trim()}", "env", null, null);
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
