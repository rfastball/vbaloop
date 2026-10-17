using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// 세션 필요 명령(push/pull/test/run) 진입부의 auto-ensure 공용부: 세션을 확보하고,
/// 새로 기동했을 때만 봉투 레벨 session 표식과 사람용 라인을 남긴다(관찰 가능성).
/// mutex+lease 획득 후에 호출할 것 — 다중 에이전트 동시 auto-start 레이스 방지.
/// </summary>
internal static class SessionEnsure
{
    public static EnsureResult ForCommand(ProjectConfig cfg, ToolLog log, Outcome o, string command)
    {
        var r = SessionService.EnsureActive(cfg, log, command);
        if (r.Action == "started")
        {
            o.Session = new { action = "started", pid = r.Meta.Pid, workbook = r.Meta.WorkbookFullName };
            o.HumanLines.Add($"세션 자동 시작 (PID {r.Meta.Pid})");
        }
        return r;
    }
}
