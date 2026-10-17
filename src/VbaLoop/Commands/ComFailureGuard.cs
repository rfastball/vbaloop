using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// 치명적 COM 오류(연결 끊김 계열) 발생 시 메타파일에 poison 기록 후 환경 오류로 전환
/// (설계 §4 공통 계약). Excel이 아직 살아있으면 자동 kill하지 않고 격리 규칙(설계 §11)을
/// 적용한다 — 호출자가 CommandMutex를 쥔 상태라 마커 기록이 다음 명령의 게이트보다 먼저다.
/// </summary>
internal static class ComFailureGuard
{
    public static T Run<T>(ProjectConfig cfg, SessionMeta meta, ToolLog log, Func<T> body)
    {
        try
        {
            return body();
        }
        catch (Exception ex) when (SessionService.PoisonIfFatal(cfg, meta, ex, log))
        {
            var (res, marker, why) = RecoveryGate.Resolve(cfg, meta, "poison", ex.Message, meta.LastCommand, log);
            throw new CliError(
                $"Session marked poisoned by fatal COM error: {ex.Message} — {RecoveryGate.Explain(res, marker, why)}",
                RecoveryGate.HintAfter(res, marker), environment: true);
        }
    }
}
