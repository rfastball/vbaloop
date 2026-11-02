using System.Runtime.InteropServices;
using System.Text.Json;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// 치명적 COM 오류(연결 끊김 계열) 발생 시 메타파일에 poison 기록 후 환경 오류로 전환
/// (설계 §4 공통 계약). Excel이 아직 살아있으면 자동 kill하지 않고 격리 규칙(설계 §11)을
/// 적용한다 — 호출자가 CommandMutex를 쥔 상태라 마커 기록이 다음 명령의 게이트보다 먼저다.
/// Excel 프로세스 자체가 사라졌으면(VBE 폼 임포트 등에서의 크래시 — 실측 2026-11-02) poison이 아니라
/// 크래시로 보고하고, Office가 다시 띄운 Excel이 있으면 kill하지 않고 알린다.
/// </summary>
internal static class ComFailureGuard
{
    public static T Run<T>(ProjectConfig cfg, SessionMeta meta, ToolLog log, Func<T> body)
    {
        // Excel 프로세스 소멸 판정용 — 핸들을 작업 전에 쥐어야 소멸 뒤에도 종료 코드를 읽는다
        using var exitWatch = ExcelExitWatch.Open(meta.Pid, meta.ExcelStartTime);
        var kind = ExcelExitKind.Alive;
        try
        {
            return body();
        }
        catch (Exception ex) when (ex is COMException or InvalidComObjectException
                                   && (kind = exitWatch.ClassifyAfter(ex)) != ExcelExitKind.Alive)
        {
            throw Crashed(cfg, meta, log, exitWatch, kind, ex.Message);
        }
        catch (CliError ce) when ((kind = exitWatch.Classify(0)) != ExcelExitKind.Alive)
        {
            // 명령이 COM 실패를 자기 오류로 바꿔 던졌어도(원복 실패 등) 원인은 Excel 소멸이다
            throw Crashed(cfg, meta, log, exitWatch, kind, ce.Message);
        }
        catch (Exception ex) when (SessionService.PoisonIfFatal(cfg, meta, ex, log))
        {
            var (res, marker, why) = RecoveryGate.Resolve(cfg, meta, "poison", ex.Message, meta.LastCommand, log);
            throw new CliError(
                $"Session marked poisoned by fatal COM error: {ex.Message} — {RecoveryGate.Explain(res, marker, why)}",
                RecoveryGate.HintAfter(res, marker), environment: true);
        }
    }

    private static CliError Crashed(ProjectConfig cfg, SessionMeta meta, ToolLog log, ExcelExitWatch exitWatch,
        ExcelExitKind kind, string cause)
    {
        var info = exitWatch.Describe(kind, meta.WorkbookFullName);
        string command = meta.LastCommand ?? "the command";
        string? step = LastStep(log);
        log.Write(kind == ExcelExitKind.Crashed ? "error" : "warn",
            kind == ExcelExitKind.Crashed ? "excel_crashed" : "excel_exited", new
            {
                command,
                step,
                pid = info.Pid,
                exitCode = ExcelCrash.FormatExitCode(info.ExitCode),
                relaunched = info.Relaunched,
                workbookLocked = info.WorkbookLocked,
                workbook = meta.WorkbookFullName,
                cause,
            });
        RecoveryGate.Resolve(cfg, meta, kind == ExcelExitKind.Crashed ? "excel_crashed" : "excel_exited",
            cause, command, log); // 기록된 프로세스는 이미 없다 — Clear로 끝난다

        string what = step != null ? $"{command} (after {step})" : command;
        string message = (kind == ExcelExitKind.Crashed
                             ? ExcelCrash.Detail(what, info.Pid, info.ExitCode)
                             : ExcelCrash.ExitedDetail(what, info.Pid))
                         + $": {cause}"
                         + (info.Relaunched.Count > 0 ? " — " + ExcelCrash.RelaunchNote(info.Relaunched, meta.WorkbookFullName, info.WorkbookLocked) : "");
        string restart = meta.Ephemeral
            ? "rerun the same command (ephemeral Excel starts fresh for each command)"
            : "session start";
        string hint = info.Relaunched.Count > 0
            ? ExcelCrash.Hint(info.Relaunched, restart)
            : $"Excel itself crashed during {command} — retry once ({restart}); "
              + "if it crashes again at the same step, report it with vbaloop bugreport";
        return new CliError(message, hint, environment: true);
    }

    /// <summary>도구 로그의 마지막 단계(예: form_imported frmMain) — 크래시 직전에 무엇을 하던 중인지.</summary>
    private static string? LastStep(ToolLog log)
    {
        for (int i = log.Records.Count - 1; i >= 0; i--)
        {
            try
            {
                using var doc = JsonDocument.Parse(log.Records[i]);
                var root = doc.RootElement;
                if (!root.TryGetProperty("event", out var ev) || ev.GetString() is not { } name) continue;
                if (name is "recovery_decision" or "session_poisoned") continue;
                string? module = root.TryGetProperty("data", out var d) && d.ValueKind == JsonValueKind.Object
                                 && d.TryGetProperty("module", out var m) && m.ValueKind == JsonValueKind.String
                    ? m.GetString() : null;
                return module != null ? $"{name} {module}" : name;
            }
            catch (JsonException)
            {
                // 깨진 줄은 건너뛴다
            }
        }
        return null;
    }
}
