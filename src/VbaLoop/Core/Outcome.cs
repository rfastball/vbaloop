using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>
/// 명령 실행 결과. 출력 계약: 성공 0, 환경 오류(세션/Excel/COM) 3, 그 외 2.
/// --json 봉투: {command, ok, exitCode, hint, errors[], payload}
/// </summary>
public sealed class Outcome
{
    public required string Command { get; init; }
    public bool Ok { get; set; } = true;
    public string? Hint { get; set; }
    public List<string> Errors { get; } = new();
    public object? Payload { get; set; }

    /// <summary>
    /// auto-ensure가 이 명령 진입에서 세션을 새로 기동했을 때만 설정되는 봉투 레벨 표식
    /// (예: {action:"started", pid, workbook}). Active 채택 평시엔 null — 출력 안정성 유지.
    /// </summary>
    public object? Session { get; set; }

    /// <summary>사람용 출력 라인(성공 경로). --json 모드에서는 무시된다.</summary>
    public List<string> HumanLines { get; } = new();

    /// <summary>실패 시 자동 인라인할 도구 로그 레코드. ToolLog.Records를 참조 공유.</summary>
    public List<string> LogRecords { get; set; } = new();

    /// <summary>환경 오류 분류(CliError.Environment 대응) — true면 exit 3.</summary>
    public bool EnvironmentError { get; set; }

    public int ExitCode => Ok ? 0 : EnvironmentError ? 3 : 2;

    public void Fail(string error, string? hint = null, bool environment = false)
    {
        Ok = false;
        Errors.Add(error);
        if (hint != null) Hint = hint;
        if (environment) EnvironmentError = true;
    }
}

public static class Output
{
    public static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static void Print(Outcome o, bool json)
    {
        if (json)
        {
            var envelope = new
            {
                command = o.Command,
                ok = o.Ok,
                exitCode = o.ExitCode,
                hint = o.Hint,
                session = o.Session,
                errors = o.Errors,
                payload = o.Payload,
            };
            Console.WriteLine(JsonSerializer.Serialize(envelope, JsonOptions));
            return;
        }

        foreach (var line in o.HumanLines)
            Console.WriteLine(line);

        if (!o.Ok)
        {
            var first = o.Errors.Count > 0 ? o.Errors[0] : "unknown error";
            var hint = o.Hint != null ? $" hint=[{o.Hint}]" : "";
            Console.Error.WriteLine($"vbaloop {o.Command}: error: {first}{hint}");
            foreach (var extra in o.Errors.Skip(1))
                Console.Error.WriteLine($"  - {extra}");
            foreach (var rec in o.LogRecords)
                Console.Error.WriteLine($"  log: {rec}");
        }
    }
}
