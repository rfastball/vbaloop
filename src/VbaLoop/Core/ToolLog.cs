using System.Text.Json;

namespace VbaLoop.Core;

/// <summary>
/// 도구 로그 &lt;프로젝트루트&gt;\_dbg\vbaloop.jsonl — append-only.
/// 이번 실행에서 기록한 레코드를 메모리에도 유지해 실패 시 출력에 자동 인라인.
/// </summary>
public sealed class ToolLog
{
    private readonly string? _path;
    private readonly string _command;
    public List<string> Records { get; } = new();

    public ToolLog(string? path, string command)
    {
        _path = path;
        _command = command;
    }

    public void Write(string level, string @event, object? data = null)
    {
        var rec = new Dictionary<string, object?>
        {
            ["ts"] = DateTime.UtcNow.ToString("o"),
            ["command"] = _command,
            ["level"] = level,
            ["event"] = @event,
        };
        if (data != null) rec["data"] = data;

        string line = JsonSerializer.Serialize(rec, new JsonSerializerOptions
        {
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        });
        Records.Add(line);

        if (_path == null) return;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
            using var fs = new FileStream(_path, FileMode.Append, FileAccess.Write, FileShare.Read);
            var bytes = CanonicalText.Utf8NoBom.GetBytes(line + "\n");
            fs.Write(bytes, 0, bytes.Length);
        }
        catch (IOException)
        {
            // 로그 실패가 명령을 죽이면 안 됨 — 메모리 레코드는 유지됨
        }
    }
}
