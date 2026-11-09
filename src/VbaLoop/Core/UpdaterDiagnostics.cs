using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VbaLoop.Core;

/// <summary>GUI/worker 공통 진단. 매 행을 닫아 강제 종료 뒤에도 회수할 수 있다.</summary>
public sealed class UpdaterDiagnostics
{
    public const string FolderVariable = "VBALOOP_DIAGNOSTIC_FOLDER";
    public const long MaxBytes = 8 * 1024 * 1024;
    public static string Root => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop", "diagnostics");
    private readonly object _gate = new();
    private readonly Stopwatch _clock = Stopwatch.StartNew();
    private readonly string _file;
    private readonly JsonSerializerOptions _json;
    private bool _stopped;

    public UpdaterDiagnostics(string folder, string name, params string?[] paths)
    {
        Directory.CreateDirectory(folder);
        _file = Path.Combine(folder, name + ".jsonl");
        var secrets = paths.Where(p => !string.IsNullOrWhiteSpace(p))
            .SelectMany(p => new[] { p!, Path.GetFileName(p!), Path.GetDirectoryName(p!) })
            .Concat([Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), Environment.UserName])
            .Where(p => !string.IsNullOrWhiteSpace(p)).Select(p => p!).Distinct().OrderByDescending(p => p.Length).ToArray();
        _json = new JsonSerializerOptions();
        _json.Converters.Add(new RedactedStrings(secrets));
        // 시작 전에 기록 가능 여부 확인. 작업 도중 로깅 실패는 업무 파일 처리에 전파하지 않는다.
        using (File.Open(_file, FileMode.CreateNew, FileAccess.Write, FileShare.ReadWrite)) { }
        Write("started", new { version = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
            pid = Environment.ProcessId, os = Environment.OSVersion.VersionString,
            runtime = Environment.Version.ToString(), process64 = Environment.Is64BitProcess,
            culture = System.Globalization.CultureInfo.CurrentCulture.Name });
    }

    public void Write(string kind, object? data = null)
    {
        lock (_gate)
        {
            if (_stopped) return;
            try
            {
                var safeData = JsonSerializer.SerializeToElement(data, _json);
                string line = JsonSerializer.Serialize(new { utc = DateTime.UtcNow, elapsedMs = _clock.ElapsedMilliseconds, kind, data = safeData });
                byte[] bytes = Encoding.UTF8.GetBytes(line + "\n");
                using var file = new FileStream(_file, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                if (file.Length + bytes.Length > MaxBytes - 1024)
                {
                    file.Write(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(new { utc = DateTime.UtcNow, kind = "log_limit_reached" }) + "\n"));
                    _stopped = true;
                    return;
                }
                file.Write(bytes);
                file.Flush();
            }
            catch { _stopped = true; /* 진단 장애로 적용/저장을 중단하지 않는다. */ }
        }
    }

    private sealed class RedactedStrings(string[] secrets) : JsonConverter<string>
    {
        public override string Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options) => throw new NotSupportedException();
        public override void Write(Utf8JsonWriter writer, string value, JsonSerializerOptions options)
        {
            foreach (string secret in secrets) value = value.Replace(secret, "[redacted]", StringComparison.OrdinalIgnoreCase);
            writer.WriteStringValue(value);
        }
    }

    // 저장 호출 예외만 본문 수집을 허용한다. 컴파일/실행 예외에는 VBA 코드/업무 데이터가 포함될 수 있다.
    public void Error(string kind, Exception error, bool includeMessage = false) => Write(kind, new {
        type = error.GetType().FullName, hresult = $"0x{error.HResult:X8}", stack = error.StackTrace,
        message = includeMessage ? error.Message[..Math.Min(error.Message.Length, 4096)] : null,
        innerType = error.InnerException?.GetType().FullName,
        innerHresult = error.InnerException is { } inner ? $"0x{inner.HResult:X8}" : null });

    public void FileState(string kind, string path)
    {
        try
        {
            var file = new FileInfo(path);
            Write(kind, new { extension = file.Extension, exists = file.Exists,
                length = file.Exists ? (long?)file.Length : null,
                modifiedUtc = file.Exists ? (DateTime?)file.LastWriteTimeUtc : null,
                attributes = file.Exists ? file.Attributes.ToString() : null });
        }
        catch (Exception ex) { Error(kind + "_unavailable", ex); }
    }
}
