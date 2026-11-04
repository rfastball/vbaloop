using System.Security.Cryptography;
using System.Text.Json;
using System.Diagnostics;

namespace VbaLoop.Core;

/// <summary>GUI 전용 프로토콜. 일반 CLI 출력과 저장 정책은 바꾸지 않는다.</summary>
public sealed class PackageUiSession(string? expectedPackage, string? expectedWorkbook, UpdaterDiagnostics? diagnostics = null)
{
    public UpdaterDiagnostics? Diagnostics { get; } = diagnostics;
    private readonly object _gate = new();
    private DateTime _lastProgress = DateTime.UtcNow;
    private bool _finished, _slowSaveReported;
    private int? _excelPid;
    private string? _excelStart;
    public string Phase { get; private set; } = "checking";
    public string SaveState { get; private set; } = "not_saved";
    public string? PackageHash { get; private set; }
    public string? WorkbookHash { get; private set; }

    public static string FileHash(string path, bool openedInExcel = false)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
            openedInExcel ? FileShare.ReadWrite : FileShare.Read);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    public void Verify(byte[] package, string workbook)
    {
        PackageHash = Convert.ToHexString(SHA256.HashData(package));
        WorkbookHash = FileHash(workbook);
        if ((expectedPackage != null && !PackageHash.Equals(expectedPackage, StringComparison.OrdinalIgnoreCase)) ||
            (expectedWorkbook != null && !WorkbookHash.Equals(expectedWorkbook, StringComparison.OrdinalIgnoreCase)))
            throw new CliError("Files changed after inspection; inspect again before applying", code: "files_changed");
    }

    public void VerifyOpenedWorkbook(string path)
    {
        if (!string.Equals(WorkbookHash, FileHash(path, openedInExcel: true), StringComparison.OrdinalIgnoreCase))
            throw new CliError("Workbook changed while opening; inspect again", code: "files_changed");
    }

    public void Report(string phase)
    {
        lock (_gate)
        {
            if (_finished) throw new CliError("Operation already stopped", code: "operation_timeout");
            Phase = phase;
            _lastProgress = DateTime.UtcNow;
            if (phase == "saving") SaveState = "unknown";
            if (phase == "saved") SaveState = "saved";
            Diagnostics?.Write("phase", new { phase, saveState = SaveState });
            WriteProgress(phase);
        }
    }

    public void OwnExcel(int pid, string started)
    {
        lock (_gate) { _excelPid = pid; _excelStart = started; }
        Diagnostics?.Write("excel_owned", new { pid, started });
    }

    internal (string Phase, string SaveState, int? Pid, string? Start) DiagnosticState()
    {
        lock (_gate) return (Phase, SaveState, _excelPid, _excelStart);
    }

    public IDisposable Watch(Action onTimeout)
        => new Timer(_ => CheckTimeout(DateTime.UtcNow, onTimeout), null, 1000, 1000);

    public void CheckTimeout(DateTime now, Action onTimeout)
    {
        lock (_gate)
        {
            if (_finished || now - _lastProgress < TimeSpan.FromMinutes(2)) return;
            // Save 진입과 타임아웃의 소유권을 같은 잠금으로 결정한다. 저장 중 강제 종료 금지.
            if (Phase == "saving")
            {
                if (!_slowSaveReported) { _slowSaveReported = true; Diagnostics?.Write("slow_saving"); WriteProgress("slow_saving"); }
                return;
            }
            _finished = true;
            Diagnostics?.Write("timeout", new { phase = Phase, saveState = SaveState });
            if (_excelPid is { } pid && _excelStart is { } start && ProcessIdentity.Identify(pid, start) == IdentityState.Match)
            {
                try { Diagnostics?.Write("timeout_kill", new { pid }); Process.GetProcessById(pid).Kill(); } catch { /* 신원 확인한 인스턴스만 종료 시도 */ }
            }
            onTimeout();
        }
    }

    public void Finish(Action writeResult)
    {
        lock (_gate)
        {
            if (_finished) return;
            _finished = true;
            writeResult();
        }
    }

    private static void WriteProgress(string phase)
    {
        try { Console.Error.WriteLine(JsonSerializer.Serialize(new { protocol = 1, phase })); }
        catch (IOException) { /* 저장 지연 중 사용자가 GUI를 닫아도 작업은 계속한다. */ }
    }

    public static bool CanSave(DistCompileVerdict.Kind kind)
        => kind is DistCompileVerdict.Kind.Compiled or DistCompileVerdict.Kind.AlreadyCompiled;
}
