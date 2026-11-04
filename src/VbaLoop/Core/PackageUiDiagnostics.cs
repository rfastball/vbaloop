using System.Diagnostics;
using VbaLoop.Excel;

namespace VbaLoop.Core;

/// <summary>COM 호출 없이 별도 타이머로 관찰한다. 창 클릭/닫기/프로세스 종료는 하지 않는다.</summary>
internal sealed class PackageUiDiagnostics : IDisposable
{
    private readonly PackageUiSession _session;
    private readonly UpdaterDiagnostics _log;
    private readonly Timer _timer;
    private int _observing;
    private bool _versionRecorded;
    private readonly string? _workbook;
    private DateTime _lastFileUse;
    private string? _lastDialogs;

    public PackageUiDiagnostics(PackageUiSession session, UpdaterDiagnostics log, string? workbook)
    {
        _session = session;
        _log = log;
        _workbook = workbook;
        _timer = new Timer(_ => Observe(), null, 0, 2000);
    }

    private void Observe()
    {
        if (Interlocked.Exchange(ref _observing, 1) != 0) return;
        try
        {
            var state = _session.DiagnosticState();
            if (state.Pid is not { } pid) return;
            var identity = ProcessIdentity.Identify(pid, state.Start);
            _log.Write("excel_observation", new { state.Phase, state.SaveState, pid, identity = identity.ToString() });
            if (identity != IdentityState.Match) return;
            using var process = Process.GetProcessById(pid);
            if (!_versionRecorded)
            {
                _versionRecorded = true;
                try { _log.Write("excel_version", new { version = process.MainModule?.FileVersionInfo.FileVersion }); }
                catch (Exception ex) { _log.Error("excel_version_unavailable", ex); }
            }
            try { _log.Write("excel_resources", new { cpuMs = process.TotalProcessorTime.TotalMilliseconds, memoryBytes = process.WorkingSet64 }); }
            catch (Exception ex) { _log.Error("excel_resources_unavailable", ex); }
            // 본문/입력란/컨트롤 텍스트/스크린샷은 수집하지 않는다. 제목은 문서명을 포함할 수 있다.
            var windows = NativeMethods.TopLevelWindowsOf(pid)
                .Where(w => w.Visible || w.ClassName is "XLMAIN" or "#32770" or "NUIDialog" or "ThunderDFrame")
                .Take(32).Select(w => new {
                handle = w.Hwnd.ToInt64(), windowClass = w.ClassName, visible = w.Visible,
                enabled = NativeMethods.IsWindowEnabled(w.Hwnd),
                title = w.ClassName is "#32770" or "NUIDialog" ? ModalHunter.GetCaption(w.Hwnd) : null,
            }).ToArray();
            _log.Write("excel_windows", windows);
            string dialogs = string.Join(",", windows.Where(w => w.windowClass is "#32770" or "NUIDialog").Select(w => w.handle));
            if (_workbook != null && state.Phase == "saving" &&
                (dialogs != _lastDialogs || DateTime.UtcNow - _lastFileUse > TimeSpan.FromSeconds(10)))
            {
                _lastDialogs = dialogs;
                _lastFileUse = DateTime.UtcNow;
                FileUseDiagnostics.Capture(_log, _workbook, dialogs.Length > 0 ? "save_dialog" : "saving", pid);
            }
        }
        catch (Exception ex) { _log.Error("observation_unavailable", ex); }
        finally { Volatile.Write(ref _observing, 0); }
    }

    public void Dispose() => _timer.Dispose();
}
