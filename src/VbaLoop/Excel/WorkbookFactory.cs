using System.Diagnostics;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// 빈 <c>.xlsm</c> 워크북 생성(setup 마법사 전용). dist와 같은 격리 패턴 — 자체 임시 인스턴스에서
/// 만들고 즉시 닫은 뒤 kill로 회수하므로 도구 세션과 무관하다. "매크로 사용 통합 문서로 저장"을
/// 사람이 고르다 실수하는 지점(기본 저장이 .xlsx라 VBA가 통째로 날아간다)을 원천 차단한다.
/// </summary>
public static class WorkbookFactory
{
    /// <summary>
    /// path에 빈 매크로 워크북을 만든다. 이미 존재하면 CliError — 마법사는 빈 폴더에서만 부른다.
    /// </summary>
    public static void CreateMacroWorkbook(string path)
    {
        if (File.Exists(path))
            throw new CliError($"이미 존재하는 파일: {path}", "다른 폴더를 고르거나 기존 워크북을 사용");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        var filter = OleMessageFilter.Register();
        ExcelLauncher.LaunchResult? launch = null;
        try
        {
            var l = ExcelLauncher.CreateNew(visible: false);
            launch = l;
            dynamic app = l.App;
            try
            {
                app.DisplayAlerts = false;
                app.EnableEvents = false;
                dynamic wb = ComRetry.Invoke<dynamic>(() => app.Workbooks.Add());
                ComRetry.Invoke(() => { wb.SaveAs(path, DistFormat.XlOpenXMLWorkbookMacroEnabled); });
                ComRetry.Invoke(() => { wb.Close(false); });
                try { app.Quit(); } catch { /* finally에서 kill 보증 */ }
            }
            catch
            {
                try { app.Quit(); } catch { /* finally에서 kill 보증 */ }
                throw;
            }
        }
        finally
        {
            filter.Revoke();
            if (launch != null && ProcessIdentity.IsAlive(launch.Pid, launch.StartTimeIso))
            {
                try { Process.GetProcessById(launch.Pid).Kill(entireProcessTree: true); }
                catch { /* 회수 실패는 doctor의 고아 스캔이 잡는다 */ }
            }
        }
    }
}
