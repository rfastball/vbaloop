using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// "이 Excel을 kill해도 사용자 작업을 잃지 않는가"를 판정하는 프로브(설계 §11).
/// 행 걸린 Excel에 대한 AccessibleObjectFromWindow/COM 호출은 무기한 블로킹될 수 있고, 블로킹된
/// STA RCW는 종료 시 파이널라이저까지 붙잡는다 — 그래서 test 러너와 같은 방식으로 자식 프로세스
/// (<c>vbaloop __probe &lt;pid&gt; &lt;startTime&gt;</c>)에서 돌리고 부모가 벽시계 제한으로 감독한다.
/// 시간 초과 = Unresponsive(→ 격리). 자식에는 메시지 필터를 등록하지 않는다 — busy 거절이
/// 재시도 없이 즉시 RPC_E_CALL_REJECTED로 돌아와 Busy로 분류된다.
/// </summary>
public static class SessionProbe
{
    public const int DefaultTimeoutMs = 8000;

    private static readonly JsonSerializerOptions ProbeJson = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    /// <summary>부모 측: 자식 프로브를 띄우고 timeoutMs 안에 결과를 받는다. 예외를 던지지 않는다.</summary>
    public static ProbeResult Run(int pid, string startTimeIso, int timeoutMs = DefaultTimeoutMs)
    {
        string? exe = Environment.ProcessPath;
        if (exe == null) return ProbeResult.Unresponsive("cannot determine own exe path");

        var psi = new ProcessStartInfo
        {
            FileName = exe,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            CreateNoWindow = true,
        };
        psi.ArgumentList.Add("__probe");
        psi.ArgumentList.Add(pid.ToString());
        psi.ArgumentList.Add(string.IsNullOrEmpty(startTimeIso) ? "none" : startTimeIso);

        try
        {
            using var proc = Process.Start(psi);
            if (proc == null) return ProbeResult.Unresponsive("probe process failed to start");
            var stdoutTask = proc.StandardOutput.ReadToEndAsync();
            _ = proc.StandardError.ReadToEndAsync();
            if (!proc.WaitForExit(timeoutMs))
            {
                // 프로브 자식은 우리 프로세스 — 이것을 죽이는 것은 Excel에 영향 없음
                try { proc.Kill(entireProcessTree: true); } catch { /* 이미 종료 */ }
                return ProbeResult.Unresponsive($"probe timed out after {timeoutMs}ms");
            }
            return Parse(stdoutTask.GetAwaiter().GetResult());
        }
        catch (Exception ex)
        {
            return ProbeResult.Unresponsive($"probe run failed: {ex.Message}");
        }
    }

    public static ProbeResult Parse(string stdout)
    {
        int brace = stdout.IndexOf('{');
        if (brace < 0) return ProbeResult.Unresponsive("probe emitted no JSON");
        try
        {
            using var doc = JsonDocument.Parse(stdout[brace..]);
            var root = doc.RootElement;
            var kind = Enum.TryParse<ProbeKind>(root.GetProperty("kind").GetString(), true, out var k)
                ? k : ProbeKind.Unresponsive;
            var wbs = root.TryGetProperty("workbooks", out var w) && w.ValueKind == JsonValueKind.Array
                ? w.EnumerateArray().Select(e => e.GetString() ?? "").ToList()
                : new List<string>();
            string? detail = root.TryGetProperty("detail", out var d) && d.ValueKind == JsonValueKind.String
                ? d.GetString() : null;
            return new ProbeResult(kind, wbs, detail);
        }
        catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException)
        {
            return ProbeResult.Unresponsive($"probe output parse failed: {ex.Message}");
        }
    }

    /// <summary>자식 측(<c>__probe</c>): stdout에 JSON 한 덩어리, 항상 exit 0.</summary>
    public static int RunChild(CliArgs opts)
    {
        ProbeResult r;
        try
        {
            if (opts.Positional.Count < 2 || !int.TryParse(opts.Positional[0], out int pid))
                r = ProbeResult.Unresponsive("__probe argument error: <pid> <startTime>");
            else
                r = ProbeInProcess(pid, opts.Positional[1] == "none" ? "" : opts.Positional[1]);
        }
        catch (Exception ex)
        {
            r = ProbeResult.Unresponsive($"{ex.GetType().Name}: {ex.Message}");
        }
        Console.WriteLine(JsonSerializer.Serialize(
            new { kind = r.Kind.ToString(), workbooks = r.Workbooks, detail = r.Detail }, ProbeJson));
        return 0;
    }

    private static ProbeResult ProbeInProcess(int pid, string startTimeIso)
    {
        if (ProcessIdentity.Identify(pid, startTimeIso) != IdentityState.Match)
            return ProbeResult.Unresponsive("process identity mismatch or unverifiable");

        var tops = NativeMethods.TopLevelWindowsOf(pid);
        var excel7 = new List<IntPtr>();
        foreach (var (hwnd, _, _) in tops)
            excel7.AddRange(Excel7Descendants(hwnd));

        // 보이는 창이 하나도 없고 워크북 창(EXCEL7)도 없으면 사람이 작업 중일 수 있는 표면이 없다
        if (excel7.Count == 0 && !tops.Any(t => t.Visible))
            return new ProbeResult(ProbeKind.Windowless, Array.Empty<string>(), $"{tops.Count} top-level windows, all hidden");
        if (excel7.Count == 0)
            return ProbeResult.Unresponsive("visible windows exist but no workbook window (EXCEL7); cannot check via COM");

        string last = "";
        bool sawBusy = false;
        foreach (var h in excel7)
        {
            var att = ExcelAttach.ApplicationFromExcel7(h, pid, retries: 0);
            if (att.App == null)
            {
                sawBusy |= att.Busy;
                last = att.FailReason ?? "";
                continue;
            }
            try
            {
                dynamic app = att.App;
                var names = new List<string>();
                int count = Convert.ToInt32(app.Workbooks.Count);
                for (int i = 1; i <= count; i++)
                    names.Add(Convert.ToString(app.Workbooks.Item(i).FullName) ?? "?");
                return new ProbeResult(ProbeKind.Responsive, names);
            }
            catch (COMException ex) when (ComRetry.IsBusy(ex.HResult))
            {
                sawBusy = true;
                last = ex.Message;
            }
            catch (Exception ex)
            {
                last = ex.Message;
            }
        }
        return sawBusy
            ? new ProbeResult(ProbeKind.Busy, Array.Empty<string>(), last)
            : ProbeResult.Unresponsive(last);
    }

    private static List<IntPtr> Excel7Descendants(IntPtr parent)
    {
        var found = new List<IntPtr>();
        NativeMethods.EnumChildWindows(parent, (h, _) =>
        {
            if (NativeMethods.GetWindowClassName(h) == "EXCEL7") found.Add(h);
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
