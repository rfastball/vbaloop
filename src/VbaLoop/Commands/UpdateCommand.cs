using System.Diagnostics;
using System.Text;
using System.Text.Json;
using VbaLoop.Cli;
using VbaLoop.Core;

namespace VbaLoop.Commands;

/// <summary>
/// update(설계 §12.5): 최신 릴리스를 받아 SHA-256을 확인하고(install.ps1과 같은 fail-closed 계약)
/// 설치본(%LOCALAPPDATA%\vbaloop\bin\vbaloop.exe)을 교체한다. 실행 중인 exe는 .old로 이름을 바꿔 두고
/// 다음 기동이 지운다. 그 뒤 새 exe로 init --all을 돌릴지 묻는다(--yes면 묻지 않고 실행, 비대화형/--json에서
/// --yes가 없으면 실행하지 않고 안내만). --check는 확인만, --force는 같은/낮은 버전도 다시 설치.
/// </summary>
public static class UpdateCommand
{
    private static readonly TimeSpan ApiTimeout = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan DownloadTimeout = TimeSpan.FromMinutes(10);

    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "update" };
        string target = SelfInstall.InstalledExePath;
        bool runningInstalled = Environment.ProcessPath != null && PathDrift.SamePath(Environment.ProcessPath, target);
        string? installed = runningInstalled ? ExeVersion.Text
            : File.Exists(target) ? QueryExeVersion(target) : null;

        GitHubRelease.Info? info;
        try
        {
            using var http = GitHubRelease.CreateClient(ApiTimeout);
            info = GitHubRelease.FetchLatestAsync(http).GetAwaiter().GetResult();
        }
        catch (HttpRequestException e) when (e.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            // releases/latest는 정식 릴리스만 본다 — prerelease만 있으면 404
            o.HumanLines.Add("아직 정식 릴리스가 없음 - 갱신할 것이 없음");
            o.Payload = new { mode = "check", current = installed, latest = (string?)null, updateAvailable = false, target };
            return o;
        }
        catch (Exception e)
        {
            o.Fail($"릴리스 정보를 가져오지 못함: {e.Message}",
                "네트워크 확인 후 재시도 (또는 install.ps1로 재설치)", environment: true);
            return o;
        }
        if (info == null)
        {
            o.Fail("최신 릴리스를 판독할 수 없음(semver 태그 아님)", $"https://github.com/{GitHubRelease.Repo}/releases 확인");
            return o;
        }

        // 방금 알아낸 최신 버전으로 알림 캐시도 갱신(다음 status/doctor가 다시 조회하지 않게)
        try { UpdateCheck.Save(UpdateCheck.DefaultCachePath, new UpdateCheck.Cache { CheckedAt = DateTimeOffset.UtcNow.ToString("o"), Latest = info.Version }); }
        catch { /* 편의 캐시 */ }

        var latest = SemVer.ParseOrNull(info.Version)!;
        var baseline = SemVer.ParseOrNull(installed);
        bool newer = baseline == null || latest.CompareTo(baseline) > 0;

        o.HumanLines.Add($"최신 릴리스: {info.Version} ({info.Tag})");
        o.HumanLines.Add($"설치본: {installed ?? "없음"} ({target})");
        if (!runningInstalled)
            o.HumanLines.Add($"실행 exe: {ExeVersion.Text} ({Environment.ProcessPath ?? "미상"})");

        if (opts.Check || (!newer && !opts.Force))
        {
            o.HumanLines.Add(newer ? $"갱신 가능: vbaloop update" : "이미 최신");
            o.Payload = new { mode = "check", current = installed, latest = info.Version, updateAvailable = newer, target };
            return o;
        }

        string tmp = Path.Combine(Path.GetTempPath(), "vbaloop-update-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(tmp);
            if (!info.Assets.TryGetValue(GitHubRelease.ExeAsset, out var exeUrl))
                throw new CliError($"릴리스 {info.Tag}에 {GitHubRelease.ExeAsset}가 없음");
            if (!info.Assets.TryGetValue(GitHubRelease.ChecksumsAsset, out var sumUrl))
                throw new CliError($"릴리스 {info.Tag}에 {GitHubRelease.ChecksumsAsset}가 없음 - 검증할 수 없는 exe는 설치하지 않음");

            string tmpExe = Path.Combine(tmp, GitHubRelease.ExeAsset);
            string tmpSums = Path.Combine(tmp, GitHubRelease.ChecksumsAsset);
            Progress(opts, $"다운로드: {info.Tag} {GitHubRelease.ExeAsset}");
            using (var http = GitHubRelease.CreateClient(DownloadTimeout))
            {
                Download(http, exeUrl, tmpExe);
                Download(http, sumUrl, tmpSums);
            }
            var expected = GitHubRelease.ExpectedHash(File.ReadAllText(tmpSums), GitHubRelease.ExeAsset)
                ?? throw new CliError($"checksums.txt에 {GitHubRelease.ExeAsset} 항목이 없음 - 검증할 수 없는 exe는 설치하지 않음");
            var actual = GitHubRelease.Sha256Hex(tmpExe);
            if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                throw new CliError($"체크섬 불일치 - 다운로드가 손상되었거나 변조됨 (expected {expected}, actual {actual})",
                    "재시도. 반복되면 설치하지 말고 리포트");
            o.HumanLines.Add($"SHA256 확인: {actual}");

            SelfReplace.Replace(tmpExe, target);
            o.HumanLines.Add($"설치: {target}" + (runningInstalled ? " (이전 exe는 다음 실행 때 정리)" : ""));
            try
            {
                if (SelfInstall.EnsureOnPath()) o.HumanLines.Add($"PATH 등록: {SelfInstall.InstallDir}");
            }
            catch (CliError e) { o.HumanLines.Add($"PATH 등록 실패(설치는 완료): {e.Message}"); }
        }
        catch (CliError e)
        {
            o.Fail(e.Message, e.Hint ?? "vbaloop update 재시도", e.Environment);
            return o;
        }
        catch (Exception e) when (e is HttpRequestException or TaskCanceledException or IOException)
        {
            o.Fail($"다운로드 실패: {e.Message}", "네트워크 확인 후 vbaloop update 재시도", environment: true);
            return o;
        }
        finally
        {
            try { Directory.Delete(tmp, recursive: true); } catch { /* 임시 폴더 */ }
        }

        var installedNow = QueryExeVersion(target);
        if (installedNow != info.Version)
            o.HumanLines.Add($"주의: 설치된 exe가 {installedNow ?? "미상"}을 보고함 (기대 {info.Version})");

        // 관리 파일 수렴은 새 exe가 해야 한다 - 이 프로세스의 임베디드 정본은 옛 버전이다
        bool runInit = opts.Yes || (!opts.Json && ConsoleHost.Confirm(
            "새 exe로 init --all을 실행해 등록된 프로젝트의 관리 파일을 수렴할까요?", defaultYes: true, fallback: false));
        object? initAll = null;
        if (runInit)
        {
            var (exit, stdout, stderr) = RunExe(target, opts.Json ? new[] { "init", "--all", "--json" } : new[] { "init", "--all" });
            if (opts.Json)
            {
                try { initAll = JsonSerializer.Deserialize<JsonElement>(stdout); }
                catch (JsonException) { initAll = new { exitCode = exit, output = stdout }; }
            }
            else
            {
                o.HumanLines.Add("== init --all ==");
                o.HumanLines.AddRange(SplitLines(stdout));
                o.HumanLines.AddRange(SplitLines(stderr));
            }
            if (exit != 0)
                o.Fail($"init --all이 exit {exit}로 끝남 (exe 갱신은 완료)", "init --all 출력의 [FAIL]/[REFUSED] 확인 후 vbaloop init --all 재실행");
        }
        else
        {
            o.HumanLines.Add("관리 파일 수렴: vbaloop init --all (또는 각 프로젝트에서 vbaloop init) 후 LogShim이 바뀐 프로젝트는 vbaloop push");
        }

        o.Payload = new
        {
            mode = "update",
            previous = installed,
            current = installedNow ?? info.Version,
            latest = info.Version,
            target,
            initAll,
            next = runInit ? null : "vbaloop init --all",
        };
        return o;
    }

    private static void Progress(CliArgs opts, string line)
    {
        if (!opts.Json) Console.Error.WriteLine(line);
    }

    private static void Download(HttpClient http, string url, string dest)
    {
        using var resp = http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
        resp.EnsureSuccessStatusCode();
        using var src = resp.Content.ReadAsStream();
        using var fs = File.Create(dest);
        src.CopyTo(fs);
    }

    /// <summary>다른 exe의 버전(`version --json`의 payload.version). 실패하면 null.</summary>
    private static string? QueryExeVersion(string exe)
    {
        try
        {
            var (exit, stdout, _) = RunExe(exe, new[] { "version", "--json" }, TimeSpan.FromSeconds(20));
            if (exit != 0) return null;
            using var doc = JsonDocument.Parse(stdout);
            return doc.RootElement.GetProperty("payload").GetProperty("version").GetString();
        }
        catch
        {
            return null;
        }
    }

    private static (int Exit, string Stdout, string Stderr) RunExe(string exe, string[] args, TimeSpan? timeout = null)
    {
        var psi = new ProcessStartInfo(exe)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            CreateNoWindow = true,
        };
        foreach (var a in args) psi.ArgumentList.Add(a);
        // 자식이 다시 업데이트 확인 자식을 띄우지 않게(이 명령이 방금 캐시를 썼다)
        psi.Environment[UpdateCheck.OptOutEnv] = "1";
        using var p = Process.Start(psi) ?? throw new CliError($"{exe} 실행 실패");
        var outTask = p.StandardOutput.ReadToEndAsync();
        var errTask = p.StandardError.ReadToEndAsync();
        if (!p.WaitForExit(timeout ?? Timeout.InfiniteTimeSpan))
        {
            try { p.Kill(entireProcessTree: true); } catch { /* 이미 종료 */ }
            return (-1, "", "timeout");
        }
        return (p.ExitCode, outTask.GetAwaiter().GetResult(), errTask.GetAwaiter().GetResult());
    }

    private static IEnumerable<string> SplitLines(string s)
        => s.Replace("\r\n", "\n").Split('\n').Where(l => l.Length > 0);
}
