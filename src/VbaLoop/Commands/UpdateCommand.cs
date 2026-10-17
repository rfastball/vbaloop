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
            o.HumanLines.Add("No official release yet - nothing to update");
            o.Payload = new { mode = "check", current = installed, latest = (string?)null, updateAvailable = false, target };
            return o;
        }
        catch (Exception e)
        {
            o.Fail($"Could not fetch release info: {e.Message}",
                "Check the network and retry (or reinstall with install.ps1)", environment: true);
            return o;
        }
        if (info == null)
        {
            o.Fail("Cannot parse the latest release (tag is not semver)", $"Check https://github.com/{GitHubRelease.Repo}/releases");
            return o;
        }

        // 방금 알아낸 최신 버전으로 알림 캐시도 갱신(다음 status/doctor가 다시 조회하지 않게)
        try { UpdateCheck.Save(UpdateCheck.DefaultCachePath, new UpdateCheck.Cache { CheckedAt = DateTimeOffset.UtcNow.ToString("o"), Latest = info.Version }); }
        catch { /* 편의 캐시 */ }

        var latest = SemVer.ParseOrNull(info.Version)!;
        var baseline = SemVer.ParseOrNull(installed);
        bool newer = baseline == null || latest.CompareTo(baseline) > 0;

        o.HumanLines.Add($"Latest release: {info.Version} ({info.Tag})");
        o.HumanLines.Add($"Installed: {installed ?? "none"} ({target})");
        if (!runningInstalled)
            o.HumanLines.Add($"Running exe: {ExeVersion.Text} ({Environment.ProcessPath ?? "unknown"})");

        if (opts.Check || (!newer && !opts.Force))
        {
            o.HumanLines.Add(newer ? $"Update available: vbaloop update" : "already up to date");
            o.Payload = new { mode = "check", current = installed, latest = info.Version, updateAvailable = newer, target };
            return o;
        }

        string tmp = Path.Combine(Path.GetTempPath(), "vbaloop-update-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(tmp);
            if (!info.Assets.TryGetValue(GitHubRelease.ExeAsset, out var exeUrl))
                throw new CliError($"Release {info.Tag} has no {GitHubRelease.ExeAsset}");
            if (!info.Assets.TryGetValue(GitHubRelease.ChecksumsAsset, out var sumUrl))
                throw new CliError($"Release {info.Tag} has no {GitHubRelease.ChecksumsAsset} - an exe that cannot be verified is not installed");

            string tmpExe = Path.Combine(tmp, GitHubRelease.ExeAsset);
            string tmpSums = Path.Combine(tmp, GitHubRelease.ChecksumsAsset);
            Progress(opts, $"Downloading: {info.Tag} {GitHubRelease.ExeAsset}");
            using (var http = GitHubRelease.CreateClient(DownloadTimeout))
            {
                Download(http, exeUrl, tmpExe);
                Download(http, sumUrl, tmpSums);
            }
            var expected = GitHubRelease.ExpectedHash(File.ReadAllText(tmpSums), GitHubRelease.ExeAsset)
                ?? throw new CliError($"checksums.txt has no entry for {GitHubRelease.ExeAsset} - an exe that cannot be verified is not installed");
            var actual = GitHubRelease.Sha256Hex(tmpExe);
            if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                throw new CliError($"Checksum mismatch - download is corrupted or tampered with (expected {expected}, actual {actual})",
                    "Retry. If it repeats, do not install; report it");
            o.HumanLines.Add($"SHA256 verified: {actual}");

            SelfReplace.Replace(tmpExe, target);
            o.HumanLines.Add($"Installed: {target}" + (runningInstalled ? " (old exe is cleaned up on next run)" : ""));
            try
            {
                if (SelfInstall.EnsureOnPath()) o.HumanLines.Add($"PATH registered: {SelfInstall.InstallDir}");
            }
            catch (CliError e) { o.HumanLines.Add($"PATH registration failed (install completed): {e.Message}"); }
        }
        catch (CliError e)
        {
            o.Fail(e.Message, e.Hint ?? "retry vbaloop update", e.Environment);
            return o;
        }
        catch (Exception e) when (e is HttpRequestException or TaskCanceledException or IOException)
        {
            o.Fail($"Download failed: {e.Message}", "Check the network and retry vbaloop update", environment: true);
            return o;
        }
        finally
        {
            try { Directory.Delete(tmp, recursive: true); } catch { /* 임시 폴더 */ }
        }

        var installedNow = QueryExeVersion(target);
        if (installedNow != info.Version)
            o.HumanLines.Add($"Warning: installed exe reports {installedNow ?? "unknown"} (expected {info.Version})");

        // 관리 파일 수렴은 새 exe가 해야 한다 - 이 프로세스의 임베디드 정본은 옛 버전이다
        bool runInit = opts.Yes || (!opts.Json && ConsoleHost.Confirm(
            "Run init --all with the new exe to converge managed files of registered projects?", defaultYes: true, fallback: false));
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
                o.Fail($"init --all exited with {exit} (exe update completed)", "Check [FAIL]/[REFUSED] in the init --all output, then rerun vbaloop init --all");
        }
        else
        {
            o.HumanLines.Add("Converge managed files: vbaloop init --all (or vbaloop init in each project), then vbaloop push for projects whose LogShim changed");
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
        using var p = Process.Start(psi) ?? throw new CliError($"Failed to run {exe}");
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
