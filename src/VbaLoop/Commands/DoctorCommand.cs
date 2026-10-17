using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>환경 점검(설계 §4 doctor). 세션 무관, 언제나 실행 가능. 애드인 체크 없음.</summary>
public static class DoctorCommand
{
    private sealed record Check(string Name, bool Ok, string Detail, string? Hint = null);

    /// <summary>
    /// PATH 문자열(`;` 구분)에 dir가 포함되는지 — 트레일링 <c>\</c> 차이 관용, 대소문자 무시,
    /// 빈/null PATH 관용. rollout이 등록하는 shim 폴더 존재 여부 판정용(정보성 doctor 체크).
    /// </summary>
    public static bool IsOnPath(string? pathVar, string dir)
    {
        if (string.IsNullOrEmpty(pathVar)) return false;
        string target = dir.TrimEnd('\\');
        return pathVar.Split(';')
            .Any(p => string.Equals(p.Trim().TrimEnd('\\'), target, StringComparison.OrdinalIgnoreCase));
    }

    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "doctor" };
        var checks = new List<Check>();

        // 0. 실행 exe 신선도 (정보성, 항상 Ok) — stale exe로 인한 "고친 버그 재현" 혼선 차단
        //    소스 레포에서는 rollout.ps1이 재빌드·전파한다.
        checks.Add(new Check("tool.build", true, VersionCommand.GetBuildInfo()));

        // 1. Excel 설치 / COM 활성화 가능 여부 (ProgID 해석)
        var progType = Type.GetTypeFromProgID("Excel.Application");
        string? curVer = Registry.ClassesRoot.OpenSubKey(@"Excel.Application\CurVer")
            ?.GetValue(null) as string;
        checks.Add(new Check("excel.com", progType != null,
            progType != null ? $"Excel.Application 등록됨 ({curVer ?? "버전 미상"})" : "Excel.Application ProgID 없음",
            progType == null ? "Excel 설치 필요" : null));

        // 2. VBIDE 신뢰 액세스 (AccessVBOM). --fix는 명시가 곧 승인(--kill-orphans와 같은 원칙) —
        //    보안 설정이라 기본 doctor는 절대 건드리지 않고 상태만 보고한다.
        var (vbomOk, vbomDetail, _) = AccessVbom.Detect();
        if (!vbomOk && opts.Fix)
        {
            try
            {
                var (changed, fixDetail) = AccessVbom.Enable(curVer);
                if (changed)
                {
                    (vbomOk, vbomDetail, _) = AccessVbom.Detect();
                    vbomDetail += $" (--fix: {fixDetail})";
                }
            }
            catch (CliError e)
            {
                vbomDetail += $" (--fix 실패: {e.Message})";
            }
        }
        checks.Add(new Check("excel.accessVBOM", vbomOk, vbomDetail,
            vbomOk ? null
            : "vbaloop doctor --fix로 켜거나, Excel 옵션 > 보안 센터 > 매크로 설정 > 'VBA 프로젝트 개체 모델에 안전하게 액세스' 체크"));

        // 3. oleacc.dll 기반 attach 가능 여부
        bool oleacc = NativeLibrary.TryLoad("oleacc.dll", out var oleaccHandle);
        if (oleacc) NativeLibrary.Free(oleaccHandle);
        checks.Add(new Check("attach.oleacc", oleacc,
            oleacc ? "oleacc.dll 로드 가능" : "oleacc.dll 로드 실패",
            oleacc ? null : "Windows 구성 손상 — oleacc.dll 확인"));

        // 4. .NET 런타임 / 비트니스
        string arch = RuntimeInformation.ProcessArchitecture.ToString();
        checks.Add(new Check("runtime.dotnet", true,
            $"{RuntimeInformation.FrameworkDescription} ({arch})"));

        string? excelBitness = ReadExcelBitness();
        bool bitnessKnown = excelBitness != null;
        checks.Add(new Check("excel.bitness", true,
            bitnessKnown
                ? $"Excel {excelBitness} / 도구 {arch} (out-of-proc COM이라 불일치도 동작하나 참고용)"
                : "Excel 비트니스 레지스트리에서 판별 불가(참고용)"));

        // 5. 시스템 ACP (코드페이지 하드코딩 금지 원칙의 실측값)
        int acp = NativeMethods.GetACP();
        checks.Add(new Check("system.acp", true, $"ACP={acp}" + (acp == 949 ? " (CP949 한국어)" : "")));

        // 5b. PATH shim 등록 상태. 정보성 — 항상 Ok(게이트/errors 비오염),
        //     상태는 Detail로만 안내. 등록은 rollout.ps1이 %LOCALAPPDATA%\vbaloop\bin에 수행.
        string shimDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop", "bin");
        string? userPath = Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.User);
        bool shimOnPath = IsOnPath(userPath, shimDir);
        bool shimExe = File.Exists(Path.Combine(shimDir, "vbaloop.exe"));
        // 안내는 exe가 소스 레포 아래인지로 갈린다 — 설치본 사용자에게 rollout.ps1은 존재하지 않는다.
        string exeDirForRepo = Environment.ProcessPath != null
            ? Path.GetDirectoryName(Environment.ProcessPath)! : Environment.CurrentDirectory;
        string reinstallHint = RepoLocator.FindRepoRoot(exeDirForRepo) != null
            ? "rollout.ps1 실행" : "vbaloop setup 실행";
        checks.Add(new Check("path.shim", true,
            (shimOnPath && shimExe) ? $"{shimDir} 등록됨"
            : shimOnPath ? $"{shimDir} PATH에 있으나 vbaloop.exe 없음 — {reinstallHint}"
            : $"{shimDir} 미등록 — {reinstallHint} 시 자동 등록"));

        // 프로젝트 컨텍스트(있을 때만): 고아 스캔의 세션 제외와 워크북 잠금 체크가 공유
        ProjectConfig? cfg = null;
        SessionMeta? meta = null;
        try
        {
            cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
            meta = SessionMeta.Load(cfg.SessionMetaPath);
        }
        catch (CliError)
        {
            // doctor는 프로젝트 밖에서도 동작해야 함 — 프로젝트 의존 체크만 생략
        }

        // 5c. 격리 마커(설계 §11) — 정보성. 격리 PID는 아래 고아 정리에서 절대 제외된다.
        RecoveryMarker? marker = cfg != null ? RecoveryMarker.Load(cfg.RecoveryMarkerPath) : null;
        if (cfg != null)
            checks.Add(new Check("session.recovery", true,
                marker == null ? "격리 없음" : marker.Describe(),
                marker == null ? null : RecoveryGate.HintFor(marker)));

        // 6. 고아 Excel 후보: 창 없는 EXCEL.EXE. 기본은 리포트만 — --kill-orphans는 사용자 승인이고,
        //    그때도 OrphanFilter(현재 Windows 세션 + /automation + 비격리)를 통과한 것만 kill한다.
        //    이 프로젝트의 일회용 인스턴스(_dbg\ephemeral)는 실행 중인 명령의 숨은 Excel일 수 있어 제외한다 —
        //    끝난 명령의 잔재는 다음 세션 명령의 sweep이 격리 규칙으로 정리한다.
        var ephemeralPids = LiveEphemeralPids(cfg);
        var orphans = ScanOrphanExcel(opts, meta, marker, ephemeralPids);
        int killable = orphans.Count(z => z.Killable);
        checks.Add(new Check("excel.orphans", true,
            orphans.Count == 0
                ? "창 없는 고아 EXCEL 프로세스 없음"
                : $"창 없는 EXCEL 프로세스 {orphans.Count}개(정리 대상 {killable}개: 현재 세션·/automation·비격리)"
                  + (opts.KillOrphans || killable == 0 ? "" : " — 정리는 사용자 확인 후에만 vbaloop doctor --kill-orphans")));
        foreach (var z in orphans)
            checks.Add(new Check($"excel.orphans.{z.Pid}", true,
                $"PID {z.Pid} (시작 {z.StartTime ?? "미상"}) {z.Action} — {z.Reason}"));

        // 7. 대상 워크북 쓰기 잠금: 잠긴 채 session start하면 조용히 read-only로 열려
        //    Save가 SaveAs 오탈출한다. 고아 스캔 뒤에 배치해 교차 안내.
        //    잠금 보유 PID 특정(Restart Manager RmGetList)은 후속 과제 — 고아 목록 연계로 충분.
        if (cfg != null)
            checks.Add(CheckWorkbookLock(cfg, meta, orphans));

        bool allOk = checks.Where(c => c.Name is "excel.com" or "excel.accessVBOM" or "attach.oleacc"
                                       or "workbook.lock")
                           .All(c => c.Ok);
        o.Ok = allOk;
        foreach (var c in checks)
        {
            o.HumanLines.Add($"[{(c.Ok ? "ok" : "FAIL")}] {c.Name}: {c.Detail}");
            if (!c.Ok) o.Errors.Add($"{c.Name}: {c.Detail}");
        }
        if (!allOk)
            o.Hint = checks.FirstOrDefault(c => !c.Ok && c.Hint != null)?.Hint;

        o.Payload = new
        {
            checks = checks.Select(c => new { name = c.Name, ok = c.Ok, detail = c.Detail, hint = c.Hint }),
            orphans = orphans.Select(z => new
            {
                pid = z.Pid, startTime = z.StartTime, action = z.Action, killable = z.Killable,
                toolLaunched = z.ToolLaunched, reason = z.Reason,
            }),
        };
        return o;
    }

    private sealed record Orphan(int Pid, string? StartTime, string Action, bool Killable, bool ToolLaunched,
        string Reason);

    /// <summary>
    /// 대상 워크북 파일의 쓰기 잠금 점검. 우리 쪽 공유를 최대 허용(FileShare.ReadWrite)해
    /// 상대 프로세스의 deny-write(Excel의 xlsm 열기 방식)에서만 걸리게 한다 —
    /// Workbooks.Open이 read-only로 폴백하는 조건과 일치.
    /// </summary>
    private static Check CheckWorkbookLock(ProjectConfig cfg, SessionMeta? meta, List<Orphan> orphans)
    {
        if (!File.Exists(cfg.WorkbookPath))
            return new Check("workbook.lock", false, $"워크북 파일 없음: {cfg.WorkbookPath}",
                "vbaloop.toml의 [excel].workbook 경로 확인");

        string? lockError;
        try
        {
            using var fs = new FileStream(cfg.WorkbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            lockError = null;
        }
        catch (IOException ex) { lockError = ex.Message; }
        catch (UnauthorizedAccessException ex) { lockError = ex.Message; }

        if (lockError == null)
            return new Check("workbook.lock", true, "대상 워크북 쓰기 잠금 없음");

        bool sessionAlive = meta != null && ProcessIdentity.IsAlive(meta.Pid, meta.ExcelStartTime);
        if (sessionAlive)
            return new Check("workbook.lock", true, $"잠금 보유: 현재 세션(정상, PID {meta!.Pid})");

        var killableOrphans = orphans.Where(z => z.Killable).Select(z => z.Pid).ToList();
        string hint = killableOrphans.Count > 0
            ? $"창 없는 자동화 EXCEL(PID {string.Join(", ", killableOrphans)})이 잠금 보유 의심 — 사용자 확인 후 vbaloop doctor --kill-orphans"
            : "Excel/타 프로세스에서 대상 파일을 닫을 것";
        return new Check("workbook.lock", false,
            $"쓰기 잠금 보유 프로세스 존재 — session start가 read-only 거부(exit 3)로 실패하게 됨: {lockError}",
            hint);
    }

    /// <summary>
    /// 창 없는 EXCEL.EXE 열거 — 현재 프로젝트의 살아있는 세션 프로세스는 제외. 각 후보를
    /// OrphanFilter로 판정하고, kill은 --kill-orphans AND Killable일 때만. 실패해도 doctor는 계속.
    /// 도구가 띄운 PID(ExcelPidRegistry)를 앞에 정렬한다.
    /// </summary>
    private static HashSet<int> LiveEphemeralPids(ProjectConfig? cfg)
    {
        var pids = new HashSet<int>();
        if (cfg == null || !Directory.Exists(cfg.EphemeralDir)) return pids;
        foreach (var file in Directory.EnumerateFiles(cfg.EphemeralDir, "*.json"))
            if (SessionMeta.Load(file) is { } m && ProcessIdentity.IsAlive(m.Pid, m.ExcelStartTime))
                pids.Add(m.Pid);
        return pids;
    }

    private static List<Orphan> ScanOrphanExcel(CliArgs opts, SessionMeta? meta, RecoveryMarker? marker,
        HashSet<int> ephemeralPids)
    {
        var result = new List<Orphan>();
        var registry = ExcelPidRegistry.Load(ExcelPidRegistry.DefaultPath);
        int mySession;
        using (var self = Process.GetCurrentProcess()) mySession = self.SessionId;

        foreach (var p in Process.GetProcessesByName("EXCEL"))
        {
            using var proc = p;
            try
            {
                if (proc.MainWindowHandle != IntPtr.Zero) continue; // 창 있는 Excel은 사람/세션 소유로 간주
                if (meta != null && proc.Id == meta.Pid
                    && ProcessIdentity.IsAlive(meta.Pid, meta.ExcelStartTime)) continue; // 현재 세션 제외
                if (ephemeralPids.Contains(proc.Id)) continue; // 이 프로젝트의 일회용 인스턴스 제외

                string? started = ProcessIdentity.GetStartTimeIso(proc.Id);
                bool quarantined = (marker != null && (marker.Corrupt || marker.Pid == proc.Id))
                                   || ExcelPidRegistry.Has(registry, proc.Id, started, ExcelPidRegistry.KindQuarantined);
                bool toolLaunched = ExcelPidRegistry.Has(registry, proc.Id, started, ExcelPidRegistry.KindLaunched);
                var verdict = OrphanFilter.Classify(new OrphanCandidate(
                    proc.Id, started, HasMainWindow: false, proc.SessionId, ProcessCommandLine.Get(proc.Id),
                    IsLiveProjectSession: false, quarantined, toolLaunched), mySession);

                string action = "reported";
                if (opts.KillOrphans && verdict.Killable)
                {
                    try
                    {
                        proc.Kill(entireProcessTree: true);
                        // Kill은 비동기 — 곧이어 도는 workbook.lock 체크가 아직 잡혀 있는
                        // 파일 핸들을 오탐하지 않도록 실제 종료를 짧게 기다린다
                        proc.WaitForExit(3000);
                        action = "killed";
                    }
                    catch (Exception ex)
                    {
                        action = $"kill 실패: {ex.Message}";
                    }
                }
                result.Add(new Orphan(proc.Id, started, action, verdict.Killable, toolLaunched, verdict.Reason));
            }
            catch
            {
                // 열거 도중 프로세스가 사라진 경우 등 — 건너뜀
            }
        }
        return result.OrderByDescending(z => z.ToolLaunched).ToList();
    }

    // AccessVBOM 점검/설정은 Core.AccessVbom이 단일 소스 — setup 마법사와 공유한다.

    private static string? ReadExcelBitness()
    {
        using var c2r = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
        if (c2r?.GetValue("Platform") is string p)
            return p; // "x86" | "x64"
        using var outlook = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\Outlook");
        return outlook?.GetValue("Bitness") as string;
    }
}
