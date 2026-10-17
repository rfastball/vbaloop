using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace VbaLoop.Core;

/// <summary>
/// setup의 자기설치 계층: 실행 중인 exe를 고정 폴더(<c>%LOCALAPPDATA%\vbaloop\bin</c>)에 복사하고
/// 사용자 PATH에 1회 등록한다. rollout.ps1 [3/4]단계를 exe 안으로 옮긴 것 — 배포본 사용자는
/// rollout.ps1을 갖고 있지 않다. PATH는 빌드 구성 경로가 아니라 고정 폴더를 가리킨다
/// (빌드 출력 경로가 바뀌어도 에이전트 셸이 같은 위치에서 exe를 찾도록).
/// </summary>
public static class SelfInstall
{
    /// <summary>고정 설치 폴더 — doctor의 path.shim 체크와 같은 경로를 본다.</summary>
    public static string InstallDir => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop", "bin");

    public static string InstalledExePath => Path.Combine(InstallDir, "vbaloop.exe");

    /// <summary>
    /// PATH 원문에 dir를 덧붙인 결과. 이미 있으면 (원문, false) — 트레일링 <c>\</c> 차이와
    /// 대소문자를 관용한다(doctor.IsOnPath와 같은 판정). 기존 항목은 원문 그대로 보존하고
    /// 끝에만 덧붙인다 — 사용자 PATH를 재포맷하지 않는다.
    /// </summary>
    public static (string NewPath, bool Changed) AddToPathValue(string? rawPath, string dir)
    {
        string target = dir.TrimEnd('\\');
        var entries = (rawPath ?? "").Split(';').Where(p => p.Length > 0).ToList();
        if (entries.Any(p => string.Equals(p.Trim().TrimEnd('\\'), target, StringComparison.OrdinalIgnoreCase)))
            return (rawPath ?? "", false);
        entries.Add(dir);
        return (string.Join(";", entries), true);
    }

    /// <summary>
    /// 실행 중인 exe를 설치 폴더로 복사. 이미 설치 폴더에서 실행 중이면 복사 없이 (false, 경로).
    /// 대상이 실행 중이면 파일 잠금으로 실패한다 — 호출부가 "실행 중인 vbaloop 종료 후 재시도"로 안내.
    /// </summary>
    public static (bool Copied, string Path) CopySelf()
    {
        string? src = Environment.ProcessPath
            ?? throw new CliError("실행 파일 경로를 알 수 없어 설치를 진행할 수 없음");
        string dest = InstalledExePath;
        if (PathDrift.SamePath(src, dest)) return (false, dest);

        Directory.CreateDirectory(InstallDir);
        try
        {
            File.Copy(src, dest, overwrite: true);
        }
        catch (IOException ex)
        {
            throw new CliError($"설치 폴더로 복사 실패: {ex.Message}",
                $"실행 중인 vbaloop(또는 {dest})를 종료한 뒤 재시도");
        }
        return (true, dest);
    }

    /// <summary>
    /// 사용자 PATH(HKCU\Environment)에 설치 폴더를 1회 등록. setx는 1024자 절단 위험이 있어
    /// 미사용 — 레지스트리에 ExpandString으로 직접 기록하고 원문(미확장)을 읽어 판정한다.
    /// 이미 등록돼 있으면 false.
    /// </summary>
    public static bool EnsureOnPath()
    {
        using var env = Registry.CurrentUser.OpenSubKey("Environment", writable: true)
            ?? throw new CliError(@"HKCU\Environment 키를 열 수 없음");
        string raw = env.GetValue("Path", "", RegistryValueOptions.DoNotExpandEnvironmentNames) as string ?? "";
        var (newPath, changed) = AddToPathValue(raw, InstallDir);
        if (!changed) return false;

        env.SetValue("Path", newPath, RegistryValueKind.ExpandString);
        BroadcastEnvironmentChange();
        // 현재 프로세스에도 즉시 반영 — 마법사가 이어서 부르는 단계들이 바로 쓸 수 있게 한다
        var live = Environment.GetEnvironmentVariable("Path") ?? "";
        Environment.SetEnvironmentVariable("Path", live.Length == 0 ? InstallDir : live + ";" + InstallDir);
        return true;
    }

    private static readonly IntPtr HWND_BROADCAST = new(0xFFFF);
    private const uint WM_SETTINGCHANGE = 0x001A;
    private const uint SMTO_ABORTIFHUNG = 0x0002;

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr SendMessageTimeoutW(
        IntPtr hWnd, uint msg, UIntPtr wParam, string lParam,
        uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

    /// <summary>
    /// 레지스트리에 직접 쓴 환경 변수는 탐색기가 다시 읽기 전까지 새 프로세스에 전파되지 않는다 —
    /// 탐색기/시작 메뉴에서 새로 여는 터미널이 vbaloop를 못 찾는다. setx와
    /// <c>Environment.SetEnvironmentVariable(User)</c>가 내부적으로 하는 WM_SETTINGCHANGE
    /// "Environment" 브로드캐스트를 같이 보낸다. 응답 없는 창은 건너뛰고(SMTO_ABORTIFHUNG),
    /// 실패해도 PATH 기록 자체는 유효하므로 best-effort(재로그온 시 어차피 반영).
    /// </summary>
    public static void BroadcastEnvironmentChange()
    {
        try
        {
            SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, UIntPtr.Zero, "Environment",
                SMTO_ABORTIFHUNG, 5000, out _);
        }
        catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException)
        {
            // 데스크톱 없는 환경(서버 코어 등) — 알릴 대상도 없다
        }
    }
}
