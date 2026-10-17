using Microsoft.Win32;

namespace VbaLoop.Core;

/// <summary>
/// "VBA 프로젝트 개체 모델에 안전하게 액세스"(AccessVBOM) 레지스트리 점검/설정.
/// 기본값이 꺼짐이고, 꺼져 있으면 push/pull/test가 전부 실패하므로 setup의 첫 관문이다.
/// 보안 설정이라 자동으로 켜지 않는다 — 켜기는 <c>doctor --fix</c>/setup의 명시 동의 경로에서만
/// (doctor의 <c>--kill-orphans</c>와 같은 "명시가 곧 승인" 원칙).
/// </summary>
public static class AccessVbom
{
    /// <summary>점검 대상 Office 버전(최신 우선). 16.0=2016 이후, 15.0=2013, 14.0=2010.</summary>
    private static readonly string[] Versions = { "16.0", "15.0", "14.0" };

    private static string KeyPath(string version) => $@"Software\Microsoft\Office\{version}\Excel\Security";

    /// <summary>
    /// <c>Excel.Application\CurVer</c>(예: "Excel.Application.16")에서 레지스트리 버전 폴더명을
    /// 유도한다. Security 키가 아직 없는 환경(설치 직후 등)에서 어느 버전에 기록할지 정하는 근거 —
    /// 순수 함수라 xUnit으로 검증한다.
    /// </summary>
    public static string? VersionFromCurVer(string? curVer)
    {
        if (string.IsNullOrWhiteSpace(curVer)) return null;
        int dot = curVer.LastIndexOf('.');
        if (dot < 0 || dot + 1 >= curVer.Length) return null;
        string major = curVer[(dot + 1)..].Trim();
        return major.Length > 0 && major.All(char.IsDigit) ? major + ".0" : null;
    }

    /// <summary>
    /// 현재 상태. Ok=true면 활성. Version은 판정 근거가 된 Office 버전(키를 못 찾으면 null).
    /// 기존 doctor 체크와 동일한 의미론 — 최초로 발견한 Security 키가 권위다.
    /// </summary>
    public static (bool Ok, string Detail, string? Version) Detect()
    {
        foreach (var v in Versions)
        {
            using var key = Registry.CurrentUser.OpenSubKey(KeyPath(v));
            if (key == null) continue;
            var val = key.GetValue("AccessVBOM");
            if (val is int i)
                return (i == 1, $"Office {v} AccessVBOM={i}", v);
            return (false, $"Office {v} AccessVBOM value missing (disabled by default)", v);
        }
        return (false, "Office Excel Security registry key not found", null);
    }

    /// <summary>
    /// AccessVBOM=1 기록. 기록 대상은 Detect가 찾은 버전, 없으면 curVer에서 유도한 버전의 키를
    /// 새로 만든다. 이미 활성이면 (false, ...). 어느 버전인지 전혀 알 수 없으면 CliError —
    /// 임의 버전에 키를 흩뿌리지 않는다.
    /// </summary>
    public static (bool Changed, string Detail) Enable(string? curVer)
    {
        var (ok, _, version) = Detect();
        if (ok) return (false, "already enabled");

        string target = version
            ?? VersionFromCurVer(curVer)
            ?? throw new CliError("Cannot enable AccessVBOM: Office version could not be determined",
                "Manually check Excel Options > Trust Center > Macro Settings > 'Trust access to the VBA project object model'");

        using var key = Registry.CurrentUser.CreateSubKey(KeyPath(target), writable: true)
            ?? throw new CliError($@"cannot create registry key: HKCU\{KeyPath(target)}");
        key.SetValue("AccessVBOM", 1, RegistryValueKind.DWord);
        return (true, $"Office {target} AccessVBOM=1 written");
    }
}
