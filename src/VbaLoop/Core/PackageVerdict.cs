namespace VbaLoop.Core;

/// <summary>
/// 패키지 적용의 대상·버전 판정(설계 §13.4). 워크북 문서 속성과 패키지 메타만 보는 순수 로직이다.
/// 같은 버전일 때만 워크북 코드를 비교해야 하므로 그 비교는 <c>codeSame</c> 콜백으로 미룬다.
/// </summary>
public static class PackageVerdict
{
    public enum Decision
    {
        /// <summary>적용한다(새 버전, 첫 적용, 허용된 다운그레이드).</summary>
        Apply,
        /// <summary>같은 버전·같은 코드 — 아무것도 하지 않고 성공(백업·저장 없음).</summary>
        AlreadyApplied,
        /// <summary>같은 버전·같은 패키지인데 현장에서 코드가 바뀜 — 복구로 다시 적용한다.</summary>
        Reapply,
        Refuse,
    }

    /// <summary>워크북 문서 속성(없으면 null).</summary>
    public sealed record WorkbookStamp(string? Target, string? Version, string? PackageSha256);

    /// <summary>vbaloop_Target 없음 거부의 hint — 드래그 앤 드롭 진입이 이것을 보고 채택 여부를 묻는다.</summary>
    public const string AdoptHint =
        "If this is the right workbook for this package, run again with --adopt (the first application records the target)";

    public sealed record Result(Decision Decision, string Reason, string? Hint = null, bool Adopting = false);

    public static Result Decide(WorkbookStamp wb, PackageFormat.Meta pkg, string pkgSha256,
        bool adopt, bool allowDowngrade, Func<bool> codeSame)
    {
        bool adopting = false;
        if (string.IsNullOrWhiteSpace(wb.Target))
        {
            if (!adopt)
                return new Result(Decision.Refuse,
                    $"workbook has no vbaloop_Target — it was never stamped by vbaloop dist or a package (package target: {pkg.Target})",
                    AdoptHint);
            adopting = true;
        }
        else if (!string.Equals(wb.Target.Trim(), pkg.Target, StringComparison.OrdinalIgnoreCase))
        {
            return new Result(Decision.Refuse,
                $"target mismatch: workbook is '{wb.Target.Trim()}', package is '{pkg.Target}'",
                "Use the package built for this workbook's project (no flag overrides a target mismatch)");
        }

        var pkgVer = SemVer.ParseOrNull(pkg.Version)
            ?? throw new CliError($"package Version is not semver: '{pkg.Version}'");
        if (string.IsNullOrWhiteSpace(wb.Version))
            return new Result(Decision.Apply, $"workbook has no version — applying {pkg.Version}", Adopting: adopting);

        var wbVer = SemVer.ParseOrNull(wb.Version);
        if (wbVer == null)
            return allowDowngrade
                ? new Result(Decision.Apply, $"workbook version '{wb.Version}' is not semver — applying {pkg.Version} (--allow-downgrade)", Adopting: adopting)
                : new Result(Decision.Refuse,
                    $"workbook version '{wb.Version}' is not semver, so it cannot be ordered against {pkg.Version}",
                    "Run again with --allow-downgrade if replacing it is intended");

        int cmp = pkgVer.CompareTo(wbVer);
        if (cmp > 0)
            return new Result(Decision.Apply, $"upgrading {wb.Version} -> {pkg.Version}", Adopting: adopting);
        if (cmp < 0)
            return allowDowngrade
                ? new Result(Decision.Apply, $"downgrading {wb.Version} -> {pkg.Version} (--allow-downgrade)", Adopting: adopting)
                : new Result(Decision.Refuse,
                    $"package {pkg.Version} is older than the workbook ({wb.Version})",
                    "Run again with --allow-downgrade if the downgrade is intended");

        if (codeSame())
            return new Result(Decision.AlreadyApplied, $"already applied: workbook code matches package {pkg.Version}", Adopting: adopting);
        if (string.Equals(wb.PackageSha256?.Trim(), pkgSha256, StringComparison.OrdinalIgnoreCase))
            return new Result(Decision.Reapply,
                $"workbook code was changed after package {pkg.Version} was applied — reapplying the same package", Adopting: adopting);
        return new Result(Decision.Refuse,
            $"workbook is already at {pkg.Version} but its code differs from this package — the same version number carries different code",
            "Raise [dist] version, build the package again, and apply that (no flag overrides this)");
    }
}
