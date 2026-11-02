namespace VbaLoop.Core;

/// <summary>
/// dist 텍스트 패키지의 모듈 목록(설계 §13.1). 사본 VBProject에 가하는 변형(DistCommand.MutateProject)과
/// 같은 규칙을 src 텍스트에 적용한다: 포함 모듈만, (선택) 주석 제거, LogShim 무해화, 합성 modBuildInfo.
/// 종류는 사본의 컴포넌트 타입에서 받는다 — 이름·헤더로는 CodeName을 바꾼 통합 문서 모듈을 가릴 수 없다.
/// 순수 함수(frx 바이트는 호출자가 읽어 넘긴다).
/// </summary>
public static class DistPackageBuilder
{
    /// <param name="kinds">컴포넌트 이름 → 사본에서 본 종류.</param>
    /// <param name="readFrx">폼 이름 → src .frx 바이트(없으면 CliError).</param>
    public static List<PackageFormat.Module> Build(
        IReadOnlyList<DistModule> included, IReadOnlyDictionary<string, PackageModuleKind> kinds,
        bool strip, string version, string rev, string builtAt, Func<string, byte[]> readFrx)
    {
        var modules = new List<PackageFormat.Module>();
        foreach (var m in included)
        {
            if (m.Name.Equals(VersionStamp.ModuleName, StringComparison.OrdinalIgnoreCase))
                continue; // 아래에서 합성한 것으로 대체
            if (!kinds.TryGetValue(m.Name, out var kind))
                throw new CliError($"{m.Name}{m.Ext} is not a component of the base workbook",
                    "Apply src to the workbook with vbaloop push, then retry");

            IReadOnlyList<string> body = m.Body;
            if (strip) body = VbaCommentStripper.Strip(body);
            if (VendoredModules.IsVendored(m.Name)) body = LogShimRelease.FlipToRelease(body);

            byte[]? frx = kind == PackageModuleKind.Form ? readFrx(m.Name) : null;
            modules.Add(new PackageFormat.Module(m.Name, m.Ext, kind, m.Header.Concat(body).ToArray(), frx));
        }

        var stamp = CanonicalText.SplitLines(VersionStamp.BuildInfoModule(version, rev, builtAt));
        modules.Add(new PackageFormat.Module(VersionStamp.ModuleName, ".bas", PackageModuleKind.Std, stamp, null));
        return modules;
    }
}
