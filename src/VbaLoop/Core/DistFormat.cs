namespace VbaLoop.Core;

public enum DistOutputKind
{
    /// <summary>.xlsm — 매크로 워크북(base 포맷 유지, wb.Save()).</summary>
    MacroWorkbook,
    /// <summary>.xlam — Excel 애드인(IsAddin=True + SaveAs xlOpenXMLAddIn).</summary>
    AddIn,
}

/// <summary>
/// dist 출력 형식/경로 판정(설계 §9.7). 형식은 별도 플래그·toml 키 없이 --out 확장자로 선언된다
/// (확장자 = 의도 선언). 순수 함수 — Excel 불필요, xUnit 검증.
/// </summary>
public static class DistFormat
{
    /// <summary>xlOpenXMLAddIn — .xlam SaveAs FileFormat 상수(Excel 열거값).</summary>
    public const int XlOpenXMLAddIn = 55;

    /// <summary>xlOpenXMLWorkbookMacroEnabled — .xlsm SaveAs FileFormat 상수(Excel 열거값).
    /// setup이 빈 워크북을 만들 때 쓴다 — 기본 저장 형식(.xlsx)은 VBA를 통째로 버린다.</summary>
    public const int XlOpenXMLWorkbookMacroEnabled = 52;

    /// <summary>--out 확장자로 출력 종류 판정. 지원 밖 확장자는 CliError로 조기 거부(Excel 기동 전).</summary>
    public static DistOutputKind FromOutPath(string outPath)
    {
        string ext = Path.GetExtension(outPath).ToLowerInvariant();
        return ext switch
        {
            ".xlsm" => DistOutputKind.MacroWorkbook,
            ".xlam" => DistOutputKind.AddIn,
            _ => throw new CliError(
                $"--out 확장자 미지원: '{(ext.Length == 0 ? "(없음)" : ext)}'",
                "지원 확장자: .xlsm(매크로 워크북), .xlam(애드인)"),
        };
    }

    /// <summary>--out 생략 시 기본 출력 경로: &lt;root&gt;/dist/&lt;워크북명&gt;.xlsm.
    /// dist/는 init-관리 .gitignore로 무시된다(빌드 아티팩트 = 비커밋, §9.7).</summary>
    public static string DefaultOut(string rootDir, string workbookPath)
        => Path.Combine(rootDir, "dist",
            Path.GetFileNameWithoutExtension(workbookPath) + ".xlsm");
}
