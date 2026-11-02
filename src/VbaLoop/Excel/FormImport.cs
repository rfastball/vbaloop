using VbaLoop.Core;

namespace VbaLoop.Excel;

/// <summary>
/// 폼 Import(설계 §3 폼 예외): 시스템 ACP로 쓴 임시 frm + frx를 VBComponents.Import하고 결과 이름을 검사한다.
/// push(src)와 패키지 적용(§13.3)이 같은 절차를 쓴다. 기존 폼의 검증형 Remove는 호출자 몫이다.
/// </summary>
public static class FormImport
{
    /// <param name="onImported">Import 직후, 이름 검사 전에 새 컴포넌트를 받는다(원복 기록용).</param>
    public static void Import(dynamic proj, string name, IReadOnlyList<string> header, IReadOnlyList<string> body,
        byte[] frx, int codePage, Action<object>? onImported = null)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "vbaloop", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            var enc = Cp949Guard.GetStrict(codePage);
            // 헤더 끝(Attribute 블록 뒤)과 본문 끝의 빈 줄을 정규화 — Export가 이 지점에
            // 빈 줄을 누적시키는 실측 동작 때문에 트림 없이는 왕복마다 파일이 자란다
            string content = CanonicalText.ToCanonical(
                CanonicalText.TrimTrailingEmpty(header).Concat(CanonicalText.TrimTrailingEmpty(body)).ToList());
            string frmTemp = Path.Combine(tempDir, name + ".frm");
            File.WriteAllBytes(frmTemp, enc.GetBytes(content));
            File.WriteAllBytes(Path.Combine(tempDir, name + ".frx"), frx);

            dynamic newComp = ComRetry.Invoke(() => proj.VBComponents.Import(frmTemp));
            onImported?.Invoke(newComp);
            string actualName = Convert.ToString(newComp.Name) ?? "";

            // Import 후 중복 이름 검사 (폼 예외 c) — <name>1 패턴이면 save 전에 실패
            if (!actualName.Equals(name, StringComparison.OrdinalIgnoreCase))
                throw new CliError(
                    $"{name}: Import result name is {actualName} — suspected duplicate copy (duplicate module signature), save aborted");
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { /* 무시 */ }
        }
    }
}
