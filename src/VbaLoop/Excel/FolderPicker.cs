namespace VbaLoop.Excel;

/// <summary>
/// 폴더 선택 대화상자. <c>Shell.Application.BrowseForFolder</c>를 COM 늦은 바인딩으로 부른다 —
/// WinForms(FolderBrowserDialog)를 쓰면 Microsoft.WindowsDesktop.App 프레임워크 참조가 붙어
/// self-contained 배포 크기가 커지는데, 이 도구는 이미 Excel COM에 같은 경로를 쓰고 있어
/// 의존성 추가가 0이다.
/// </summary>
public static class FolderPicker
{
    // BIF_RETURNONLYFSDIRS(0x0001) | BIF_NEWDIALOGSTYLE(0x0040) — 파일시스템 폴더만, 새 폴더 버튼 포함
    private const int Options = 0x0041;

    /// <summary>선택된 폴더의 전체 경로. 취소하거나 대화상자를 띄울 수 없으면 null.</summary>
    public static string? Pick(string title)
    {
        var shellType = Type.GetTypeFromProgID("Shell.Application");
        if (shellType == null) return null;

        dynamic? shell = null;
        try
        {
            shell = Activator.CreateInstance(shellType);
            if (shell == null) return null;
            dynamic? folder = shell.BrowseForFolder(0, title, Options);
            if (folder == null) return null;
            // Self는 FolderItem — 파일시스템 폴더에서만 Path가 실제 경로다(가상 폴더는 빈 문자열)
            string? path = folder.Self?.Path as string;
            return string.IsNullOrWhiteSpace(path) ? null : path;
        }
        catch
        {
            return null; // 대화상자 불가(세션 0, Shell 미등록 등) — 호출부가 콘솔 입력으로 폴백
        }
    }
}
