using System.Security.Cryptography;

namespace VbaLoop.Core;

/// <summary>GUI는 원본을 Excel에 열지 않는다. 불변 스냅샷에서 작업하고 새 파일로만 저장한다.</summary>
public sealed class GuiWorkbookCopy : IDisposable
{
    public string WorkPath { get; }
    public string? OutputPath { get; }
    private readonly string _workFolder;

    public GuiWorkbookCopy(string source, string expectedHash, bool plan, string? root = null)
    {
        root ??= Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop");
        string id = DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N");
        _workFolder = Path.Combine(root, "work", id);
        WorkPath = Path.Combine(_workFolder, Path.GetFileName(source));
        Directory.CreateDirectory(_workFolder);
        try
        {
            // 공유 삭제/이름 변경을 막는 읽기 핸들은 허용하되, 복사 중 쓰기는 허용하지 않는다.
            using (var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var output = new FileStream(WorkPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                input.CopyTo(output);
            if (!PackageUiSession.FileHash(WorkPath).Equals(expectedHash, StringComparison.OrdinalIgnoreCase))
                throw new CliError("Workbook changed while making the working copy; inspect again", code: "files_changed");
            if (!plan)
            {
                string folder = Path.Combine(root, "updated", id);
                Directory.CreateDirectory(folder);
                OutputPath = Path.Combine(folder, Path.GetFileName(source));
            }
        }
        catch { Dispose(); throw; }
    }

    public static void RequireNewOutput(string path)
    {
        if (File.Exists(path) || Directory.Exists(path))
            throw new CliError("The output path already exists; no existing file will be replaced", code: "output_exists");
    }

    public static string VerifyOutput(string path)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (file.Length == 0) throw new CliError("The output file is empty; save state is unknown", code: "save_uncertain");
        return Convert.ToHexString(SHA256.HashData(file));
    }

    public void Dispose()
    {
        // 이 실행이 만든 입력 사본만 정리한다. 결과/원본은 실패해도 삭제하지 않는다.
        try { File.Delete(WorkPath); Directory.Delete(_workFolder, recursive: false); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
