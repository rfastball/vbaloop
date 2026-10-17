namespace VbaLoop.Core;

/// <summary>
/// 실행 중일 수 있는 exe 교체(설계 §12.5). Windows는 실행 중인 exe를 지우거나 덮어쓸 수 없지만 이름은
/// 바꿀 수 있다 — 대상을 vbaloop.exe.old로 옮기고 새 exe를 원래 이름으로 복사한다. .old는 다음 기동에서
/// 지운다(CleanupOld). 복사가 실패하면 .old를 되돌려 설치본이 비지 않게 한다.
/// </summary>
public static class SelfReplace
{
    public const string OldSuffix = ".old";

    /// <summary>newExe를 target 자리에 설치. target이 없으면 그냥 복사. 실패 시 CliError(원상 복구 후).</summary>
    public static void Replace(string newExe, string target)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(target)!);
        string? moved = null;
        if (File.Exists(target))
        {
            moved = target + OldSuffix;
            if (!TryDelete(moved))
                moved = $"{target}.{Guid.NewGuid().ToString("N")[..8]}{OldSuffix}"; // 이전 .old가 아직 실행 중
            try
            {
                File.Move(target, moved);
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
                throw new CliError($"기존 exe를 옮기지 못함: {e.Message}", $"{target}을 쓰는 프로그램을 종료한 뒤 재시도");
            }
        }
        try
        {
            File.Copy(newExe, target, overwrite: false);
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            if (moved != null)
            {
                try { File.Move(moved, target, overwrite: true); } catch { /* 아래 에러가 상황을 말한다 */ }
            }
            throw new CliError($"새 exe 복사 실패: {e.Message}", "잠시 후 vbaloop update 재시도");
        }
    }

    /// <summary>dir의 vbaloop.exe*.old를 best-effort로 지운다(아직 실행 중이면 다음 기회에).</summary>
    public static int CleanupOld(string dir)
    {
        int n = 0;
        try
        {
            foreach (var f in Directory.EnumerateFiles(dir, "*" + OldSuffix))
            {
                var name = Path.GetFileName(f);
                if (!name.StartsWith("vbaloop.exe", StringComparison.OrdinalIgnoreCase)) continue;
                if (TryDelete(f)) n++;
            }
        }
        catch
        {
            // 정리는 부수적 — 실패가 명령을 깨면 안 된다
        }
        return n;
    }

    private static bool TryDelete(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
