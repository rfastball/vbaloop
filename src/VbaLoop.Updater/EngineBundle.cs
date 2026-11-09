using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;

namespace VbaLoop.Updater;

public sealed record EngineManifest(string Version, string Sha256, long Size, int Protocol);

/// <summary>런타임 없는 CLI DLL을 내장한다. PATH/설치본을 참조하지 않고 해시별 캐시만 로드한다.</summary>
public static class EngineBundle
{
    public static string DefaultCache => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "vbaloop-updater", "engine");

    public static (string Path, EngineManifest Manifest) Extract(string? cacheRoot = null)
    {
        var assembly = typeof(EngineBundle).Assembly;
        using var metadata = assembly.GetManifestResourceStream("VbaLoop.Updater.Engine.json")
            ?? throw new InvalidOperationException("실행 도구가 내장되지 않은 개발 빌드입니다. build-updater.ps1로 빌드해 주세요.");
        var manifest = JsonSerializer.Deserialize<EngineManifest>(metadata, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException("내장 실행 도구 정보를 읽을 수 없습니다.");
        using var source = assembly.GetManifestResourceStream("VbaLoop.Updater.Engine.dll")
            ?? throw new InvalidDataException("내장 실행 도구가 없습니다.");
        return (ExtractVerified(source, manifest, cacheRoot ?? DefaultCache), manifest);
    }

    public static string ExtractVerified(Stream source, EngineManifest manifest, string cacheRoot)
    {
        if (manifest.Protocol != 1 || manifest.Size <= 0 || manifest.Sha256.Length != 64 || !manifest.Sha256.All(Uri.IsHexDigit))
            throw new InvalidDataException("내장 실행 도구 정보가 올바르지 않습니다.");
        // 해시가 릴리스 버전뿐 아니라 같은 버전의 재빌드까지 구분한다.
        string folder = Path.Combine(cacheRoot, manifest.Sha256.ToLowerInvariant());
        Directory.CreateDirectory(folder);
        string target = Path.Combine(folder, "vbaloop.dll");
        using var mutex = new Mutex(false, @"Local\vbaloop-updater-extract-" + manifest.Sha256[..24]);
        bool held;
        try { held = mutex.WaitOne(TimeSpan.FromSeconds(30)); }
        catch (AbandonedMutexException) { held = true; }
        if (!held) throw new IOException("다른 업데이터가 준비 중입니다. 잠시 후 다시 시도해 주세요.");
        string temporary = Path.Combine(folder, Guid.NewGuid().ToString("N") + ".tmp");
        try
        {
            if (Matches(target, manifest)) return target;
            using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                source.CopyTo(output);
                output.Flush(true);
            }
            if (!Matches(temporary, manifest))
                throw new InvalidDataException("실행 도구 무결성 검사에 실패했습니다. 배포 파일을 다시 받아 주세요.");
            File.Move(temporary, target, overwrite: true);
            return target;
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
            mutex.ReleaseMutex();
        }
    }

    private static bool Matches(string path, EngineManifest manifest)
    {
        if (!File.Exists(path) || new FileInfo(path).Length != manifest.Size) return false;
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert.ToHexString(SHA256.HashData(stream)).Equals(manifest.Sha256, StringComparison.OrdinalIgnoreCase);
    }
}
