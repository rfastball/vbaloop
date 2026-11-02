using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace VbaLoop.Core;

/// <summary>패키지 안 모듈의 종류(설계 §13.2). dist가 사본의 컴포넌트 타입을 보고 기록한다 — 이름이나
/// 헤더로 추측하면 CodeName을 바꾼 통합 문서 모듈과 시트 모듈을 구분할 수 없다.</summary>
public enum PackageModuleKind { Std, Class, Form, Document, Workbook }

/// <summary>
/// 텍스트 패키지 포맷(설계 §13.2): 메타 블록 + 모듈별 <c>@@FILE</c> 블록 + 꼬리 줄 하나의 해시.
/// 해시는 꼬리 줄 직전까지 전체(메타 포함)를 정규화(BOM 제거, CRLF 통일)해 계산한다 — 망연계가 줄끝만
/// 바꾼 경우는 통과하고, 메타 수정·잘림·재인코딩은 거부한다. 순수 함수(IO/COM 없음).
/// </summary>
public static class PackageFormat
{
    public const string FormatVersion = "1";

    public const string MarkPackage = "@@VBALOOP_PACKAGE";
    public const string MarkEndMeta = "@@END_META";
    public const string MarkFile = "@@FILE";
    public const string MarkEndFile = "@@END_FILE";
    public const string MarkEndPackage = "@@END_PACKAGE";

    private const string FooterPrefix = MarkEndPackage + " sha256=";
    private const int Base64LineWidth = 76;

    private static readonly string[] ReservedMarkers = { MarkPackage, MarkEndMeta, MarkFile, MarkEndFile, MarkEndPackage };

    /// <summary>파일명은 경로 구분자·와일드카드·공백 없는 단일 이름 + 허용 확장자. 반입물이 쓰기 경로를 정하지 못하게 한다.</summary>
    private static readonly Regex SafeFileName = new(@"^[^\\/:*?""<>|\s.]+\.(bas|cls|frm|frx)$",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    public sealed record Meta(string Target, string Version, string? Rev, string? BuiltAt);

    /// <param name="Name">컴포넌트 이름(확장자 제외).</param>
    /// <param name="Ext">".bas" / ".cls" / ".frm".</param>
    /// <param name="Lines">src 규약 텍스트(헤더 + 본문).</param>
    /// <param name="Frx">폼의 .frx 바이트(폼 외에는 null).</param>
    public sealed record Module(string Name, string Ext, PackageModuleKind Kind, IReadOnlyList<string> Lines, byte[]? Frx)
    {
        public string FileName => Name + Ext;
    }

    public sealed record Package(Meta Meta, IReadOnlyList<Module> Modules, string Sha256);

    // ---- 쓰기 ----

    /// <summary>패키지 텍스트(CRLF, 끝 개행 1개). 본문 행이 예약 마커로 시작하거나 메타 값에 개행이 있으면 CliError.</summary>
    public static string Write(Meta meta, IReadOnlyList<Module> modules)
    {
        var lines = new List<string>
        {
            MarkPackage,
            "FormatVersion=" + FormatVersion,
            "Target=" + MetaValue("Target", meta.Target),
            "Version=" + MetaValue("Version", meta.Version),
        };
        if (meta.Rev != null) lines.Add("Rev=" + MetaValue("Rev", meta.Rev));
        if (meta.BuiltAt != null) lines.Add("BuiltAt=" + MetaValue("BuiltAt", meta.BuiltAt));
        lines.Add("Modules=" + string.Join(",", modules.Select(m => m.FileName)));
        lines.Add(MarkEndMeta);

        foreach (var m in modules)
        {
            for (int i = 0; i < m.Lines.Count; i++)
                if (IsReservedLine(m.Lines[i]))
                    throw new CliError(
                        $"{m.FileName}:{i + 1}: line starts with a reserved package marker ({m.Lines[i].Trim()})",
                        "Change that line (a VBA line cannot start with @@ outside a string or comment continuation)");
            lines.Add("");
            lines.Add($"{MarkFile} name={m.FileName} kind={KindText(m.Kind)} encoding=utf-8");
            lines.AddRange(m.Lines);
            lines.Add(MarkEndFile);

            if (m.Frx != null)
            {
                lines.Add("");
                lines.Add($"{MarkFile} name={m.Name}.frx encoding=base64");
                string b64 = Convert.ToBase64String(m.Frx);
                for (int i = 0; i < b64.Length; i += Base64LineWidth)
                    lines.Add(b64.Substring(i, Math.Min(Base64LineWidth, b64.Length - i)));
                lines.Add(MarkEndFile);
            }
        }
        lines.Add("");

        string hash = Sha256Of(lines);
        lines.Add(FooterPrefix + hash);
        return CanonicalText.ToCanonical(lines);
    }

    private static string MetaValue(string key, string value)
    {
        if (value.IndexOfAny(new[] { '\r', '\n' }) >= 0)
            throw new CliError($"package {key} contains a line break: '{value}'");
        return value.Trim();
    }

    private static bool IsReservedLine(string line)
    {
        string t = line.TrimStart();
        return t.StartsWith("@@", StringComparison.Ordinal)
               && ReservedMarkers.Any(m => t.StartsWith(m, StringComparison.Ordinal));
    }

    /// <summary>해시 대상 행들(꼬리 줄 제외)의 정규형 SHA-256 소문자 hex.</summary>
    public static string Sha256Of(IReadOnlyList<string> lines)
        => Convert.ToHexString(SHA256.HashData(CanonicalText.Utf8NoBom.GetBytes(CanonicalText.ToCanonical(lines))))
            .ToLowerInvariant();

    // ---- 읽기 ----

    /// <summary>
    /// 패키지 바이트를 검증하며 읽는다. 엄격 UTF-8 → 머리 표식 → FormatVersion → 해시 → 메타 → 블록 순으로
    /// 거부 사유를 정한다. 거부는 전부 CliError(exit 2) — 패키지를 고쳐 다시 만드는 것 말고 다음 행동이 없다.
    /// </summary>
    public static Package Parse(byte[] bytes)
    {
        string text;
        try { text = StrictUtf8.GetString(bytes); }
        catch (DecoderFallbackException)
        {
            throw new CliError("package is not valid UTF-8 — it was re-encoded in transit (e.g. to CP949)",
                "Transfer the original .txt file without conversion (binary mode), or build it again with vbaloop dist");
        }
        if (text.Length > 0 && text[0] == '﻿') text = text[1..];

        var lines = CanonicalText.TrimTrailingEmpty(CanonicalText.SplitLines(text));
        if (lines.Length == 0 || lines[0] != MarkPackage)
            throw new CliError($"not a vbaloop package (first line is not {MarkPackage})",
                "Pass the .txt file made by vbaloop dist --out <file>.txt");

        // 포맷 버전이 먼저다 — 새 포맷은 해시 규칙부터 다를 수 있다
        string? formatVersion = lines.TakeWhile(l => l != MarkEndMeta)
            .Where(l => l.StartsWith("FormatVersion=", StringComparison.Ordinal))
            .Select(l => l["FormatVersion=".Length..].Trim())
            .FirstOrDefault();
        if (formatVersion != FormatVersion)
            throw new CliError($"unsupported package FormatVersion: {formatVersion ?? "(missing)"} (this exe reads {FormatVersion})",
                "Update vbaloop (in an offline environment, bring in the newer vbaloop.exe as well)");

        string last = lines[^1];
        if (!last.StartsWith(FooterPrefix, StringComparison.Ordinal))
            throw new CliError($"package is truncated — the last line is not '{FooterPrefix}<hash>'",
                "Transfer the whole file again, or build it again with vbaloop dist");
        string expected = last[FooterPrefix.Length..].Trim();
        string actual = Sha256Of(lines[..^1]);
        if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
            throw new CliError($"package hash mismatch — the content was changed after dist (expected {expected}, got {actual})",
                "Do not edit the package by hand; build it again with vbaloop dist");

        int i = 1;
        var meta = ParseMeta(lines, ref i);
        var files = ParseBlocks(lines, i);
        return new Package(meta.Meta, Assemble(meta.ModuleNames, files), actual);
    }

    private static (Meta Meta, IReadOnlyList<string> ModuleNames) ParseMeta(string[] lines, ref int i)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (; i < lines.Length && lines[i] != MarkEndMeta; i++)
        {
            int eq = lines[i].IndexOf('=');
            if (eq <= 0) throw new CliError($"package meta line {i + 1} is not key=value: '{lines[i]}'");
            string key = lines[i][..eq].Trim();
            if (!values.TryAdd(key, lines[i][(eq + 1)..].Trim()))
                throw new CliError($"package meta key appears twice: {key}");
        }
        if (i >= lines.Length) throw new CliError($"package meta block has no {MarkEndMeta}");
        i++;

        // 적용 후 훅은 예약만 돼 있다(§13.8) — 조용히 건너뛰면 마이그레이션 없이 새 코드만 들어간다
        if (values.ContainsKey("PostUpdate"))
            throw new CliError("package declares a PostUpdate hook, which this exe does not support",
                "Update vbaloop, or build the package without [dist] postUpdate");
        var known = new HashSet<string>(StringComparer.Ordinal) { "FormatVersion", "Target", "Version", "Rev", "BuiltAt", "Modules" };
        var unknown = values.Keys.Where(k => !known.Contains(k)).ToList();
        if (unknown.Count > 0)
            throw new CliError($"unknown package meta keys: {string.Join(", ", unknown)}",
                "Update vbaloop (the package may come from a newer version)");

        string target = Required(values, "Target");
        string version = Required(values, "Version");
        if (!SemVer.TryParse(version, out _))
            throw new CliError($"package Version is not semver: '{version}'");
        var modules = Required(values, "Modules").Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        return (new Meta(target, version, values.GetValueOrDefault("Rev"), values.GetValueOrDefault("BuiltAt")), modules);
    }

    private static string Required(Dictionary<string, string> values, string key)
        => values.TryGetValue(key, out var v) && v.Length > 0 ? v
            : throw new CliError($"package meta is missing {key}");

    private sealed record RawFile(string Name, string? Kind, string Encoding, List<string> Lines);

    private static Dictionary<string, RawFile> ParseBlocks(string[] lines, int i)
    {
        var files = new Dictionary<string, RawFile>(StringComparer.OrdinalIgnoreCase);
        int end = lines.Length - 1; // 꼬리 줄
        while (i < end)
        {
            if (lines[i].Length == 0) { i++; continue; }
            if (!lines[i].StartsWith(MarkFile + " ", StringComparison.Ordinal))
                throw new CliError($"package line {i + 1}: expected {MarkFile}, got '{Excerpt(lines[i])}'");

            var attrs = ParseAttributes(lines[i], i + 1);
            string name = attrs.GetValueOrDefault("name") ?? throw new CliError($"package line {i + 1}: {MarkFile} has no name=");
            if (!SafeFileName.IsMatch(name))
                throw new CliError($"package line {i + 1}: file name not allowed: '{name}'");
            string encoding = attrs.GetValueOrDefault("encoding") ?? throw new CliError($"package line {i + 1}: {MarkFile} has no encoding=");
            bool isFrx = name.EndsWith(".frx", StringComparison.OrdinalIgnoreCase);
            if (encoding != (isFrx ? "base64" : "utf-8"))
                throw new CliError($"package line {i + 1}: {name} must use encoding={(isFrx ? "base64" : "utf-8")}");
            attrs.TryGetValue("kind", out var kind);
            if (isFrx == (kind != null))
                throw new CliError($"package line {i + 1}: {name} " + (isFrx ? "must not have kind=" : "has no kind="));

            int start = ++i;
            while (i < end && lines[i] != MarkEndFile) i++;
            if (i >= end) throw new CliError($"package block {name} has no {MarkEndFile}");
            if (!files.TryAdd(name, new RawFile(name, kind, encoding, lines[start..i].ToList())))
                throw new CliError($"package has two blocks for {name}");
            i++;
        }
        return files;
    }

    private static Dictionary<string, string> ParseAttributes(string header, int lineNo)
    {
        var attrs = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var token in header[MarkFile.Length..].Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            int eq = token.IndexOf('=');
            if (eq <= 0 || !attrs.TryAdd(token[..eq], token[(eq + 1)..]))
                throw new CliError($"package line {lineNo}: bad {MarkFile} attribute '{token}'");
        }
        return attrs;
    }

    /// <summary><c>Modules</c> 목록과 블록을 대조해 모듈로 묶는다. 목록이 mirror 판정의 기준이라 어긋나면 거부.</summary>
    private static IReadOnlyList<Module> Assemble(IReadOnlyList<string> moduleNames, Dictionary<string, RawFile> files)
    {
        var modules = new List<Module>();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int workbooks = 0;
        foreach (var fileName in moduleNames)
        {
            if (!files.TryGetValue(fileName, out var f))
                throw new CliError($"package Modules lists {fileName} but has no block for it");
            if (!used.Add(fileName)) throw new CliError($"package Modules lists {fileName} twice");
            string ext = Path.GetExtension(fileName).ToLowerInvariant();
            if (ext == ".frx") throw new CliError($"package Modules must not list .frx files: {fileName}");
            var kind = ParseKind(f.Kind!, fileName);
            bool kindFits = ext switch
            {
                ".bas" => kind == PackageModuleKind.Std,
                ".frm" => kind == PackageModuleKind.Form,
                _ => kind is PackageModuleKind.Class or PackageModuleKind.Document or PackageModuleKind.Workbook,
            };
            if (!kindFits) throw new CliError($"package {fileName}: kind={f.Kind} does not fit the extension");
            if (kind == PackageModuleKind.Workbook && ++workbooks > 1)
                throw new CliError("package has more than one kind=workbook module");

            byte[]? frx = null;
            if (kind == PackageModuleKind.Form)
            {
                string frxName = Path.GetFileNameWithoutExtension(fileName) + ".frx";
                if (!files.TryGetValue(frxName, out var fx))
                    throw new CliError($"package form {fileName} has no {frxName} block");
                used.Add(frxName);
                try { frx = Convert.FromBase64String(string.Concat(fx.Lines)); }
                catch (FormatException) { throw new CliError($"package {frxName} is not valid base64"); }
            }
            modules.Add(new Module(Path.GetFileNameWithoutExtension(fileName), ext, kind, f.Lines, frx));
        }
        var stray = files.Keys.Where(k => !used.Contains(k)).ToList();
        if (stray.Count > 0)
            throw new CliError($"package has blocks not listed in Modules: {string.Join(", ", stray)}");
        return modules;
    }

    public static string KindText(PackageModuleKind kind) => kind switch
    {
        PackageModuleKind.Std => "std",
        PackageModuleKind.Class => "class",
        PackageModuleKind.Form => "form",
        PackageModuleKind.Document => "document",
        _ => "workbook",
    };

    private static PackageModuleKind ParseKind(string text, string fileName) => text switch
    {
        "std" => PackageModuleKind.Std,
        "class" => PackageModuleKind.Class,
        "form" => PackageModuleKind.Form,
        "document" => PackageModuleKind.Document,
        "workbook" => PackageModuleKind.Workbook,
        _ => throw new CliError($"package {fileName}: unknown kind={text}"),
    };

    private static string Excerpt(string s) => s.Length <= 80 ? s : s[..80] + "...";
}
