namespace VbaLoop.Core;

/// <summary>
/// diff 명령의 순수 비교 로직(IO/COM 없음, 설계 §4 diff). 비교 의미론은
/// "pull이 이 src 파일을 다시 쓸 것인가" — PullCommand의 직렬화 파이프라인
/// (본문 SplitLines + 기존 src 헤더 보존/합성 + ToCanonical + Utf8NoBom)을 그대로
/// 미러링해 바이트 기준으로 판정한다. BOM/개행 위반은 formatOnly로 구분해 드러낸다.
/// 폼은 코드비하인드만 라인 비교(.frx/레이아웃은 Export 비결정성으로 비교 제외).
/// </summary>
public static class ModuleDiff
{
    /// <summary>Kind가 null이면 미지원 컴포넌트 타입(RawType으로 보고) — verdict skipped.</summary>
    public sealed record WorkbookModule(string Name, VbaComponentKind? Kind, string Body, int RawType = 0);

    /// <summary>src 파일의 원시 바이트 — BOM/개행 위반이 비교에 드러나도록 디코딩 전 상태로 받는다.</summary>
    public sealed record SrcFile(string BaseName, string Extension, byte[] Bytes);

    public sealed record Detail(
        int? FirstDiffLine = null,
        string? WorkbookLine = null,
        string? SrcLine = null,
        int? WorkbookLines = null,
        int? SrcLines = null,
        bool? FormatOnly = null,
        bool? LayoutCompared = null);

    /// <summary>Verdict: same | changed | casingOnly | onlyInWorkbook | onlyInSrc | skipped | managed.</summary>
    public sealed record Entry(
        string Name, string Kind, string? File, string Verdict, Detail? Detail = null, string? Note = null);

    private const int ExcerptMax = 160;

    public static IReadOnlyList<Entry> Compare(IReadOnlyList<WorkbookModule> workbook, IReadOnlyList<SrcFile> src)
    {
        var srcByFile = new Dictionary<string, SrcFile>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in src)
            srcByFile[f.BaseName + f.Extension] = f;

        var componentNames = new HashSet<string>(workbook.Select(m => m.Name), StringComparer.OrdinalIgnoreCase);
        var entries = new List<Entry>();

        foreach (var m in workbook)
        {
            if (m.Kind == null)
            {
                entries.Add(new Entry(m.Name, "unknown", null, "skipped",
                    Note: $"component type not supported: {m.RawType}"));
                continue;
            }
            // 벤더링 관리 모듈(LogShim)은 정본이 init이지 워크북↔src 비교가 아니다. COM 왕복이
            // 본문을 미세 변형해도 changed로 오보하지 않고 managed로
            // 분리 보고 — clean 판정에서 제외된다. src 무결성은 init --check가 담당.
            if (VendoredModules.IsVendored(m.Name))
            {
                entries.Add(new Entry(m.Name, KindName(m.Kind.Value), null, "managed",
                    Note: "vendored managed module — canonical copy is owned by init (verify with vbaloop init --check)"));
                continue;
            }
            string ext = m.Kind switch
            {
                VbaComponentKind.StdModule => ".bas",
                VbaComponentKind.MsForm => ".frm",
                _ => ".cls",
            };
            srcByFile.TryGetValue(m.Name + ext, out var file);
            entries.Add(m.Kind == VbaComponentKind.MsForm
                ? CompareForm(m, file)
                : CompareCode(m, file, m.Kind.Value));
        }

        // pull의 orphan 판정 미러: 어느 컴포넌트와도 이름이 안 맞는 src 파일 (.frx는 대상 외)
        foreach (var f in src)
        {
            if (componentNames.Contains(f.BaseName)) continue;
            entries.Add(new Entry(f.BaseName, KindFromExtension(f.Extension), f.BaseName + f.Extension, "onlyInSrc"));
        }
        return entries;
    }

    private static Entry CompareCode(WorkbookModule m, SrcFile? file, VbaComponentKind kind)
    {
        string kindName = KindName(kind);
        var bodyLines = m.Body.Length > 0 ? CanonicalText.SplitLines(m.Body) : Array.Empty<string>();
        if (file == null)
            return new Entry(m.Name, kindName, null, "onlyInWorkbook",
                new Detail(WorkbookLines: bodyLines.Length));

        string fileName = file.BaseName + file.Extension;
        var srcLines = CanonicalText.SplitLines(DecodeStripBom(file.Bytes));

        // PullCommand.PullCodeComponent 미러: 기존 src 헤더 엔벨로프 보존, VB_Name 부재 시 합성
        IReadOnlyList<string> headerLines = VbaHeader.Split(srcLines, fileName, strict: false).Header;
        if (!headerLines.Any(l => l.TrimStart().StartsWith("Attribute VB_Name", StringComparison.Ordinal)))
            headerLines = VbaHeader.SynthesizeHeader(kind, m.Name);

        var expectedLines = headerLines.Concat(bodyLines).ToArray();
        byte[] expectedBytes = CanonicalText.Utf8NoBom.GetBytes(CanonicalText.ToCanonical(expectedLines));
        if (file.Bytes.AsSpan().SequenceEqual(expectedBytes))
            return new Entry(m.Name, kindName, fileName, "same");

        return Changed(m.Name, kindName, fileName, expectedLines, srcLines,
            lineOffset: 0, layoutCompared: null, casingAware: true);
    }

    private static Entry CompareForm(WorkbookModule m, SrcFile? file)
    {
        // 폼 왕복 수렴 규칙: 끝 빈 줄 제거(pull의 Import 상쇄 미러) + 선행 빈 줄 제거 —
        // 폼 CodeModule은 선언부 자리의 선행 빈 줄을 포함하지만 Export 텍스트(=src .frm)에는
        // 없다(SampleRpa 실측). pull은 Export 경로라 이 줄을 못 보므로 diff도 무시해야
        // "pull이 다시 쓸 것인가" 의미론이 유지된다.
        // 멤버 Attribute(.VB_VarHelpID 등) 스트립: pull이 더는 src에 쓰지 않지만(버그리포트
        // 2026-07-11), 이미 박힌 기존 src가 CodeModule.Lines(해당 라인 미포함)와 달라 changed로
        // 오탐하는 것을 방어 — pull의 PullForm 미러.
        var wbBody = CanonicalText.StripMemberAttributes(TrimSurroundingEmpty(
            m.Body.Length > 0 ? CanonicalText.SplitLines(m.Body) : Array.Empty<string>(), out _));
        if (file == null)
            return new Entry(m.Name, "form", null, "onlyInWorkbook",
                new Detail(WorkbookLines: wbBody.Count, LayoutCompared: false));

        string fileName = file.BaseName + file.Extension;
        var split = VbaHeader.Split(CanonicalText.SplitLines(DecodeStripBom(file.Bytes)), fileName, strict: false);
        var srcBody = CanonicalText.StripMemberAttributes(TrimSurroundingEmpty(split.Body, out int srcLeading));

        if (wbBody.SequenceEqual(srcBody, StringComparer.Ordinal))
            return new Entry(m.Name, "form", fileName, "same", new Detail(LayoutCompared: false));

        return Changed(m.Name, "form", fileName, wbBody, srcBody,
            lineOffset: split.Header.Count + srcLeading, layoutCompared: false, casingAware: true);
    }

    private static IReadOnlyList<string> TrimSurroundingEmpty(IReadOnlyList<string> lines, out int leading)
    {
        var trimmedTail = CanonicalText.TrimTrailingEmpty(lines);
        leading = 0;
        while (leading < trimmedTail.Length && trimmedTail[leading].Trim().Length == 0) leading++;
        return trimmedTail.Skip(leading).ToArray();
    }

    /// <summary>
    /// firstDiffLine은 1-기준 src 파일 행 번호(코드 모듈은 전체 파일 라인끼리 비교하므로
    /// offset 0, 폼은 본문끼리 비교하므로 src 헤더 길이를 더한다). 한쪽이 다른 쪽의
    /// 접두이면 짧은 쪽 길이+1 — 그 행 발췌는 존재하는 쪽만 채운다.
    /// </summary>
    private static Entry Changed(string name, string kindName, string file,
        IReadOnlyList<string> expected, IReadOnlyList<string> actual, int lineOffset,
        bool? layoutCompared, bool casingAware = false)
    {
        bool formatOnly = expected.SequenceEqual(actual, StringComparer.Ordinal);
        // 케이싱만 상이(라인 수 동일 + 모든 쌍이 식별자 케이싱만 다름): VBE 프로젝트 전역 재케이싱.
        // pull이 src 철자를 보존해 재기록하지 않으므로 clean 판정에서
        // 제외한다. 라인별 VbaCasing.LineEquivalent로 판정 — 문자열 리터럴·주석 변경은 케이싱으로
        // 오분류하지 않는다(구 전체-라인 OrdinalIgnoreCase 결함 해소). 라인 수가 다르면(삽입/삭제)
        // 구조 변경이므로 changed. 코드·폼 모두 casingAware(pull이 양쪽 다 케이싱 보존).
        bool casingOnly = casingAware && !formatOnly
            && expected.Count == actual.Count
            && AllPairsLineEquivalent(expected, actual);
        int? firstDiff = null;
        string? wbLine = null, srcLine = null;
        if (!formatOnly)
        {
            int n = Math.Min(expected.Count, actual.Count);
            int i = 0;
            while (i < n && string.Equals(expected[i], actual[i], StringComparison.Ordinal)) i++;
            firstDiff = lineOffset + i + 1;
            wbLine = i < expected.Count ? Truncate(expected[i]) : null;
            srcLine = i < actual.Count ? Truncate(actual[i]) : null;
        }
        return new Entry(name, kindName, file, casingOnly ? "casingOnly" : "changed", new Detail(
            FirstDiffLine: firstDiff, WorkbookLine: wbLine, SrcLine: srcLine,
            WorkbookLines: expected.Count, SrcLines: actual.Count,
            FormatOnly: formatOnly, LayoutCompared: layoutCompared),
            casingOnly
                ? "identifier casing differs only — VBE global re-casing; pull preserves src spelling"
                : null);
    }

    private static string KindName(VbaComponentKind kind) => kind switch
    {
        VbaComponentKind.StdModule => "stdModule",
        VbaComponentKind.ClassModule => "classModule",
        VbaComponentKind.Document => "document",
        _ => "form",
    };

    private static string KindFromExtension(string ext) => ext switch
    {
        ".bas" => "stdModule",
        ".cls" => "classModule",
        ".frm" => "form",
        _ => "unknown",
    };

    /// <summary>CanonicalText.ReadFileStripBom의 바이트 배열 대응 — 디코더가 남긴 선행 U+FEFF 제거.</summary>
    private static string DecodeStripBom(byte[] bytes)
    {
        var text = System.Text.Encoding.UTF8.GetString(bytes);
        if (text.Length > 0 && text[0] == '\uFEFF') text = text[1..];
        return text;
    }

    /// <summary>라인 수가 같은 두 블록의 모든 대응 쌍이 식별자 케이싱만 다른지(위치 정렬).</summary>
    private static bool AllPairsLineEquivalent(IReadOnlyList<string> a, IReadOnlyList<string> b)
    {
        for (int i = 0; i < a.Count; i++)
            if (!VbaCasing.LineEquivalent(a[i], b[i])) return false;
        return true;
    }

    private static string Truncate(string s) => s.Length <= ExcerptMax ? s : s[..ExcerptMax];
}
