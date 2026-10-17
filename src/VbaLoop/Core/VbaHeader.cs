namespace VbaLoop.Core;

public enum VbaComponentKind
{
    StdModule,     // vbext_ct_StdModule = 1
    ClassModule,   // vbext_ct_ClassModule = 2
    MsForm,        // vbext_ct_MSForm = 3
    Document,      // vbext_ct_Document = 100
}

/// <summary>
/// Export 호환 헤더의 스트립/보존/합성 (설계 §2 파일 포맷, §4 pull/push).
/// 스트립 로직은 _VbaHarness ModuleIO.bas Mio_InjectDocument의 일반화 +
/// 폼 레이아웃의 중첩 Begin/End 블록 깊이 추적 확장.
/// </summary>
public static class VbaHeader
{
    public sealed record SplitResult(IReadOnlyList<string> Header, IReadOnlyList<string> Body);

    /// <summary>
    /// 선두 헤더 엔벨로프(VERSION / Begin..End 블록 / Attribute / 코드 전 빈 줄)와 본문을 분리.
    /// strict=true(push 경로)면 코드 시작 이후 Attribute 라인 발견 시 CliError —
    /// 프로시저 수준 속성은 CodeModule 경로로 왕복 불가(설계 §4 push 1단계).
    /// strict=false(pull의 기존 헤더 보존)는 본문 검증 없이 헤더 프리픽스만 취한다.
    /// </summary>
    public static SplitResult Split(string[] lines, string sourceName, bool strict = true)
    {
        var header = new List<string>();
        var body = new List<string>();
        bool codeStarted = false;
        int beginDepth = 0;

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            string t = line.Trim();

            if (!codeStarted)
            {
                if (beginDepth > 0)
                {
                    header.Add(line);
                    if (IsBlockEnd(t)) beginDepth--;
                    else if (IsNestedBeginLine(t)) beginDepth++;
                    continue;
                }

                if (t.StartsWith("VERSION ", StringComparison.OrdinalIgnoreCase))
                {
                    header.Add(line);
                }
                else if (IsTopLevelBeginLine(t))
                {
                    header.Add(line);
                    beginDepth++;
                }
                else if (t.StartsWith("Attribute ", StringComparison.Ordinal))
                {
                    header.Add(line);
                }
                else if (t.Length == 0)
                {
                    header.Add(line); // 코드 시작 전 빈 줄은 헤더 쪽에 귀속
                }
                else
                {
                    codeStarted = true;
                    body.Add(line);
                }
            }
            else
            {
                if (strict && t.StartsWith("Attribute ", StringComparison.Ordinal))
                    throw new CliError(
                        $"{sourceName}:{i + 1}: 코드 시작 이후 Attribute 라인 발견 — 프로시저 수준 속성은 이 경로로 왕복 불가",
                        "해당 속성을 제거하거나 이 모듈만 수동 Export/Import로 관리");
                body.Add(line);
            }
        }

        // 헤더 끝의 빈 줄들은 본문 소속으로 돌리지 않고 그대로 헤더에 남긴다(보존 우선).
        return new SplitResult(header, body);
    }

    /// <summary>최상위 헤더 블록: cls의 "BEGIN" 단독 또는 폼의 "Begin {GUID} 이름"만 인정
    /// ("Begin = 5" 같은 코드를 헤더로 오인하지 않도록).</summary>
    private static bool IsTopLevelBeginLine(string t)
        => t.Equals("BEGIN", StringComparison.OrdinalIgnoreCase)
           || t.StartsWith("Begin {", StringComparison.Ordinal);

    /// <summary>블록 내부 중첩 컨트롤: "Begin MSForms.CommandButton cmdOK" 형태 허용.</summary>
    private static bool IsNestedBeginLine(string t)
        => t.Equals("BEGIN", StringComparison.OrdinalIgnoreCase)
           || t.StartsWith("Begin ", StringComparison.Ordinal);

    private static bool IsBlockEnd(string t)
        => t.Equals("End", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 신규 컴포넌트용 타입별 표준 헤더 합성(설계 §4 pull).
    /// VB_PredeclaredId 등은 VBIDE로 읽을 수 없으므로 비표준 클래스는 합성 불가 —
    /// 호출부가 경고를 출력한다.
    /// </summary>
    public static string[] SynthesizeHeader(VbaComponentKind kind, string name)
    {
        return kind switch
        {
            VbaComponentKind.StdModule => new[]
            {
                $"Attribute VB_Name = \"{name}\"",
            },
            VbaComponentKind.ClassModule => ClassHeader(name, predeclared: false, exposed: false),
            VbaComponentKind.Document => ClassHeader(name, predeclared: true, exposed: true),
            _ => throw new CliError($"헤더 합성 불가 타입: {kind} ({name})",
                "폼은 Export 경로로만 생성됨"),
        };
    }

    private static string[] ClassHeader(string name, bool predeclared, bool exposed) => new[]
    {
        "VERSION 1.0 CLASS",
        "BEGIN",
        "  MultiUse = -1  'True",
        "END",
        $"Attribute VB_Name = \"{name}\"",
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        $"Attribute VB_PredeclaredId = {(predeclared ? "True" : "False")}",
        $"Attribute VB_Exposed = {(exposed ? "True" : "False")}",
    };
}
