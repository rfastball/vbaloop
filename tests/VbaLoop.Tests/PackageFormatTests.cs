using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class PackageFormatTests
{
    private static readonly PackageFormat.Meta Meta = new("지출관리", "1.3.0", "8acb552", "2026-11-02 12:00:00");

    private static List<PackageFormat.Module> Modules() => new()
    {
        new("modMain", ".bas", PackageModuleKind.Std,
            new[] { "Attribute VB_Name = \"modMain\"", "Option Explicit", "' 한글 주석", "Public Sub Go()", "End Sub" }, null),
        new("ThisWorkbook", ".cls", PackageModuleKind.Workbook,
            new[] { "VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "END", "Attribute VB_Name = \"ThisWorkbook\"", "Option Explicit" }, null),
        new("frmInput", ".frm", PackageModuleKind.Form,
            new[] { "VERSION 5.00", "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInput", "End", "Attribute VB_Name = \"frmInput\"", "Option Explicit" },
            Enumerable.Range(0, 200).Select(i => (byte)i).ToArray()),
    };

    private static byte[] Bytes(string text) => CanonicalText.Utf8NoBom.GetBytes(text);

    [Fact]
    public void WriteThenParse_RoundTrips()
    {
        string text = PackageFormat.Write(Meta, Modules());
        var pkg = PackageFormat.Parse(Bytes(text));

        Assert.Equal(Meta, pkg.Meta);
        Assert.Equal(new[] { "modMain.bas", "ThisWorkbook.cls", "frmInput.frm" }, pkg.Modules.Select(m => m.FileName));
        Assert.Equal(Modules()[0].Lines, pkg.Modules[0].Lines);
        Assert.Equal(PackageModuleKind.Workbook, pkg.Modules[1].Kind);
        Assert.Equal(Modules()[2].Frx, pkg.Modules[2].Frx);
        Assert.Null(pkg.Modules[0].Frx);
        Assert.Matches("^[0-9a-f]{64}$", pkg.Sha256);
        Assert.EndsWith($"@@END_PACKAGE sha256={pkg.Sha256}\r\n", text);
    }

    [Fact]
    public void Write_IsCanonicalCrlf()
    {
        string text = PackageFormat.Write(Meta, Modules());
        Assert.DoesNotContain("\n", text.Replace("\r\n", ""));
        Assert.StartsWith("@@VBALOOP_PACKAGE\r\nFormatVersion=1\r\nTarget=지출관리\r\n", text);
    }

    [Fact]
    public void Parse_LfAndBomTransfer_Passes()
    {
        string text = PackageFormat.Write(Meta, Modules());
        var lf = Encoding.UTF8.GetPreamble().Concat(Bytes(text.Replace("\r\n", "\n") + "\n\n")).ToArray();
        Assert.Equal(PackageFormat.Parse(Bytes(text)).Sha256, PackageFormat.Parse(lf).Sha256);
    }

    [Theory]
    [InlineData("Public Sub Go()", "Public Sub Gx()")]   // 본문 1바이트
    [InlineData("Version=1.3.0", "Version=9.3.0")]       // 메타도 해시 대상
    [InlineData("Target=지출관리", "Target=다른것")]
    public void Parse_EditedContent_HashMismatch(string from, string to)
    {
        string text = PackageFormat.Write(Meta, Modules());
        var ex = Assert.Throws<CliError>(() => PackageFormat.Parse(Bytes(text.Replace(from, to))));
        Assert.Contains("hash mismatch", ex.Message);
    }

    [Fact]
    public void Parse_Truncated_Refused()
    {
        string text = PackageFormat.Write(Meta, Modules());
        var ex = Assert.Throws<CliError>(() => PackageFormat.Parse(Bytes(text[..(text.Length / 2)])));
        Assert.Contains("truncated", ex.Message);
    }

    [Fact]
    public void Parse_Cp949Reencoded_Refused()
    {
        string text = PackageFormat.Write(Meta, Modules());
        var cp949 = CodePagesEncodingProvider.Instance.GetEncoding(949)!.GetBytes(text);
        var ex = Assert.Throws<CliError>(() => PackageFormat.Parse(cp949));
        Assert.Contains("UTF-8", ex.Message);
    }

    [Fact]
    public void Parse_UnknownFormatVersion_AsksForExeUpdate()
    {
        string text = PackageFormat.Write(Meta, Modules()).Replace("FormatVersion=1", "FormatVersion=2");
        var ex = Assert.Throws<CliError>(() => PackageFormat.Parse(Bytes(text)));
        Assert.Contains("FormatVersion", ex.Message);
        Assert.Contains("Update vbaloop", ex.Hint);
    }

    [Fact]
    public void Parse_NotAPackage_Refused()
        => Assert.Throws<CliError>(() => PackageFormat.Parse(Bytes("hello\r\n")));

    /// <summary>헤더·블록을 고친 뒤 해시를 다시 맞춘 패키지 — 구조 검사가 해시와 별개로 동작하는지 본다.</summary>
    private static byte[] Resealed(string text, Func<List<string>, List<string>> edit)
    {
        var lines = edit(CanonicalText.SplitLines(text).ToList());
        lines.RemoveAt(lines.Count - 1);
        lines.Add("@@END_PACKAGE sha256=" + PackageFormat.Sha256Of(lines));
        return Bytes(CanonicalText.ToCanonical(lines));
    }

    [Fact]
    public void Parse_PostUpdateHook_Refused()
    {
        var bytes = Resealed(PackageFormat.Write(Meta, Modules()), l => { l.Insert(2, "PostUpdate=modMig.Run"); return l; });
        var ex = Assert.Throws<CliError>(() => PackageFormat.Parse(bytes));
        Assert.Contains("PostUpdate", ex.Message);
    }

    [Fact]
    public void Parse_UnknownMetaKey_Refused()
    {
        var bytes = Resealed(PackageFormat.Write(Meta, Modules()), l => { l.Insert(2, "Scope=manifest"); return l; });
        Assert.Contains("unknown package meta", Assert.Throws<CliError>(() => PackageFormat.Parse(bytes)).Message);
    }

    [Theory]
    [InlineData(@"..\evil.bas")]
    [InlineData("C:evil.bas")]
    [InlineData("sub/evil.bas")]
    [InlineData("evil.exe")]
    public void Parse_UnsafeFileName_Refused(string name)
    {
        var bytes = Resealed(PackageFormat.Write(Meta, Modules()),
            l => l.Select(x => x.Replace("name=modMain.bas", "name=" + name)).ToList());
        Assert.Throws<CliError>(() => PackageFormat.Parse(bytes));
    }

    [Fact]
    public void Parse_ModulesListMismatch_Refused()
    {
        var bytes = Resealed(PackageFormat.Write(Meta, Modules()),
            l => l.Select(x => x.StartsWith("Modules=") ? "Modules=modMain.bas,ThisWorkbook.cls" : x).ToList());
        Assert.Contains("not listed", Assert.Throws<CliError>(() => PackageFormat.Parse(bytes)).Message);
    }

    [Fact]
    public void Parse_NonSemverVersion_Refused()
    {
        var bytes = Resealed(PackageFormat.Write(Meta, Modules()), l => l.Select(x => x == "Version=1.3.0" ? "Version=20261102_1200" : x).ToList());
        Assert.Contains("semver", Assert.Throws<CliError>(() => PackageFormat.Parse(bytes)).Message);
    }

    [Theory]
    [InlineData("@@END_FILE")]
    [InlineData("  @@FILE name=x.bas")]
    [InlineData("@@END_PACKAGE sha256=00")]
    public void Write_ReservedMarkerInBody_Refused(string line)
    {
        var mods = Modules();
        mods[0] = mods[0] with { Lines = mods[0].Lines.Append(line).ToArray() };
        Assert.Throws<CliError>(() => PackageFormat.Write(Meta, mods));
    }

    [Fact]
    public void Write_LineBreakInMeta_Refused()
        => Assert.Throws<CliError>(() => PackageFormat.Write(Meta with { Target = "a\r\nb" }, Modules()));
}
