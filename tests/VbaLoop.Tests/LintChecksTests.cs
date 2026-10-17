using System.Text;
using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class LintChecksTests
{
    static LintChecksTests()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    private static string NewDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "vbaloop-lint", Path.GetRandomFileName());
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void WriteRaw(string dir, string name, byte[] bytes)
        => File.WriteAllBytes(Path.Combine(dir, name), bytes);

    private static void WriteOk(string dir, string name, string vbName)
        => WriteRaw(dir, name, Encoding.UTF8.GetBytes(
            $"Attribute VB_Name = \"{vbName}\"\r\nOption Explicit\r\n"));

    private static List<string> Codes(List<LintChecks.Finding> f) => f.Select(x => x.Code).ToList();

    [Fact]
    public void CleanFile_NoFindings()
    {
        var dir = NewDir();
        try
        {
            WriteOk(dir, "modA.bas", "modA");
            Assert.Empty(LintChecks.CheckProject(dir, 949));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Bom_LoneLf_EofViolations_Detected()
    {
        var dir = NewDir();
        try
        {
            // BOM + LF-only + EOF 개행 없음
            var bom = new byte[] { 0xEF, 0xBB, 0xBF };
            var body = Encoding.UTF8.GetBytes("Attribute VB_Name = \"modB\"\nOption Explicit");
            WriteRaw(dir, "modB.bas", bom.Concat(body).ToArray());

            var codes = Codes(LintChecks.CheckProject(dir, 949));
            Assert.Contains("BOM", codes);
            Assert.Contains("EOL_LF", codes);
            Assert.Contains("EOF_NEWLINE", codes);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void TrailingBlankLine_IsLegitimateContent_NotFlagged()
    {
        // 본문 끝 빈 줄이 있는 모듈은 이중 CRLF로 끝나는 게 정당한 직렬화 — 위반 아님
        var dir = NewDir();
        try
        {
            WriteRaw(dir, "modC.bas", Encoding.UTF8.GetBytes(
                "Attribute VB_Name = \"modC\"\r\nOption Explicit\r\n\r\n"));
            Assert.DoesNotContain("EOF_NEWLINE", Codes(LintChecks.CheckProject(dir, 949)));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Encoding_NameMismatch_AttrAfterCode_Detected()
    {
        var dir = NewDir();
        try
        {
            // em dash(U+2014) + VB_Name 불일치
            WriteRaw(dir, "modD.bas", Encoding.UTF8.GetBytes(
                "Attribute VB_Name = \"다른이름\"\r\n' em dash — here\r\n"));
            // 코드 이후 Attribute
            WriteRaw(dir, "modE.bas", Encoding.UTF8.GetBytes(
                "Attribute VB_Name = \"modE\"\r\nPublic Function F() As Long\r\nAttribute F.VB_UserMemId = 0\r\nEnd Function\r\n"));
            // VB_Name 부재
            WriteRaw(dir, "modF.bas", Encoding.UTF8.GetBytes("Option Explicit\r\n"));

            var findings = LintChecks.CheckProject(dir, 949);
            var codes = Codes(findings);
            Assert.Contains("ENCODING", codes);
            Assert.Contains("NAME_MISMATCH", codes);
            Assert.Contains("HEADER_ATTR_AFTER_CODE", codes);
            Assert.Contains("NAME_MISSING", codes);

            var enc = findings.First(f => f.Code == "ENCODING");
            Assert.Equal(2, enc.Line); // em dash 라인 위치
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void FrxPairing_And_BasenameConflict_Detected()
    {
        var dir = NewDir();
        try
        {
            WriteOk(dir, "frmA.frm", "frmA");          // frx 없음
            WriteRaw(dir, "orphan.frx", new byte[] { 1, 2 }); // frm 없음
            WriteOk(dir, "dup.bas", "dup");
            WriteOk(dir, "DUP.cls", "DUP");            // basename 충돌 (대소문자 무시)

            var codes = Codes(LintChecks.CheckProject(dir, 949));
            Assert.Contains("FRX_MISSING", codes);
            Assert.Contains("FRX_ORPHAN", codes);
            Assert.Contains("BASENAME_CONFLICT", codes);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void MissingSrcRoot_Reported()
    {
        var findings = LintChecks.CheckProject(
            Path.Combine(Path.GetTempPath(), "vbaloop-none-" + Guid.NewGuid()), 949);
        Assert.Contains("SRC_MISSING", Codes(findings));
    }
}
