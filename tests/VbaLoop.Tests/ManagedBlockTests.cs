using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ManagedBlockTests
{
    private static readonly string[] Body = { "## Development loop (vbaloop)", "", "본문 라인 1", "- 규칙 A" };

    private static string Block(string eol = "\r\n")
        => ManagedBlock.StartMarker + eol + string.Join(eol, Body) + eol + ManagedBlock.EndMarker;

    [Fact]
    public void Upsert_NoFile_CreatesBlockOnly()
    {
        var (text, changed) = ManagedBlock.Upsert(null, Body);
        Assert.True(changed);
        Assert.Equal(Block() + "\r\n", text);
    }

    [Fact]
    public void Upsert_Append_WhenNoHeading()
    {
        var existing = "# 프로젝트\r\n\r\n설명.\r\n";
        var (text, changed) = ManagedBlock.Upsert(existing, Body);
        Assert.True(changed);
        Assert.Equal("# 프로젝트\r\n\r\n설명.\r\n\r\n" + Block() + "\r\n", text);
    }

    [Fact]
    public void Upsert_Replace_BetweenMarkers()
    {
        var existing = "# 헤더\r\n\r\n" + ManagedBlock.StartMarker + "\r\n옛 내용\r\n"
            + ManagedBlock.EndMarker + "\r\n\r\n## 아키텍처\r\n뒤쪽 유지\r\n";
        var (text, changed) = ManagedBlock.Upsert(existing, Body);
        Assert.True(changed);
        Assert.Equal("# 헤더\r\n\r\n" + Block() + "\r\n\r\n## 아키텍처\r\n뒤쪽 유지\r\n", text);
    }

    [Fact]
    public void Upsert_UnmarkedOwnHeading_ReplacedUpToNextH2()
    {
        var existing = "# 헤더\r\n\r\n## 개발 루프 (vbaloop)\r\n\r\n옛 절차 1\r\n옛 절차 2\r\n\r\n## 아키텍처\r\n뒤쪽 유지\r\n";
        var (text, changed) = ManagedBlock.Upsert(existing, Body);
        Assert.True(changed);
        // 다음 섹션과의 구분 빈 줄은 보존됨
        Assert.Equal("# 헤더\r\n\r\n" + Block() + "\r\n\r\n## 아키텍처\r\n뒤쪽 유지\r\n", text);
    }

    [Fact]
    public void Upsert_UnmarkedOwnHeading_ReplacedAtEof()
    {
        var existing = "# 헤더\r\n\r\n## 개발 루프 (vbaloop)\r\n옛 절차\r\n";
        var (text, changed) = ManagedBlock.Upsert(existing, Body);
        Assert.True(changed);
        Assert.Equal("# 헤더\r\n\r\n" + Block() + "\r\n", text);
    }

    [Fact]
    public void Upsert_UnmarkedEnglishHeading_ReplacedUpToNextH2()
    {
        var existing = "# 헤더\r\n\r\n## Development loop (vbaloop)\r\n옛 절차\r\n\r\n## 아키텍처\r\n뒤쪽 유지\r\n";
        var (text, changed) = ManagedBlock.Upsert(existing, Body);
        Assert.True(changed);
        Assert.Equal("# 헤더\r\n\r\n" + Block() + "\r\n\r\n## 아키텍처\r\n뒤쪽 유지\r\n", text);
    }

    [Fact]
    public void Upsert_Idempotent()
    {
        var existing = "# 헤더\r\n";
        var (first, _) = ManagedBlock.Upsert(existing, Body);
        var (second, changed) = ManagedBlock.Upsert(first, Body);
        Assert.False(changed);
        Assert.Equal(first, second);
    }

    [Fact]
    public void Upsert_PreservesLfFile()
    {
        var existing = "# 헤더\n\n## 개발 루프 (vbaloop)\n옛 절차\n\n## 아키텍처\n유지\n";
        var (text, changed) = ManagedBlock.Upsert(existing, Body);
        Assert.True(changed);
        Assert.DoesNotContain("\r\n", text);
        Assert.Equal("# 헤더\n\n" + Block("\n") + "\n\n## 아키텍처\n유지\n", text);
    }

    [Fact]
    public void Upsert_H3NotABoundary()
    {
        var existing = "## 개발 루프 (vbaloop)\r\n옛 절차\r\n### 세부\r\n세부 내용\r\n## 아키텍처\r\n유지\r\n";
        var (text, _) = ManagedBlock.Upsert(existing, Body);
        // "### 세부"는 경계가 아니므로 레거시 섹션과 함께 교체되어 사라져야 함
        Assert.DoesNotContain("### 세부", text);
        Assert.Contains("## 아키텍처", text);
    }

    [Fact]
    public void Upsert_StartMarkerWithoutEnd_Recovers()
    {
        // end 마커 유실 + 블록 자신의 헤딩이 남아있는 파손 상태에서
        // 다음 진짜 섹션(## 아키텍처)은 보존하고 잔해만 교체해야 함
        var existing = ManagedBlock.StartMarker + "\r\n## 개발 루프 (vbaloop)\r\n잔해 라인\r\n## 아키텍처\r\n유지\r\n";
        var (text, changed) = ManagedBlock.Upsert(existing, Body);
        Assert.True(changed);
        Assert.Equal(Block() + "\r\n## 아키텍처\r\n유지\r\n", text);
        Assert.DoesNotContain("잔해 라인", text);
    }
}
