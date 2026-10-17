using VbaLoop.Core;
using Xunit;

namespace VbaLoop.Tests;

public class ExtraModulePlanTests
{
    private static HashSet<string> Src(params string[] names)
        => new(names, StringComparer.OrdinalIgnoreCase);

    private const int Std = 1, Cls = 2, Form = 3, Doc = 100;

    [Fact]
    public void DocumentModuleOnlyInWorkbook_IsIgnored_NotPrunable()
    {
        // 매크로가 만든 시트 코드모듈(Sheet2~)은 문서모듈 — push를 막으면 안 된다(이슈 A)
        var r = ExtraModulePlan.Classify(
            new[] { ("Sheet2", Doc), ("Sheet3", Doc), ("modProbe", Std) },
            Src("modProbe"));
        Assert.Empty(r.Prunable);
        Assert.Equal(new[] { "Sheet2", "Sheet3" }, r.IgnoredDocuments);
    }

    [Fact]
    public void StandardModuleOnlyInWorkbook_IsPrunable()
    {
        var r = ExtraModulePlan.Classify(
            new[] { ("modStale", Std), ("clsGone", Cls), ("frmOld", Form) },
            Src());
        Assert.Equal(new[] { "modStale", "clsGone", "frmOld" }, r.Prunable);
        Assert.Empty(r.IgnoredDocuments);
    }

    [Fact]
    public void ModulesInSrc_AreNeitherPrunableNorIgnored()
    {
        var r = ExtraModulePlan.Classify(
            new[] { ("modA", Std), ("ThisWorkbook", Doc) },
            Src("modA", "ThisWorkbook"));
        Assert.Empty(r.Prunable);
        Assert.Empty(r.IgnoredDocuments);
    }

    [Fact]
    public void EmptyProjectFirstPush_DocModulesDoNotBlock()
    {
        // 빈 프로젝트 최초 push: ThisWorkbook + Sheet1(둘 다 문서모듈)이 src에 없어도
        // 표준모듈 push를 차단하지 않는다 — 첫 힌트가 --prune을 오유도하던 문제 제거(이슈 B)
        var r = ExtraModulePlan.Classify(
            new[] { ("ThisWorkbook", Doc), ("Sheet1", Doc), ("modNew", Std) },
            Src("modNew"));
        Assert.Empty(r.Prunable);
        Assert.Equal(new[] { "ThisWorkbook", "Sheet1" }, r.IgnoredDocuments);
    }

    [Fact]
    public void NameMatch_IsCaseInsensitive()
    {
        var r = ExtraModulePlan.Classify(
            new[] { ("MODA", Std) },
            Src("moda"));
        Assert.Empty(r.Prunable);
    }
}
