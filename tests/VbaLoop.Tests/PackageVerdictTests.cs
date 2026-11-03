using VbaLoop.Core;
using Xunit;
using static VbaLoop.Core.PackageVerdict;

namespace VbaLoop.Tests;

public class PackageVerdictTests
{
    private const string Sha = "abc123";
    private static readonly PackageFormat.Meta Pkg = new("App", "1.3.0", null, null);

    private static Result Decide(string? target, string? version, string? sha = null,
        bool adopt = false, bool downgrade = false, bool codeSame = false)
        => PackageVerdict.Decide(new WorkbookStamp(target, version, sha), Pkg, Sha, adopt, downgrade, () => codeSame);

    [Fact]
    public void NoTarget_Refused_UnlessAdopt()
    {
        var r = Decide(null, null);
        Assert.Equal(Decision.Refuse, r.Decision);
        Assert.Contains("--adopt", r.Hint);
        Assert.Equal("adoption_required", r.Code);

        var a = Decide(null, null, adopt: true);
        Assert.Equal(Decision.Apply, a.Decision);
        Assert.True(a.Adopting);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void TargetMismatch_AlwaysRefused(bool adopt, bool downgrade)
        => Assert.Equal(Decision.Refuse, Decide("Other", "1.0.0", adopt: adopt, downgrade: downgrade).Decision);

    [Fact]
    public void Target_IgnoresCase()
        => Assert.Equal(Decision.Apply, Decide("app", "1.2.0").Decision);

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("1.2.9")]
    [InlineData("1.3.0-rc.1")]   // prerelease < 정식
    public void NewerOrMissingVersion_Applies(string? wbVersion)
    {
        var r = Decide("App", wbVersion);
        Assert.Equal(Decision.Apply, r.Decision);
        Assert.False(r.Adopting);
    }

    [Fact]
    public void Older_Refused_UnlessAllowDowngrade()
    {
        Assert.Equal(Decision.Refuse, Decide("App", "2.0.0").Decision);
        Assert.Equal(Decision.Apply, Decide("App", "2.0.0", downgrade: true).Decision);
    }

    [Fact]
    public void NonSemverWorkbookVersion_Refused_UnlessAllowDowngrade()
    {
        var r = Decide("App", "20261102_1200");
        Assert.Equal(Decision.Refuse, r.Decision);
        Assert.Contains("semver", r.Reason);
        Assert.Equal(Decision.Apply, Decide("App", "20261102_1200", downgrade: true).Decision);
    }

    [Fact]
    public void SameVersion_SameCode_AlreadyApplied()
        => Assert.Equal(Decision.AlreadyApplied, Decide("App", "1.3.0", codeSame: true).Decision);

    [Fact]
    public void SameVersion_BuildMetadataOnlyDiffers_IsSameVersion()
        => Assert.Equal(Decision.AlreadyApplied, Decide("App", "1.3.0+other", codeSame: true).Decision);

    [Fact]
    public void SameVersion_CodeChangedInField_SamePackage_Reapplies()
        => Assert.Equal(Decision.Reapply, Decide("App", "1.3.0", sha: "ABC123").Decision);

    [Theory]
    [InlineData(null)]
    [InlineData("ffff")]
    public void SameVersion_DifferentCode_DifferentPackage_Refused(string? sha)
    {
        var r = Decide("App", "1.3.0", sha: sha, downgrade: true);
        Assert.Equal(Decision.Refuse, r.Decision);
        Assert.Contains("[dist] version", r.Hint);
    }

    [Fact]
    public void CodeCompared_OnlyForSameVersion()
    {
        bool called = false;
        PackageVerdict.Decide(new WorkbookStamp("App", "1.0.0", null), Pkg, Sha, false, false, () => called = true);
        Assert.False(called);
    }
}
