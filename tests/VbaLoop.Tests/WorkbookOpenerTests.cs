using VbaLoop.Excel;
using VbaLoop.Core;
using Xunit;
using static VbaLoop.Excel.ModalHunter;

namespace VbaLoop.Tests;

public class WorkbookOpenerTests
{
    // Test delta: 권장 설정 때문에 쓰기 요청이 읽기 전용이 되는 위험. 기존 opener 소유자에서
    // COM 경계의 옵션만 대역으로 검증한다. 2개 사례, Excel/프로세스/파일 기동 없음.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReadOnlyRecommendation_DoesNotOverrideRequestedAccess(bool readOnly)
    {
        var app = new RecommendedWorkbookApp();
        var result = WorkbookOpener.Open(app, 0, "recommended.xlsm", readOnly, new ToolLog(null, "test"));
        Assert.Equal(readOnly, ((OpenedWorkbook)result.Workbook).ReadOnly);
    }

    public sealed class RecommendedWorkbookApp
    {
        public RecommendedWorkbooks Workbooks { get; } = new();
    }

    public sealed class RecommendedWorkbooks
    {
        // DisplayAlerts=false인 Excel은 권장 프롬프트를 생략하면 읽기 전용으로 연다.
        public OpenedWorkbook Open(string path, int updateLinks, bool readOnly, bool IgnoreReadOnlyRecommended = false)
            => new(readOnly || !IgnoreReadOnlyRecommended);
    }

    public sealed record OpenedWorkbook(bool ReadOnly);

    private const string File = "지출원인행위정보 송신 목록.xlsm";

    // 실측 CollectInfo(한국어 Excel 16, 2026-11-02)
    private const string CrashPrompt =
        "Microsoft Excel | 예(&Y) | 아니요(&N) | 마지막으로 '지출원인행위정보 송신 목록.xlsm'을(를) 열었을 때 오류가 발생했습니다. 그래도 이 문서를 여시겠습니까? | 700164";

    private static ButtonInfo Yes => new("예(&Y)", WorkbookOpener.IDYES, true);
    private static ButtonInfo No => new("아니요(&N)", IDNO, true);

    [Fact]
    public void CrashPrompt_Matches()
    {
        Assert.True(WorkbookOpener.IsCrashReopenPrompt(new[] { Yes, No }, CrashPrompt, File));
    }

    [Fact]
    public void CrashPrompt_MatchesRegardlessOfLanguage()
    {
        string en = "Microsoft Excel | &Yes | &No | The last time you opened '" + File
                    + "', it caused a serious error. Do you still want to open it? | 700164";
        Assert.True(WorkbookOpener.IsCrashReopenPrompt(
            new[] { new ButtonInfo("&Yes", WorkbookOpener.IDYES, true), new ButtonInfo("&No", IDNO, true) }, en, File));
    }

    [Fact]
    public void OtherYesNoPrompt_ForSameFile_IsNotAccepted()
    {
        // 같은 예/아니요 구조의 복구 프롬프트 — 알림 ID가 다르면 "예"를 누르지 않는다
        string repair = "Microsoft Excel | 예(&Y) | 아니요(&N) | '" + File
                        + "'의 일부 내용에 문제가 있습니다. 최대한 복구하시겠습니까? | 700110";
        Assert.False(WorkbookOpener.IsCrashReopenPrompt(new[] { Yes, No }, repair, File));
    }

    [Fact]
    public void CrashPrompt_ForAnotherFile_IsNotAccepted()
    {
        Assert.False(WorkbookOpener.IsCrashReopenPrompt(new[] { Yes, No }, CrashPrompt, "other.xlsm"));
    }

    [Fact]
    public void ButtonShapeMustBeYesNo()
    {
        Assert.False(WorkbookOpener.IsCrashReopenPrompt(new[] { No }, CrashPrompt, File));
        Assert.False(WorkbookOpener.IsCrashReopenPrompt(
            new[] { Yes, No, new ButtonInfo("취소", IDCANCEL, true) }, CrashPrompt, File));
        Assert.False(WorkbookOpener.IsCrashReopenPrompt(
            new[] { Yes with { Enabled = false }, No }, CrashPrompt, File));
        // 도움말 버튼은 판정에서 빠진다
        Assert.True(WorkbookOpener.IsCrashReopenPrompt(
            new[] { Yes, No, new ButtonInfo("도움말", IDHELP, true) }, CrashPrompt, File));
    }

    [Fact]
    public void AlertIdMustBeAWholeToken()
    {
        string partial = CrashPrompt.Replace("| 700164", "| 오류 7001645");
        Assert.False(WorkbookOpener.IsCrashReopenPrompt(new[] { Yes, No }, partial, File));
    }
}
