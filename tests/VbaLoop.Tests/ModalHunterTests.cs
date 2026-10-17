using VbaLoop.Excel;
using Xunit;
using static VbaLoop.Excel.ModalHunter;

namespace VbaLoop.Tests;

public class ModalHunterTests
{
    [Theory]
    [InlineData("&확인", "확인")]
    [InlineData("확인(Y)", "확인")]
    [InlineData("&아니요(N)", "아니요")]
    [InlineData("  OK  ", "OK")]
    [InlineData("종료(&E)", "종료")] // 앰퍼샌드 제거 후 (E) 꼬리 제거
    [InlineData("", "")]
    [InlineData(null, "")]
    public void NormalizeButtonText_StripsMnemonics(string? raw, string expected)
    {
        Assert.Equal(expected, ModalHunter.NormalizeButtonText(raw));
    }

    [Fact]
    public void PickSafeButton_LegacyLabelCases_WithUnknownIds()
    {
        // 구 캡션 매칭 시절 사례가 ID 미상(1000번대)에서도 같은 버튼으로 수렴하는지
        // 컴파일 오류 모달: [확인, 도움말] → 확인
        AssertPick(new[] { B("확인"), B("도움말(&H)") }, vbaContext: true, ModalKind.Compile, 0, "single_button");
        // 미처리 런타임 오류(액세스 키 없음): [Continue, End, Debug, Help] → End (Continue가 활성이라 라벨까지)
        AssertPick(new[] { B("Continue"), B("End"), B("Debug"), B("Help") },
            vbaContext: true, ModalKind.Runtime, 1, "caption_label");
        // 둘 다 확인류면 DismissLabels 순서(확인 우선)
        AssertPick(new[] { B("OK"), B("확인") }, vbaContext: true, ModalKind.Compile, 1, "caption_label");
        // 디버그·도움말뿐 → 고르지 않음 (호출부가 WM_CLOSE 폴백)
        Assert.Null(PickSafeButton(new[] { B("디버그"), B("도움말") }, ModalKind.Compile));
    }

    [Fact]
    public void IsVbaModalCaption_MatchesVbePrefix()
    {
        Assert.True(ModalHunter.IsVbaModalCaption("Microsoft Visual Basic"));
        Assert.True(ModalHunter.IsVbaModalCaption("Microsoft Visual Basic for Applications"));
        Assert.False(ModalHunter.IsVbaModalCaption("Microsoft Excel"));
    }

    [Theory]
    [InlineData("확인", true)]       // 컴파일 오류 OK → VBE 닫아야 제어 반환
    [InlineData("OK", true)]
    [InlineData("종료(&E)", false)]  // 런타임 오류 End → 클릭만으로 반환, VBE 닫기 금지
    [InlineData("End", false)]
    [InlineData("wm_close", true)]   // 버튼 못 찾음 → 보수적으로 VBE 닫음
    // L3 토큰 형식: action이 결정한다(버튼 언어 무관)
    [InlineData("ok:확인 [compile/control_id]", true)]
    [InlineData("ok:OK [compile/single_button/hidden]", true)]
    [InlineData("end:&Beenden [runtime/caption_label]", false)]
    [InlineData("end:종료(&E) [runtime/access_key_e]", false)]
    [InlineData("cancel:Cancel [compile/caption_label]", false)]
    [InlineData("wm_close [unknown/no_safe_button]", true)]
    [InlineData(null, false)]
    public void RequiresVbeClose_OnlyForCompileErrorDismissal(string? label, bool expected)
    {
        Assert.Equal(expected, ModalHunter.RequiresVbeClose(label));
    }

    // ── L3 언어 중립 판별: 합성 버튼 스냅샷 ──
    // 컨트롤 ID 1000번대는 "알 수 없는 ID"(IDOK/IDCLOSE/IDHELP 아님)를 뜻한다.

    private static ButtonInfo B(string text, int id = 1000, bool enabled = true, bool visible = true)
        => new(text, id, enabled, visible);

    private static void AssertPick(ButtonInfo[] buttons, bool vbaContext,
        ModalKind expectedKind, int expectedIndex, string expectedRule)
    {
        var kind = Classify(buttons, vbaContext);
        Assert.Equal(expectedKind, kind);
        var pick = PickSafeButton(buttons, kind);
        Assert.NotNull(pick);
        Assert.Equal(expectedIndex, pick!.Index);
        Assert.Equal(expectedRule, pick.Rule);
    }

    [Fact]
    public void Runtime_Korean_ScreenshotButtons_PicksEndByAccessKey()
    {
        // 실제 화면: 계속(C) 비활성, 종료(E), 디버그(D), 도움말(H) — 표시상 '&'는 숨는다
        AssertPick(new[]
        {
            B("계속(&C)", 1001, enabled: false), B("종료(&E)", 1002), B("디버그(&D)", 1003), B("도움말(&H)", 1004),
        }, vbaContext: true, ModalKind.Runtime, 1, "access_key_e");
    }

    [Fact]
    public void Runtime_KoreanWithoutAmpersand_UsesTrailingParenAccessKey()
    {
        AssertPick(new[] { B("계속(C)", enabled: false), B("종료(E)"), B("디버그(D)"), B("도움말(H)") },
            vbaContext: true, ModalKind.Runtime, 1, "access_key_e");
    }

    [Fact]
    public void Runtime_English_PicksEndByAccessKey()
    {
        AssertPick(new[] { B("&Continue", enabled: false), B("&End"), B("&Debug"), B("&Help") },
            vbaContext: true, ModalKind.Runtime, 1, "access_key_e");
    }

    [Fact]
    public void Runtime_Japanese_PicksEndByAccessKey()
    {
        AssertPick(new[] { B("継続(&C)", enabled: false), B("終了(&E)"), B("デバッグ(&D)"), B("ヘルプ(&H)") },
            vbaContext: true, ModalKind.Runtime, 1, "access_key_e");
    }

    [Fact]
    public void Runtime_German_HelpById_PicksSingleRemaining()
    {
        // &Beenden은 액세스 키가 B라 E 규칙 불발. 비활성·Debug(D)·IDHELP 제외 후 하나만 남는다
        AssertPick(new[] { B("&Fortfahren", enabled: false), B("&Beenden"), B("&Debuggen"), B("&Hilfe", IDHELP) },
            vbaContext: true, ModalKind.Runtime, 1, "single_remaining");
    }

    [Fact]
    public void Runtime_German_UnknownIds_HelpByLabel_PicksSingleRemaining()
    {
        // IDHELP가 아니어도 도움말 라벨(보조 신호)로 제외되어 하나만 남는다
        AssertPick(new[] { B("&Fortfahren", enabled: false), B("&Beenden"), B("&Debuggen"), B("&Hilfe") },
            vbaContext: true, ModalKind.Runtime, 1, "single_remaining");
    }

    [Fact]
    public void Runtime_French_PicksSingleRemaining()
    {
        AssertPick(new[] { B("&Continuer", enabled: false), B("&Fin"), B("&Débogage"), B("&Aide") },
            vbaContext: true, ModalKind.Runtime, 1, "single_remaining");
    }

    [Fact]
    public void Runtime_ContinueEnabled_FallsBackToCaptionLabel()
    {
        // Continue까지 활성이면(Stop 문 등) 구조로는 둘 중 하나를 못 고른다 → 종료 라벨(최후)
        AssertPick(new[] { B("&Continuer"), B("&Fin"), B("&Débogage"), B("&Aide") },
            vbaContext: true, ModalKind.Runtime, 1, "caption_label");
    }

    [Fact]
    public void Runtime_NeverPressesDebug()
    {
        // End가 비활성이고 남은 게 Debug와 도움말뿐 → 고르지 않는다(호출부 WM_CLOSE 폴백)
        var buttons = new[] { B("&Continue", enabled: false), B("&End", enabled: false), B("&Debug"), B("&Help", IDHELP) };
        Assert.Equal(ModalKind.Runtime, Classify(buttons, vbaContext: true));
        Assert.Null(PickSafeButton(buttons, ModalKind.Runtime));

        // 액세스 키 없는 디버그 라벨도 제외
        var noKeys = new[] { B("Continue", enabled: false), B("End", enabled: false), B("Debug"), B("Help", IDHELP) };
        Assert.Null(PickSafeButton(noKeys, ModalKind.Runtime));
    }

    [Fact]
    public void Runtime_DisabledEndButton_IsNotPicked()
    {
        var buttons = new[] { B("&Continue"), B("&End", enabled: false), B("&Debug"), B("&Help", IDHELP) };
        // E는 비활성, 남은 후보는 Continue 하나 — 구조 규칙상 유일 후보(Debug 아님)
        var pick = PickSafeButton(buttons, ModalKind.Runtime);
        Assert.Equal(0, pick!.Index);
        Assert.Equal("single_remaining", pick.Rule);
    }

    [Fact]
    public void Compile_Korean_ScreenshotButtons_ByControlIdOrCaption()
    {
        // 컨트롤 ID가 표준이면 IDOK
        AssertPick(new[] { B("확인", IDOK), B("도움말", IDHELP) },
            vbaContext: true, ModalKind.Compile, 0, "control_id");
        // ID를 모를 때: 도움말이 IDHELP면 단일 버튼
        AssertPick(new[] { B("확인"), B("도움말", IDHELP) },
            vbaContext: true, ModalKind.Compile, 0, "single_button");
        // ID를 전혀 모를 때: 도움말 라벨(보조)로 제외 후 단일 버튼
        AssertPick(new[] { B("확인"), B("도움말") },
            vbaContext: true, ModalKind.Compile, 0, "single_button");
    }

    [Fact]
    public void Compile_OkButtonWithCancelId_StillClosesVbe()
    {
        // 실측: 한국어 VBE 컴파일 오류의 "확인"은 IDCANCEL이었다. 그래도 동작은 ok여야
        // RequiresVbeClose가 참이 되어 Application.Run이 반환된다.
        var ok = B("확인", IDCANCEL);
        Assert.Equal("ok", ActionFor(ModalKind.Compile, ok));
        Assert.True(RequiresVbeClose($"{ActionFor(ModalKind.Compile, ok)}:확인 [compile/single_button]"));
        // 일반 메시지 대화상자의 취소는 여전히 cancel
        Assert.Equal("cancel", ActionFor(ModalKind.Message, B("취소", IDCANCEL)));
    }

    [Fact]
    public void Compile_ControlIdOnly_IgnoresText()
    {
        // 텍스트가 어떤 언어든(여기선 판독 불가 문자열) ID만으로 결정
        AssertPick(new[] { B("???", IDHELP), B("!!!", IDOK) },
            vbaContext: true, ModalKind.Compile, 1, "control_id");
        AssertPick(new[] { B("???", IDHELP), B("!!!", IDCLOSE) },
            vbaContext: true, ModalKind.Compile, 1, "control_id");
        // IDOK가 IDCLOSE보다 우선
        AssertPick(new[] { B("a", IDCLOSE), B("b", IDOK) },
            vbaContext: true, ModalKind.Compile, 1, "control_id");
    }

    [Fact]
    public void Compile_DisabledOrHiddenOk_FallsToNextRule()
    {
        AssertPick(new[] { B("OK", IDOK, enabled: false), B("Schließen", IDCLOSE), B("Hilfe", IDHELP) },
            vbaContext: true, ModalKind.Compile, 1, "control_id");
        // 숨은 버튼(WS_VISIBLE 없음)도 후보 아님
        AssertPick(new[] { B("OK", IDOK, visible: false), B("Fermer"), B("Aide", IDHELP) },
            vbaContext: true, ModalKind.Compile, 1, "single_button");
    }

    [Fact]
    public void Compile_HelpButton_IsExcludedFromSingleButtonRule()
    {
        AssertPick(new[] { B("Aide", IDHELP), B("Valider") },
            vbaContext: true, ModalKind.Compile, 1, "single_button");
    }

    [Fact]
    public void Compile_NothingSafe_ReturnsNull()
    {
        var buttons = new[] { B("foo"), B("bar") }; // ID·라벨 모두 미상, 후보 2개
        Assert.Equal(ModalKind.Compile, Classify(buttons, vbaContext: true));
        Assert.Null(PickSafeButton(buttons, ModalKind.Compile));
    }

    [Fact]
    public void Classify_IsStructural()
    {
        Assert.Equal(ModalKind.Unknown, Classify(Array.Empty<ButtonInfo>(), vbaContext: true));
        // 2버튼 VBA 모달 = 컴파일 오류
        Assert.Equal(ModalKind.Compile, Classify(new[] { B("OK", IDOK), B("Help", IDHELP) }, vbaContext: true));
        // VBA 문맥이 아니면 3버튼이어도 런타임 아님
        Assert.Equal(ModalKind.Message,
            Classify(new[] { B("&Yes", 6), B("&No", 7), B("Cancel", 2) }, vbaContext: false));
        Assert.Equal(ModalKind.Message,
            Classify(new[] { B("&End"), B("&Debug"), B("&Help") }, vbaContext: false));
        // VBA 3버튼이라도 종료/디버그 신호가 없으면 런타임 아님
        Assert.Equal(ModalKind.Compile,
            Classify(new[] { B("&Yes", 6), B("&No", 7), B("Cancel", 2) }, vbaContext: true));
    }

    [Theory]
    [InlineData("&End", 'E')]
    [InlineData("종료(&E)", 'E')]
    [InlineData("終了(&E)", 'E')]
    [InlineData("&Beenden", 'B')]
    [InlineData("&fin", 'F')]
    [InlineData("종료(E)", 'E')]   // '&' 없는 꼬리 표기
    [InlineData("R&&D", null)]     // '&&'는 리터럴
    [InlineData("R&&D &Go", 'G')]
    [InlineData("OK", null)]
    [InlineData("", null)]
    [InlineData(null, null)]
    public void AccessKey_ParsesMnemonic(string? text, char? expected)
    {
        Assert.Equal(expected, AccessKey(text));
    }

    [Fact]
    public void IsSingleAnswer_OnlyForOneEnabledNonHelpButton()
    {
        Assert.True(IsSingleAnswer(new[] { B("OK", IDOK) }));
        Assert.True(IsSingleAnswer(new[] { B("OK", IDOK), B("Help", IDHELP) }));
        Assert.True(IsSingleAnswer(new[] { B("확인"), B("도움말") })); // 도움말은 라벨 보조로 제외
        Assert.True(IsSingleAnswer(new[] { B("OK", IDOK), B("Cancel", 2, enabled: false) }));
        Assert.False(IsSingleAnswer(new[] { B("&Yes", 6), B("&No", 7) }));
        Assert.False(IsSingleAnswer(new[] { B("OK", IDOK, enabled: false) }));
        Assert.False(IsSingleAnswer(Array.Empty<ButtonInfo>()));
    }

    [Fact]
    public void FormatDismissal_RoundTripsThroughRequiresVbeClose()
    {
        string end = FormatDismissal("end", "종료(&E)", ModalKind.Runtime, "access_key_e", hidden: true);
        Assert.Equal("end:종료(&E) [runtime/access_key_e/hidden]", end);
        Assert.False(RequiresVbeClose(end));

        string ok = FormatDismissal("ok", "확인", ModalKind.Compile, "control_id", hidden: false);
        Assert.Equal("ok:확인 [compile/control_id]", ok);
        Assert.True(RequiresVbeClose(ok));

        string fb = FormatDismissal("wm_close", null, ModalKind.Runtime, "no_safe_button", hidden: false);
        Assert.Equal("wm_close [runtime/no_safe_button]", fb);
        Assert.True(RequiresVbeClose(fb));
    }

    [Theory]
    [InlineData(1000, 1000, 0, 1500, false)]    // 방금 시작
    [InlineData(2500, 1000, 0, 1500, true)]     // 모달 없이 1.5s 경과
    [InlineData(2500, 1000, 2000, 1500, false)] // 0.5s 전 모달 관측 → 아직 대기
    [InlineData(3500, 1000, 2000, 1500, true)]
    [InlineData(1200, 1000, 900, 1500, false)]  // 시작 전 관측은 시작 시각 기준
    public void SettleAfterCompile_QuietWindow(long now, long start, long lastSeen, int quietMs, bool expected)
    {
        Assert.Equal(expected, ModalWatchdog.IsQuiet(now, start, lastSeen, quietMs));
    }
}
