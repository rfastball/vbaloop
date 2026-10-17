using System.Text;
using System.Text.RegularExpressions;

namespace VbaLoop.Excel;

/// <summary>
/// VBE 모달 탐지/해제 프리미티브 (Watch-Modal.ps1 이식, 순수 Win32 — COM 미사용).
/// 배경: 본문 컴파일 오류·미처리 런타임 오류는 트랩 불가 VBE 모달(#32770)로 뜨고,
/// 다이얼로그만 닫으면 Application.Run이 반환되지 않는다 — VBE 창까지 닫아야 제어가 복귀
/// (2026-07-03 실증). VBE 강제 닫기는 dirty 워크북에 NUIDialog("변경 저장?")를 유발하며
/// 이것은 DirectUI라 BM_CLICK 불가, WM_CLOSE(=취소)만 유효 (2026-07-04 실증).
/// 인프로세스 스레드에서 쓰므로 모든 크로스-프로세스 메시지는 행 불가 형태만 사용:
/// 캡션=GetWindowTextW / 자식 텍스트=SendMessageTimeout(500ms) / 클릭=SendMessageTimeout(2s,
/// 타임아웃 시 재클릭 금지) / 닫기=PostMessage.
/// 버튼 선택은 언어 중립(DESIGN §8): 액세스 키·컨트롤 ID·버튼 구조가 1차, 캡션 라벨은
/// 최후 보조 신호. Debug는 어떤 규칙으로도 누르지 않는다(break 모드 → 이후 COM attach 파손).
/// </summary>
public static class ModalHunter
{
    /// <summary>캡션 폴백(최후 보조 신호) 우선순위: 확인/OK(컴파일 오류) → 종료/End(미처리 런타임
    /// 오류) → 부정/취소. 1차 판별은 <see cref="PickSafeButton"/>의 액세스 키·컨트롤 ID 규칙.</summary>
    public static readonly string[] DismissLabels =
        { "확인", "OK", "종료", "End", "아니요", "No", "취소", "Cancel" };

    // 캡션 보조 신호 — 구조 규칙이 실패했을 때만 쓴다. 대소문자 무시, 니모닉 제거 후 비교.
    private static readonly string[] EndLabels =
        { "종료", "End", "終了", "Beenden", "Fin", "Fine", "Finalizar", "Terminar" };
    private static readonly string[] DebugLabels =
        { "디버그", "Debug", "デバッグ", "Debuggen", "Débogage", "Depurar" };
    private static readonly string[] HelpLabels =
        { "도움말", "Help", "ヘルプ", "Hilfe", "Aide", "Aiuto", "Ayuda", "Ajuda" };
    private static readonly string[] NegativeLabels = { "아니요", "No", "취소", "Cancel" };

    public const int IDOK = 1, IDCANCEL = 2, IDNO = 7, IDCLOSE = 8, IDHELP = 9;

    public enum ModalKind { Runtime, Compile, Message, Unknown }

    /// <summary>버튼 1개 스냅샷. <paramref name="Visible"/>은 버튼 자신의 WS_VISIBLE 비트 —
    /// 부모 가시성과 무관하므로 아직 숨은 다이얼로그의 버튼도 true다.</summary>
    public sealed record ButtonInfo(string Text, int CtrlId, bool Enabled, bool Visible = true);

    /// <summary>선택 결과: 스냅샷 인덱스 + 발동 규칙(로그/페이로드 표기용).</summary>
    public sealed record ButtonPick(int Index, string Rule);

    /// <summary>해제 결과. <paramref name="Dismissal"/>은 페이로드 dismissal 문자열
    /// (<see cref="FormatDismissal"/>), Fallback=true면 안전 버튼을 못 찾아 WM_CLOSE한 것.</summary>
    public sealed record DismissResult(
        string Dismissal, ModalKind Kind, string Rule, bool RequiresVbeClose, bool Fallback);

    /// <summary>이 해제 뒤 VBE 창을 닫아야 Application.Run이 반환되는가(컴파일 오류 #32770,
    /// 실측). 런타임 오류의 End는 클릭만으로 제어가 반환되므로 VBE를 닫지 않는다 —
    /// 닫으면 dirty 프로젝트 저장 프롬프트와 XLMAIN 소멸을 유발(2026-07-08 실측).
    /// 입력은 <see cref="FormatDismissal"/> 토큰(<c>ok:</c>/<c>end:</c>/<c>cancel:</c>/<c>wm_close</c>)
    /// 또는 구 형식의 bare 버튼 라벨.</summary>
    public static bool RequiresVbeClose(string? dismissal)
    {
        if (dismissal == null) return false;
        if (dismissal.StartsWith("ok:", StringComparison.Ordinal)) return true;
        if (dismissal.StartsWith("end:", StringComparison.Ordinal)
            || dismissal.StartsWith("cancel:", StringComparison.Ordinal)) return false;
        // 버튼 못 찾아 강제 닫음 → 보수적으로 VBE도 닫음
        if (dismissal.StartsWith("wm_close", StringComparison.OrdinalIgnoreCase)) return true;
        string n = NormalizeButtonText(dismissal);
        return n.Equals("확인", StringComparison.OrdinalIgnoreCase)
               || n.Equals("OK", StringComparison.OrdinalIgnoreCase);
    }

    public const string VbeWindowClass = "wndclass_desked_gsk";

    /// <summary>Visible=false여도 대상이다 — 런타임 오류 모달은 Excel이 포커스를 얻기 전까지
    /// 숨은 채로 있을 수 있다. OwnerChain은 소유자 클래스명을 '&gt;'로 이은 것(분류·로그용).</summary>
    public sealed record DialogInfo(IntPtr Hwnd, string ClassName, string Caption,
        bool Visible = true, string OwnerChain = "")
    {
        public bool OwnedByVbe => OwnerChain.Contains(VbeWindowClass, StringComparison.Ordinal);
    }

    /// <summary>대상 PID 소유의 top-level 다이얼로그(#32770=VBE 모달, NUIDialog=저장 프롬프트).
    /// EnumWindows는 숨은 창도 열거하므로 가시성으로 거르지 않는다.</summary>
    public static List<DialogInfo> FindDialogs(int pid)
    {
        var found = new List<DialogInfo>();
        NativeMethods.EnumWindows((h, _) =>
        {
            NativeMethods.GetWindowThreadProcessId(h, out uint wpid);
            if (wpid == (uint)pid)
            {
                string cls = NativeMethods.GetWindowClassName(h);
                if (cls is "#32770" or "NUIDialog")
                    found.Add(new DialogInfo(h, cls, GetCaption(h),
                        NativeMethods.IsWindowVisible(h), GetOwnerChain(h)));
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static string GetCaption(IntPtr hwnd)
    {
        var sb = new StringBuilder(512);
        _ = NativeMethods.GetWindowTextW(hwnd, sb, sb.Capacity);
        return sb.ToString();
    }

    /// <summary>"Microsoft Visual Basic"은 제품명이라 UI 언어와 무관하게 지역화되지 않는다
    /// (런타임 오류="Microsoft Visual Basic", 컴파일 오류="... for Applications").</summary>
    public static bool IsVbaModalCaption(string caption)
        => caption.StartsWith("Microsoft Visual Basic", StringComparison.Ordinal);

    /// <summary>다이얼로그 캡션 + 자식 컨트롤 텍스트를 한 줄로 수집(오류 본문 캡처).</summary>
    public static string CollectInfo(IntPtr dlg)
    {
        var acc = new StringBuilder();
        acc.Append(GetCaption(dlg));
        NativeMethods.EnumChildWindows(dlg, (h, _) =>
        {
            string t = GetTextSafe(h);
            if (t.Length > 0)
            {
                acc.Append(" | ");
                acc.Append(t);
            }
            return true;
        }, IntPtr.Zero);
        return acc.ToString().Replace("\r", " ").Replace("\n", " ");
    }

    /// <summary>다이얼로그의 푸시 버튼 스냅샷(체크박스/라디오/그룹박스 제외). 자식 Z-순서 =
    /// 템플릿 생성 순서.</summary>
    public static List<(IntPtr Hwnd, ButtonInfo Info)> SnapshotButtons(IntPtr dlg)
    {
        var buttons = new List<(IntPtr, ButtonInfo)>();
        NativeMethods.EnumChildWindows(dlg, (h, _) =>
        {
            if (NativeMethods.GetWindowClassName(h) == "Button")
            {
                int style = NativeMethods.GetWindowLong(h, NativeMethods.GWL_STYLE);
                int type = style & NativeMethods.BS_TYPEMASK;
                if (type is NativeMethods.BS_PUSHBUTTON or NativeMethods.BS_DEFPUSHBUTTON)
                    buttons.Add((h, new ButtonInfo(
                        GetTextSafe(h).Trim(),
                        NativeMethods.GetDlgCtrlID(h),
                        NativeMethods.IsWindowEnabled(h),
                        (style & NativeMethods.WS_VISIBLE) != 0)));
            }
            return true;
        }, IntPtr.Zero);
        return buttons;
    }

    /// <summary>
    /// 다이얼로그를 언어 중립 규칙으로 해제한다. 안전 버튼이 있으면 누르고, 없으면
    /// <paramref name="allowFallback"/>일 때만 WM_CLOSE(=dismissal "wm_close [...]",
    /// RequiresVbeClose=true). 폴백이 금지됐고 버튼도 없으면 null(호출부가 다음 폴에서 재시도 —
    /// 생성 중이라 버튼이 아직 없는 다이얼로그를 성급히 닫지 않기 위함).
    /// </summary>
    public static DismissResult? Dismiss(DialogInfo dlg, bool allowFallback = true)
    {
        var buttons = SnapshotButtons(dlg.Hwnd);
        var infos = buttons.Select(b => b.Info).ToList();
        bool vbaContext = IsVbaModalCaption(dlg.Caption) || dlg.OwnedByVbe;
        var kind = Classify(infos, vbaContext);
        var pick = PickSafeButton(infos, kind);
        bool hidden = !dlg.Visible;

        if (pick != null)
        {
            var (hwnd, info) = buttons[pick.Index];
            Press(dlg, hwnd, info.CtrlId);
            string action = ActionFor(kind, info);
            return new DismissResult(FormatDismissal(action, info.Text, kind, pick.Rule, hidden),
                kind, pick.Rule, action == "ok", Fallback: false);
        }

        if (!allowFallback) return null;
        CloseWindow(dlg.Hwnd);
        return new DismissResult(FormatDismissal("wm_close", null, kind, "no_safe_button", hidden),
            kind, "no_safe_button", RequiresVbeClose: true, Fallback: true);
    }

    // ── 언어 중립 판별 (순수 로직, 단위 테스트 대상) ──

    /// <summary>
    /// 구조로 분류한다(본문 텍스트 미사용). Runtime = VBA 문맥 + 푸시 버튼 3개 이상 + 종료/디버그
    /// 신호(액세스 키 E·D, 보조로 라벨). Compile = 그 밖의 VBA 문맥(캡션 "Microsoft Visual Basic*"
    /// 또는 VBE 소유). Message = VBA 문맥이 아닌 일반 #32770. 버튼이 없으면 Unknown.
    /// </summary>
    public static ModalKind Classify(IReadOnlyList<ButtonInfo> buttons, bool vbaContext)
    {
        if (buttons.Count == 0) return ModalKind.Unknown;
        if (vbaContext && buttons.Count >= 3 && buttons.Any(b =>
                AccessKey(b.Text) is 'E' or 'D'
                || InLabels(b.Text, EndLabels) || InLabels(b.Text, DebugLabels)))
            return ModalKind.Runtime;
        return vbaContext ? ModalKind.Compile : ModalKind.Message;
    }

    /// <summary>
    /// 누를 안전 버튼을 고른다. 공통 제외: 비활성, 숨김(WS_VISIBLE 없음), Debug(액세스 키 D
    /// 또는 디버그 라벨).
    /// 도움말 = IDHELP 또는(보조) 도움말 라벨 — 단일 버튼 규칙에서 빠진다.
    /// Runtime: ① 액세스 키 E → ② 도움말 제외 후 단 하나 남은 버튼 → ③ 종료 라벨.
    /// Compile/Message: ① 컨트롤 ID IDOK → IDCLOSE → ② 도움말 제외 단일 버튼 → ③ 캡션 라벨
    /// (<see cref="DismissLabels"/> 순서). Unknown이거나 해당 없음 → null(호출부 폴백).
    /// </summary>
    public static ButtonPick? PickSafeButton(IReadOnlyList<ButtonInfo> buttons, ModalKind kind)
    {
        bool Usable(ButtonInfo b) => b.Enabled && b.Visible && !IsDebug(b);
        var usable = Enumerable.Range(0, buttons.Count).Where(i => Usable(buttons[i])).ToList();
        var notHelp = usable.Where(i => !IsHelp(buttons[i])).ToList();

        switch (kind)
        {
            case ModalKind.Runtime:
                foreach (int i in usable)
                    if (AccessKey(buttons[i].Text) == 'E') return new ButtonPick(i, "access_key_e");
                if (notHelp.Count == 1) return new ButtonPick(notHelp[0], "single_remaining");
                foreach (int i in notHelp)
                    if (InLabels(buttons[i].Text, EndLabels)) return new ButtonPick(i, "caption_label");
                return null;

            case ModalKind.Compile:
            case ModalKind.Message:
                foreach (int id in new[] { IDOK, IDCLOSE })
                    foreach (int i in usable)
                        if (buttons[i].CtrlId == id) return new ButtonPick(i, "control_id");
                if (notHelp.Count == 1) return new ButtonPick(notHelp[0], "single_button");
                foreach (var want in DismissLabels)
                    foreach (int i in notHelp)
                        if (LabelEquals(buttons[i].Text, want)) return new ButtonPick(i, "caption_label");
                return null;

            default:
                return null;
        }
    }

    /// <summary>답이 하나뿐인 알림(활성 버튼이 도움말 빼고 정확히 1개) — 자동 해제해도 오답이
    /// 있을 수 없는 다이얼로그. 가드 윈도의 "Microsoft Excel" #32770 처리 조건(본문 텍스트 대체).</summary>
    public static bool IsSingleAnswer(IReadOnlyList<ButtonInfo> buttons)
        => buttons.Count(b => b.Enabled && b.Visible && !IsHelp(b)) == 1;

    /// <summary>버튼 텍스트의 액세스 키(대문자). "&amp;End"→E, "종료(&amp;E)"→E, "&amp;&amp;"는 리터럴.
    /// '&amp;'가 없으면 꼬리 "(X)" 표기를 본다. 없으면 null.</summary>
    public static char? AccessKey(string? text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        for (int i = 0; i < text.Length - 1; i++)
        {
            if (text[i] != '&') continue;
            if (text[i + 1] == '&') { i++; continue; }
            return char.ToUpperInvariant(text[i + 1]);
        }
        var m = Regex.Match(text, @"\(([A-Za-z0-9])\)\s*$");
        return m.Success ? char.ToUpperInvariant(m.Groups[1].Value[0]) : null;
    }

    /// <summary>페이로드 dismissal 문자열: <c>{action}:{버튼 원문} [{kind}/{rule}(/hidden)]</c>,
    /// 폴백은 <c>wm_close [{kind}/no_safe_button(/hidden)]</c>. action ∈ ok|end|cancel.</summary>
    public static string FormatDismissal(string action, string? buttonText, ModalKind kind, string rule, bool hidden)
    {
        string tag = $"[{kind.ToString().ToLowerInvariant()}/{rule}{(hidden ? "/hidden" : "")}]";
        return action == "wm_close" ? $"wm_close {tag}" : $"{action}:{buttonText} {tag}";
    }

    /// <summary>눌린 버튼의 의미. Compile은 거절 선택지가 없는 확인 전용 모달이라 버튼 ID와 무관하게
    /// "ok"다 — 실측(한국어 VBE)에서 확인 버튼이 IDCANCEL이어서 "cancel"로 오분류되면 VBE를 닫지 않아
    /// Application.Run이 영영 미반환했다(러너 120s 타임아웃 회귀).</summary>
    public static string ActionFor(ModalKind kind, ButtonInfo b)
    {
        if (kind == ModalKind.Runtime) return "end";
        if (kind == ModalKind.Compile) return "ok";
        bool negative = b.CtrlId is IDCANCEL or IDNO || InLabels(b.Text, NegativeLabels);
        return negative ? "cancel" : "ok";
    }

    private static bool IsDebug(ButtonInfo b)
        => AccessKey(b.Text) == 'D' || InLabels(b.Text, DebugLabels);

    private static bool IsHelp(ButtonInfo b)
        => b.CtrlId == IDHELP || InLabels(b.Text, HelpLabels);

    private static bool InLabels(string text, string[] labels)
        => labels.Any(l => LabelEquals(text, l));

    private static bool LabelEquals(string text, string label)
        => string.Equals(NormalizeButtonText(text), NormalizeButtonText(label), StringComparison.OrdinalIgnoreCase);

    /// <summary>"&amp;확인(Y)" → "확인" — 니모닉 앰퍼샌드와 꼬리 단축키 표기 제거.</summary>
    public static string NormalizeButtonText(string? text)
    {
        if (text == null) return "";
        string t = text.Replace("&", "").Trim();
        t = Regex.Replace(t, @"\s*\([A-Za-z0-9]\)\s*$", "");
        return t.Trim();
    }

    /// <summary>
    /// 버튼 누르기. 보이는 다이얼로그는 실증된 BM_CLICK, 숨은 다이얼로그는 WM_COMMAND(ID,
    /// BN_CLICKED)를 다이얼로그에 직접 — BM_CLICK은 비활성(포커스 없는) 다이얼로그에서 실패할 수
    /// 있다고 문서화되어 있다. 타임아웃돼도 sent-message는 대상 큐에 남아 나중에 처리됨 —
    /// 재클릭 금지, 다음 폴에서 창 잔존 여부로 재판단한다.
    /// </summary>
    private static void Press(DialogInfo dlg, IntPtr button, int ctrlId)
    {
        const uint flags = NativeMethods.SMTO_ABORTIFHUNG | NativeMethods.SMTO_NORMAL;
        if (dlg.Visible || ctrlId == 0)
            _ = NativeMethods.SendMessageTimeoutW(button, NativeMethods.BM_CLICK, IntPtr.Zero, IntPtr.Zero,
                flags, 2000, out _);
        else
            _ = NativeMethods.SendMessageTimeoutW(dlg.Hwnd, NativeMethods.WM_COMMAND,
                (IntPtr)(ctrlId & 0xFFFF), button, flags, 2000, out _); // HIWORD=BN_CLICKED(0)
    }

    /// <summary>WM_CLOSE 전송(=X 클릭/취소). PostMessage라 블록 불가.</summary>
    public static void CloseWindow(IntPtr hwnd)
        => NativeMethods.PostMessageW(hwnd, NativeMethods.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);

    /// <summary>VBE 메인 윈도우 전부에 WM_CLOSE — 이게 없으면 Application.Run이 영영 미반환.</summary>
    public static int CloseVbeWindows(int pid)
    {
        var vbe = new List<IntPtr>();
        NativeMethods.EnumWindows((h, _) =>
        {
            NativeMethods.GetWindowThreadProcessId(h, out uint wpid);
            if (wpid == (uint)pid && NativeMethods.GetWindowClassName(h) == VbeWindowClass)
                vbe.Add(h);
            return true;
        }, IntPtr.Zero);
        foreach (var h in vbe)
            NativeMethods.PostMessageW(h, NativeMethods.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        return vbe.Count;
    }

    public static bool WaitUntilGone(IntPtr hwnd, int timeoutMs)
    {
        var deadline = Environment.TickCount64 + timeoutMs;
        while (Environment.TickCount64 < deadline)
        {
            if (!NativeMethods.IsWindow(hwnd)) return true;
            Thread.Sleep(50);
        }
        return !NativeMethods.IsWindow(hwnd);
    }

    /// <summary>
    /// 1회성 스윕(부모 post-mortem용): 잔존 VBE 모달을 해제하고 VBE를 닫는다.
    /// 러너가 모달을 남기고 비정상 종료하면 이후 모든 명령이 COM busy로 막히는 갭을 봉합.
    /// 처리한 모달 설명 목록 반환.
    /// </summary>
    public static List<string> SweepOnce(int pid)
    {
        var handled = new List<string>();
        bool needVbeClose = false;
        foreach (var dlg in FindDialogs(pid))
        {
            if (!IsVbaModalCaption(dlg.Caption)) continue;
            string info = CollectInfo(dlg.Hwnd);
            var r = Dismiss(dlg)!; // 폴백 허용 → 항상 non-null
            WaitUntilGone(dlg.Hwnd, 1500);
            if (r.RequiresVbeClose) needVbeClose = true;
            handled.Add($"{r.Dismissal}: {info}");
        }
        if (needVbeClose)
        {
            Thread.Sleep(150);
            CloseVbeWindows(pid);
        }
        // VBE 닫기 후속 저장 프롬프트가 세션을 영구 차단하므로 항상 능동 드레인.
        // (모달을 못 봤어도 잔존 save 프롬프트가 있을 수 있음 — 부모 안전망 겸용)
        handled.AddRange(DrainSavePrompts(pid, maxMs: 5000));
        return handled;
    }

    /// <summary>
    /// "변경 저장?" NUIDialog(DirectUI, WM_CLOSE=취소)를 안정될 때까지 드레인.
    /// 프롬프트가 없는 상태가 1.2s 유지되거나 maxMs 초과 시 종료.
    /// </summary>
    public static List<string> DrainSavePrompts(int pid, int maxMs)
    {
        var handled = new List<string>();
        long start = Environment.TickCount64;
        long lastActivity = start;
        while (Environment.TickCount64 - start < maxMs)
        {
            var prompt = FindDialogs(pid).FirstOrDefault(
                d => d.Caption == "Microsoft Excel" && d.ClassName == "NUIDialog");
            if (prompt != null)
            {
                CloseWindow(prompt.Hwnd);
                handled.Add("save_prompt: closed");
                lastActivity = Environment.TickCount64;
            }
            else if (Environment.TickCount64 - lastActivity > 1200)
            {
                break;
            }
            Thread.Sleep(100);
        }
        return handled;
    }

    /// <summary>소유자 체인의 클래스명(최대 4단계, '&gt;' 구분) — GetWindow(GW_OWNER)는 메시지를
    /// 보내지 않아 행 불가.</summary>
    private static string GetOwnerChain(IntPtr hwnd)
    {
        var parts = new List<string>();
        IntPtr h = NativeMethods.GetWindow(hwnd, NativeMethods.GW_OWNER);
        for (int depth = 0; depth < 4 && h != IntPtr.Zero; depth++)
        {
            parts.Add(NativeMethods.GetWindowClassName(h));
            h = NativeMethods.GetWindow(h, NativeMethods.GW_OWNER);
        }
        return string.Join(">", parts);
    }

    private static string GetTextSafe(IntPtr h)
    {
        var sb = new StringBuilder(4096);
        var ok = NativeMethods.SendMessageTimeoutW(h, NativeMethods.WM_GETTEXT, (IntPtr)sb.Capacity, sb,
            NativeMethods.SMTO_ABORTIFHUNG | NativeMethods.SMTO_NORMAL, 500, out _);
        return ok == IntPtr.Zero ? "" : sb.ToString();
    }
}
