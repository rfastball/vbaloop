using System.Collections.Concurrent;

namespace VbaLoop.Excel;

/// <summary>
/// VBE 모달 자동 해제 워치독 — 러너 프로세스 내 백그라운드 스레드(250ms 폴링).
/// _VbaHarness는 별도 PowerShell 프로세스였지만(블로킹된 런스페이스 때문),
/// C#에서는 블로킹된 STA 메인 스레드와 별개의 스레드 하나로 충분하다(Win32만 사용, COM 금지).
/// 해제 시퀀스(전부 실증 기반, 순서 변경 금지): 모달 감지 → 버튼 클릭/WM_CLOSE →
/// 소멸 대기 1.5s → VBE 창 닫기(이게 없으면 Run 미반환) → 3초 가드 윈도에서 후속
/// NUIDialog(저장 프롬프트, WM_CLOSE=취소)와 답이 하나뿐인 "Microsoft Excel" #32770 알림 정리.
/// 버튼 선택은 언어 중립: <see cref="ModalHunter.Dismiss"/>가 액세스 키·컨트롤 ID로 고르고,
/// 안전 버튼을 못 찾은 WM_CLOSE 폴백은 다이얼로그가 1s 이상 관측된 뒤에만(생성 중 오판 방지).
/// 스레드 간 공유는 락 없이: volatile 라벨(main→wd), ConcurrentQueue(wd→main).
/// 이 클래스는 stdout과 jsonl에 절대 쓰지 않는다 — 러너 JSON 프로토콜과
/// offset 기반 A레코드 판독의 단일 기록자 전제를 보호.
/// </summary>
public sealed class ModalWatchdog : IDisposable
{
    public sealed record ModalEvent(
        string Ts, string? Label, string Kind, string Caption, string Body, string Dismissal);

    private readonly int _pid;
    private readonly Thread _thread;
    private readonly ConcurrentQueue<ModalEvent> _events = new();
    private volatile string? _currentLabel;
    private volatile bool _stop;
    private long _lastDismissTick; // Environment.TickCount64, Interlocked 접근
    private long _lastVbaModalTick; // 마지막으로 VBA 모달을 관측한 시각, Interlocked 접근

    /// <summary>안전 버튼이 없을 때 WM_CLOSE 폴백 전에 다이얼로그를 관측해야 하는 최소 시간 —
    /// 숨은 상태로 생성 중인 다이얼로그는 버튼이 아직 없을 수 있다.</summary>
    public const int FallbackGraceMs = 1000;

    public ModalWatchdog(int excelPid)
    {
        _pid = excelPid;
        _thread = new Thread(Loop) { IsBackground = true, Name = "vbaloop-modal-watchdog" };
        _thread.Start();
    }

    /// <summary>실행 중인 테스트/매크로 라벨 — 모달 이벤트 귀속용.</summary>
    public void SetLabel(string? label) => _currentLabel = label;

    public List<ModalEvent> Drain()
    {
        var list = new List<ModalEvent>();
        while (_events.TryDequeue(out var e)) list.Add(e);
        return list;
    }

    /// <summary>
    /// 러너 종료 직전 링거: VBE 닫기 후속 "변경 저장?" NUIDialog를 능동적으로 드레인한다.
    /// 3초 가드 폴링만으로는 프롬프트가 가드 경계 근처에 뜨면 놓쳐 COM이 영구 차단되는
    /// 것을 실측(2026-07-08) — 프롬프트가 안정적으로 사라질 때까지(또는 8s 상한) 직접 닫는다.
    /// 해제 이력이 없으면 즉시 반환.
    /// </summary>
    public void LingerForGuardWindow()
    {
        long last = Interlocked.Read(ref _lastDismissTick);
        if (last == 0) return;

        long start = Environment.TickCount64;
        long lastActivity = start;
        while (Environment.TickCount64 - start < 8000 && !_stop)
        {
            var prompt = ModalHunter.FindDialogs(_pid).FirstOrDefault(
                d => d.Caption == "Microsoft Excel" && d.ClassName == "NUIDialog");
            if (prompt != null)
            {
                ModalHunter.CloseWindow(prompt.Hwnd); // WM_CLOSE = 취소(미저장 유지)
                Enqueue("save_prompt", prompt.Caption, "", "wm_close");
                lastActivity = Environment.TickCount64;
            }
            else if (Environment.TickCount64 - lastActivity > 1200)
            {
                break; // 프롬프트 없이 안정 — 종료 안전
            }
            Thread.Sleep(100);
        }
    }

    /// <summary>
    /// VBE 컴파일 호출 반환 직후 호출: 컴파일 오류 모달은 호출이 반환된 뒤 늦게 뜰 수 있으므로
    /// VBA 모달이 <paramref name="quietMs"/> 동안 관측되지 않을 때까지(상한 <paramref name="maxMs"/>)
    /// 블록한다 — 그사이 뜬 모달은 워치독 스레드가 평소대로 해제하고 Drain()에 남긴다.
    /// 조용해지면 true, 상한 초과(모달이 계속 뜸)면 false.
    /// </summary>
    public bool SettleAfterCompile(int quietMs = 1500, int maxMs = 10000)
    {
        long start = Environment.TickCount64;
        while (!_stop)
        {
            long now = Environment.TickCount64;
            if (IsQuiet(now, start, Interlocked.Read(ref _lastVbaModalTick), quietMs)) return true;
            if (now - start >= maxMs) return false;
            Thread.Sleep(50);
        }
        return false;
    }

    /// <summary>호출 시작(start) 또는 마지막 VBA 모달 관측(lastSeen) 중 늦은 쪽부터 quietMs가 지났는가.</summary>
    public static bool IsQuiet(long now, long start, long lastSeen, int quietMs)
        => now - Math.Max(start, lastSeen) >= quietMs;

    public void Dispose()
    {
        _stop = true;
        _thread.Join(2000);
    }

    private void Loop()
    {
        long saveGuardUntil = 0; // TickCount64 기준
        var firstSeen = new Dictionary<IntPtr, long>(); // 폴백 유예용 최초 관측 시각

        while (!_stop)
        {
            try
            {
                var dialogs = ModalHunter.FindDialogs(_pid);
                foreach (var stale in firstSeen.Keys.Where(h => dialogs.All(d => d.Hwnd != h)).ToList())
                    firstSeen.Remove(stale);

                foreach (var dlg in dialogs)
                {
                    if (ModalHunter.IsVbaModalCaption(dlg.Caption))
                    {
                        long now = Environment.TickCount64;
                        Interlocked.Exchange(ref _lastVbaModalTick, now);
                        if (!firstSeen.TryGetValue(dlg.Hwnd, out long seen))
                            firstSeen[dlg.Hwnd] = seen = now;

                        string body = ModalHunter.CollectInfo(dlg.Hwnd);
                        var r = ModalHunter.Dismiss(dlg, allowFallback: now - seen >= FallbackGraceMs);
                        if (r == null) continue; // 버튼 미생성일 수 있음 — 유예 후 재판단

                        // enqueue는 VBE 닫기 이전 — Run은 VBE가 닫혀야 반환되므로(실증),
                        // Run 반환 시점에 큐에는 반드시 이 이벤트가 있다 (동기화 불필요 근거).
                        // 폴백이면 dismissal이 "wm_close [kind/no_safe_button]"으로 페이로드에 드러난다.
                        Enqueue("vbe_modal", dlg.Caption, body, r.Dismissal);

                        bool gone = ModalHunter.WaitUntilGone(dlg.Hwnd, 1500);
                        if (!gone) continue; // 다음 폴에서 재시도

                        Interlocked.Exchange(ref _lastDismissTick, Environment.TickCount64);
                        // VBE 닫기는 컴파일 오류(확인/OK)와 폴백에만 필요 — 런타임 오류(End)는
                        // 클릭만으로 제어 반환. 불필요한 VBE 닫기는 저장 프롬프트/XLMAIN 소멸 유발.
                        if (r.RequiresVbeClose)
                        {
                            Thread.Sleep(150);
                            ModalHunter.CloseVbeWindows(_pid);
                            // 가드 8s — 저장 프롬프트가 VBE 닫기 후 지연되어 뜨는 것을 실측(가드
                            // 경계 근처에 뜨면 놓쳐 COM 영구 차단).
                            saveGuardUntil = Environment.TickCount64 + 8000;
                        }
                    }
                    else if (dlg.Caption == "Microsoft Excel" && dlg.ClassName == "NUIDialog"
                             && Environment.TickCount64 < saveGuardUntil)
                    {
                        // VBE 강제 닫기의 부작용인 "변경 저장?" 프롬프트 — WM_CLOSE(=취소)로
                        // 종료를 중단해 워크북 유지 (DirectUI라 BM_CLICK 불가).
                        // 가드를 0으로 리셋하지 않는다 — 늦게 뜨는 2차 프롬프트도 잡기 위해.
                        ModalHunter.CloseWindow(dlg.Hwnd);
                        Enqueue("save_prompt", dlg.Caption, "", "wm_close");
                    }
                    else if (dlg.Caption == "Microsoft Excel" && dlg.ClassName == "#32770"
                             && Environment.TickCount64 < saveGuardUntil)
                    {
                        // VBE 닫기 후속 "대화 상자가 열려 있습니다" 알림 계열. 본문 텍스트 대신 구조로
                        // 판별: 답이 하나뿐인 알림(도움말 제외 활성 버튼 1개)만 누른다 — 오답이
                        // 있을 수 없으므로 언어와 무관하게 안전. 선택지가 둘 이상이면 건드리지 않는다.
                        var buttons = ModalHunter.SnapshotButtons(dlg.Hwnd).Select(b => b.Info).ToList();
                        if (ModalHunter.IsSingleAnswer(buttons))
                        {
                            string info = ModalHunter.CollectInfo(dlg.Hwnd);
                            var r = ModalHunter.Dismiss(dlg)!; // 폴백 허용 → 항상 non-null
                            Enqueue("excel_dialog_open", dlg.Caption, info, r.Dismissal);
                        }
                    }
                    // 가드 윈도 밖의 "Microsoft Excel" 다이얼로그(MsgBox 등)는 건드리지 않는다
                    // — 정당한 프롬프트 오답 위험. 타임아웃 kill 폴백이 담당 (명시적 non-goal).
                }
            }
            catch
            {
                // 워치독은 절대 죽지 않는다 — 다음 폴에서 재시도
            }
            Thread.Sleep(250);
        }
    }

    private void Enqueue(string kind, string caption, string body, string dismissal)
        => _events.Enqueue(new ModalEvent(
            DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), _currentLabel, kind, caption, body, dismissal));
}
