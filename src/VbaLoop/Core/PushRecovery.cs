namespace VbaLoop.Core;

/// <summary>
/// push 부분 실패 복구 판정(순수, 설계 §4 push 실패 복구). push는 실패 시 저장하지 않으므로 디스크 정본은
/// 항상 이번 push 이전 상태다 — 문제는 세션 Excel의 메모리 워크북이다. 앞선 모듈이 이미 주입된 채
/// 남으면 test/run이 혼합 코드를 돌리고, 기본 session stop(Save)이 혼합 상태를 정본에 굳힌다.
/// 그래서 실패 시 둘 중 하나만 허용한다:
///   (a) 이번 push가 건드린 모든 모듈을 스냅샷으로 되돌리고 검증(Restored), 또는
///   (b) 되돌릴 수 없거나 되돌리기가 하나라도 실패하면 세션을 poison(Poisoned) — 다음 명령이 Excel을
///       kill(저장 없음) 후 디스크 정본을 다시 연다.
/// 이 클래스는 "무엇을 되돌릴지 / poison할지 / 사용자에게 뭐라고 말할지"만 정한다. 실행은 PushCommand.
/// </summary>
public static class PushRecovery
{
    /// <summary>push가 메모리 워크북에 가한 변경 종류.</summary>
    public enum TouchKind
    {
        /// <summary>기존 모듈 본문 교체 — 교체 전 스냅샷 재주입으로 되돌릴 수 있다.</summary>
        CodeReplaced,
        /// <summary>신규 컴포넌트 생성(표준/클래스/신규 폼) — 제거로 되돌릴 수 있다.</summary>
        ComponentAdded,
        /// <summary>기존 폼 Remove → Import — 원본 폼 바이너리가 없어 되돌릴 수 없다.</summary>
        FormReimported,
        /// <summary>--prune 삭제 — 되돌릴 수 없다.</summary>
        ComponentPruned,
    }

    public sealed record Touch(string Module, TouchKind Kind);

    /// <summary>실패가 난 단계.</summary>
    public enum Phase
    {
        /// <summary>모듈별 주입 루프.</summary>
        Inject,
        /// <summary>--prune 제거.</summary>
        Prune,
        /// <summary>save 직전 무결성 체크(이름 집합/빈 본문) — 상태 자체를 신뢰할 수 없다.</summary>
        Integrity,
    }

    public enum WorkbookState
    {
        /// <summary>메모리 워크북에 아무 변경도 없었다.</summary>
        Untouched,
        /// <summary>건드린 모듈 전부 원복·검증됨(메모리, 저장 안 함).</summary>
        Restored,
        /// <summary>혼합 상태 — 세션 poison, 저장 금지. 디스크 정본은 push 이전 상태.</summary>
        Poisoned,
    }

    /// <param name="RestoreOrder">되돌릴 모듈(역순 — 마지막 변경부터). Poison이면 비어 있다.</param>
    public sealed record Plan(IReadOnlyList<Touch> RestoreOrder, bool Poison, string? PoisonReason);

    public static bool IsRestorable(TouchKind k) => k is TouchKind.CodeReplaced or TouchKind.ComponentAdded;

    /// <summary>
    /// 복구 계획. 모듈별로 가장 나쁜 변경 하나로 접는다(ComponentAdded는 이후 본문 교체를 포함하고,
    /// 되돌릴 수 없는 변경이 하나라도 섞이면 그 모듈은 되돌릴 수 없다).
    /// </summary>
    public static Plan Decide(IReadOnlyList<Touch> touches, Phase phase)
    {
        if (phase == Phase.Integrity)
            return new Plan(Array.Empty<Touch>(), true,
                "integrity check failed — in-memory workbook state cannot be trusted");
        if (touches.Count == 0)
            return new Plan(Array.Empty<Touch>(), false, null);

        var collapsed = Collapse(touches);
        var unrestorable = collapsed.Where(t => !IsRestorable(t.Kind)).ToList();
        if (unrestorable.Count > 0)
            return new Plan(Array.Empty<Touch>(), true,
                "changes that cannot be rolled back from the snapshot: " +
                string.Join(", ", unrestorable.Select(t => $"{t.Module}({Label(t.Kind)})")));

        return new Plan(collapsed.AsEnumerable().Reverse().ToList(), false, null);
    }

    /// <summary>계획 실행 결과로 최종 상태 판정.</summary>
    public static WorkbookState Resolve(Plan plan, int touchCount, int restoreFailures, bool poisonedByFatalCom = false)
    {
        if (plan.Poison || restoreFailures > 0 || poisonedByFatalCom) return WorkbookState.Poisoned;
        return touchCount == 0 ? WorkbookState.Untouched : WorkbookState.Restored;
    }

    /// <summary>사용자/에이전트에게 보일 상태 요약(에러 문구 꼬리)과 hint — 워크북이 실제 어떤 상태인지만 말한다.</summary>
    public static (string Summary, string Hint) Describe(
        WorkbookState state, int restoredCount, string? backupId, string? poisonReason)
    {
        string rollback = backupId != null
            ? $"vbaloop rollback --latest (pre-push backup {backupId})"
            : "rollback unavailable (backup disabled by [backup] keep = 0)";
        return state switch
        {
            WorkbookState.Untouched => (
                "workbook unchanged (not saved)",
                "fix the cause, then retry"),
            WorkbookState.Restored => (
                $"all {restoredCount} modules touched by this push rolled back to the snapshot and verified (in memory, not saved)",
                $"fix the cause, then retry. If the workbook is suspect, run session stop --no-save, then {rollback}"),
            _ => (
                "in-memory workbook is in a mixed state, so the session is poisoned — not saved; the canonical workbook on disk is still in its pre-push state" +
                (poisonReason != null ? $" ({poisonReason})" : ""),
                "The next command restarts the session Excel without saving and reopens the canonical workbook from disk (unsaved edits in the session window are discarded). " +
                $"If the canonical workbook on disk is also suspect, run vbaloop session stop, then {rollback}"),
        };
    }

    /// <summary>
    /// 스냅샷 원복 검증(순수): 말미 빈 줄을 무시하고 줄 단위 대소문자 무시 비교. AddFromString은 말미 빈 줄을
    /// 덧붙일 수 있고, VBE는 다른 모듈의 선언을 따라 식별자 대소문자를 재케이싱한다(2026-07-09) — 둘 다
    /// 원복 실패가 아니다. 그 외 차이는 원복 실패로 보고 poison한다.
    /// </summary>
    public static bool SnapshotMatches(string expected, string actual)
        => CanonicalText.TrimTrailingEmpty(CanonicalText.SplitLines(expected))
            .SequenceEqual(CanonicalText.TrimTrailingEmpty(CanonicalText.SplitLines(actual)),
                StringComparer.OrdinalIgnoreCase);

    private static List<Touch> Collapse(IReadOnlyList<Touch> touches)
    {
        var order = new List<string>();
        var worst = new Dictionary<string, Touch>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in touches)
        {
            if (!worst.TryGetValue(t.Module, out var prev))
            {
                order.Add(t.Module);
                worst[t.Module] = t;
            }
            else if (Severity(t.Kind) > Severity(prev.Kind))
            {
                worst[t.Module] = t with { Module = prev.Module };
            }
        }
        return order.Select(m => worst[m]).ToList();
    }

    private static int Severity(TouchKind k) => k switch
    {
        TouchKind.CodeReplaced => 0,
        TouchKind.ComponentAdded => 1,
        _ => 2,
    };

    private static string Label(TouchKind k) => k switch
    {
        TouchKind.CodeReplaced => "body replaced",
        TouchKind.ComponentAdded => "newly created",
        TouchKind.FormReimported => "form re-Import",
        TouchKind.ComponentPruned => "prune delete",
        _ => k.ToString(),
    };
}
