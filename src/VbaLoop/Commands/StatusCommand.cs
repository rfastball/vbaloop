using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// 세션 상태(none/active/poisoned/stale), 워크북 경로, src 최종 수정시각 vs
/// 마지막 push 성공 타임스탬프 비교 출력(설계 §4 status). 메타파일을 읽기만 한다.
/// </summary>
public static class StatusCommand
{
    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "status" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        // 격리 중이면 세션 Excel에 COM 탐침(Inspect)조차 하지 않는다(설계 §11)
        var marker = RecoveryMarker.Load(cfg.RecoveryMarkerPath);
        var info = marker != null
            ? new SessionInfo { State = SessionState.None, Meta = SessionMeta.Load(cfg.SessionMetaPath) }
            : SessionService.Inspect(cfg);

        string state = marker != null ? "quarantined" : info.State.ToString().ToLowerInvariant();

        DateTime? srcLastModified = null;
        if (Directory.Exists(cfg.SrcRoot))
        {
            var files = Directory.EnumerateFiles(cfg.SrcRoot)
                .Where(f => Fingerprint.SourceExtensions.Contains(Path.GetExtension(f).ToLowerInvariant()))
                .ToList();
            if (files.Count > 0)
                srcLastModified = files.Max(f => File.GetLastWriteTimeUtc(f));
        }

        var sync = SyncCache.Load(SyncCache.PathFor(cfg));
        DateTime? lastPushAt = null;
        if (sync.LastPushAt is string s && DateTime.TryParse(s, null,
                System.Globalization.DateTimeStyles.RoundtripKind, out var parsed))
            lastPushAt = parsed.ToUniversalTime();

        // 최신성 판정은 mtime이 아니라 지문 비교 — pull도 src를 재기록하므로 mtime 기준은
        // pull 직후 "push 필요"로 오판한다(SampleRpa 파일럿 피드백). mtime은 참고 표기로만.
        string? currentFingerprint = srcLastModified != null ? Fingerprint.Compute(cfg.SrcRoot) : null;
        bool? srcDiffersFromPush =
            currentFingerprint != null && sync.LastPushFingerprint != null
                ? currentFingerprint != sync.LastPushFingerprint
                : null;

        string freshness =
            currentFingerprint == null ? "src 파일 없음"
            : sync.LastPushFingerprint == null ? "push 이력 없음 (src가 미반영 상태일 수 있음)"
            : srcDiffersFromPush == true ? "src가 마지막 push와 다름 (push 필요)"
            : "src와 마지막 push 일치 (반영됨)";

        var lease = LeaseService.Inspect(cfg);

        o.HumanLines.Add($"session : {state}" + (marker != null ? $" ({marker.Describe()})"
            : info.Detail != null ? $" ({info.Detail})" : ""));
        o.HumanLines.Add(LeaseCommand.DescribeState(lease));
        o.HumanLines.Add($"workbook: {cfg.WorkbookPath}" + (File.Exists(cfg.WorkbookPath) ? "" : " (파일 없음!)"));
        o.HumanLines.Add($"config  : {cfg.TomlPath}" + (ProjectConfig.IsLegacyToml(cfg.TomlPath) ? " (옛 이름 — vbaloop init으로 이전)" : ""));
        o.HumanLines.Add($"src     : {cfg.SrcRoot} (최종 수정 {Fmt(srcLastModified)})");
        o.HumanLines.Add($"push    : {Fmt(lastPushAt)} — {freshness}");

        if (marker != null)
            o.Hint = RecoveryGate.HintFor(marker);
        else if (info.State == SessionState.Busy)
            o.Hint = "Excel에서 셀 편집/열린 대화상자를 마친 뒤 재시도";
        else if (info.State is SessionState.None or SessionState.Stale or SessionState.Poisoned)
            o.Hint = "session start";

        o.Payload = new
        {
            state,
            detail = info.Detail,
            workbook = cfg.WorkbookPath,
            workbookExists = File.Exists(cfg.WorkbookPath),
            tomlPresent = true,
            srcRoot = cfg.SrcRoot,
            srcLastModifiedUtc = srcLastModified?.ToString("o"),
            lastPushAtUtc = lastPushAt?.ToString("o"),
            lastPullAtUtc = sync.LastPullAt,
            srcFingerprint = currentFingerprint,
            lastPushFingerprint = sync.LastPushFingerprint,
            lastPullFingerprint = sync.LastPullFingerprint,
            srcDiffersFromPush,
            recovery = RecoveryCommand.DescribeOrNull(marker),
            session = info.Meta == null ? null : new
            {
                pid = info.Meta.Pid,
                owner = info.Meta.Owner ?? SessionMeta.OwnerTool,
                hwnd = info.Meta.Hwnd,
                createdAt = info.Meta.CreatedAt,
                lastCommand = info.Meta.LastCommand,
                poisonReason = info.Meta.PoisonReason,
                hResult = info.Meta.HResult,
            },
            lease = new
            {
                state = lease.Owner == null ? (lease.FilePresent ? "corrupt" : "none")
                    : lease.Mine ? "owned" : lease.Reclaimable ? "reclaimable" : "busy",
                ownerThread = lease.Owner?.OwnerThread,
                idleSeconds = double.IsNaN(lease.IdleSeconds) ? (double?)null : lease.IdleSeconds,
            },
        };
        return o;
    }

    private static string Fmt(DateTime? dt) => dt?.ToString("yyyy-MM-dd HH:mm:ss'Z'") ?? "-";
}
