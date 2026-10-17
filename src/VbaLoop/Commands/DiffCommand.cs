using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop diff (설계 §4): 워크북 VBProject vs src 읽기 전용 비교. 에이전트가 세션
/// Excel에 직접 COM 접근하는 우회를 없애기 위한 공식 조회 경로 — src/워크북/sync.json
/// 어디에도 쓰지 않는다. persistent 모드는 attach 전용(auto-ensure 없음 — 진단이 세션 Excel을
/// 띄우는 것 자체가 부작용), 일회용 모드는 읽기 전용 일회용 인스턴스, lease 미획득(조회는 쓰기 의도가 아님 — status 선례), CommandMutex만 획득
/// (동시 push/test와 같은 Excel STA 인터리브 차단). 차이 존재는 실패가 아니다 —
/// exit 0 + payload.clean으로 보고한다.
/// </summary>
public static class DiffCommand
{
    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "diff" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "diff");
        o.LogRecords = log.Records;

        using var cmdLock = CommandMutex.Acquire(cfg);
        // 일회용 모드는 세션이 없는 것이 평상시라 읽기 전용 일회용 인스턴스로 본다(명시 세션이 있으면 그것).
        // persistent 모드는 기존대로 attach 전용 — 진단이 세션 Excel을 띄우는 것 자체가 부작용이다.
        using var ses = cfg.SessionMode == SessionMode.Ephemeral
            ? ExcelScope.Acquire(cfg, log, o, "diff", ScopeAccess.Read)
            : null;
        dynamic app;
        SessionMeta meta;
        if (ses != null) (app, meta) = (ses.App, ses.Meta);
        else (app, meta) = SessionService.AcquireOrThrow(cfg, "diff", log);
        var filter = OleMessageFilter.Register();
        try
        {
            return ComFailureGuard.Run(cfg, meta, log, () => Execute(o, opts, cfg, app, meta, log));
        }
        finally
        {
            filter.Revoke();
        }
    }

    private static Outcome Execute(Outcome o, CliArgs opts, ProjectConfig cfg, dynamic app, SessionMeta meta, ToolLog log)
    {
        dynamic wb = SessionService.FindWorkbook(app, meta.WorkbookFullName);
        dynamic proj = SessionService.GetVbProject(wb);

        var modules = new List<ModuleDiff.WorkbookModule>();
        var bodies = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var ci in VbaProjectOps.ListComponents(proj))
        {
            VbaComponentKind? kind = ci.Type switch
            {
                VbaProjectOps.TypeStdModule => VbaComponentKind.StdModule,
                VbaProjectOps.TypeClassModule => VbaComponentKind.ClassModule,
                VbaProjectOps.TypeMsForm => VbaComponentKind.MsForm,
                VbaProjectOps.TypeDocument => VbaComponentKind.Document,
                _ => null,
            };
            string body = kind != null ? VbaProjectOps.ReadModuleBody(ci.Comp.CodeModule) : "";
            if (kind != null) bodies[ci.Name] = body;
            modules.Add(new ModuleDiff.WorkbookModule(ci.Name, kind, body, ci.Type));
        }

        var srcFiles = new List<ModuleDiff.SrcFile>();
        if (Directory.Exists(cfg.SrcRoot))
        {
            foreach (var file in Directory.EnumerateFiles(cfg.SrcRoot))
            {
                var ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext is not (".bas" or ".cls" or ".frm")) continue;
                srcFiles.Add(new ModuleDiff.SrcFile(
                    Path.GetFileNameWithoutExtension(file), ext, File.ReadAllBytes(file)));
            }
        }

        IReadOnlyList<ModuleDiff.Entry> entries = ModuleDiff.Compare(modules, srcFiles);

        var summary = new
        {
            same = entries.Count(e => e.Verdict == "same"),
            changed = entries.Count(e => e.Verdict == "changed"),
            casingOnly = entries.Count(e => e.Verdict == "casingOnly"),
            onlyInWorkbook = entries.Count(e => e.Verdict == "onlyInWorkbook"),
            onlyInSrc = entries.Count(e => e.Verdict == "onlyInSrc"),
            skipped = entries.Count(e => e.Verdict == "skipped"),
            managed = entries.Count(e => e.Verdict == "managed"),
        };
        // managed(벤더링)·casingOnly(VBE 재케이싱, pull이 src 철자 보존)는 clean 판정에서 제외 —
        // 워크북↔src 바이트 차이가 pull 재기록으로 이어지지 않으므로.
        bool clean = summary.changed == 0 && summary.onlyInWorkbook == 0 && summary.onlyInSrc == 0;

        // --module: 표시 대상을 해당 모듈로 좁히고 워크북 쪽 본문 전체를 덤프.
        // clean/summary는 전체 비교 기준 유지 — 단건 조회가 전체 상태를 가리지 않도록.
        object? moduleDump = null;
        IReadOnlyList<ModuleDiff.Entry> shown = entries;
        if (opts.Module != null)
        {
            var matches = entries
                .Where(e => string.Equals(e.Name, opts.Module, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (matches.Length == 0)
                throw new CliError($"미지 모듈: {opts.Module} - 워크북/src 어느 쪽에도 없음",
                    "vbaloop diff --json으로 전체 모듈 목록 확인");
            shown = matches;
            bool found = bodies.TryGetValue(opts.Module, out var body);
            // verdict/detail을 덤프에 함께 실어 단건 조회가 라인 근거(firstDiffLine 등)를 바로
            // 노출하도록 한다.
            moduleDump = new
            {
                name = matches[0].Name,
                verdict = matches[0].Verdict,
                detail = matches[0].Detail,
                note = matches[0].Note,
                found,
                workbookBody = found
                    ? (body!.Length > 0 ? CanonicalText.ToCanonical(CanonicalText.SplitLines(body)) : "")
                    : null,
            };
        }

        o.HumanLines.Add(
            $"diff: same {summary.same}, changed {summary.changed}, casingOnly {summary.casingOnly}, " +
            $"onlyInWorkbook {summary.onlyInWorkbook}, onlyInSrc {summary.onlyInSrc}, " +
            $"skipped {summary.skipped}, managed {summary.managed}");
        foreach (var e in shown.Where(e => e.Verdict is not ("same" or "managed")))
        {
            var extra = e.Detail?.FirstDiffLine != null ? $" (첫 차이 {e.Detail.FirstDiffLine}행)" : "";
            var note = e.Note != null ? $" - {e.Note}" : "";
            o.HumanLines.Add($"  {e.Verdict}: {e.Name}{extra}{note}");
        }

        log.Write("info", "diff_done", new { summary, clean });

        o.Payload = new { clean, summary, modules = shown, module = moduleDump };
        return o;
    }
}
