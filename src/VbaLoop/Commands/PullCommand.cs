using System.Text;
using VbaLoop.Cli;
using VbaLoop.Core;
using VbaLoop.Excel;

namespace VbaLoop.Commands;

/// <summary>
/// vbaloop pull (설계 §4): CodeModule.Lines(BSTR 직접) → Export 호환 직렬화 →
/// UTF-8(BOM 없음)/CRLF 기록. 폼만 Export 파일 경로 예외(폼 예외).
/// </summary>
public static class PullCommand
{
    private sealed class ModuleResult
    {
        public required string Name { get; init; }
        public required string Action { get; set; } // added|replaced|skipped|failed|pruned
        public string? Reason { get; set; }
    }

    public static Outcome Run(CliArgs opts)
    {
        var o = new Outcome { Command = "pull" };
        var cfg = ProjectConfig.Load(opts.Config, Environment.CurrentDirectory);
        var log = new ToolLog(cfg.ToolLogPath, "pull");
        o.LogRecords = log.Records;

        using var cmdLock = CommandMutex.Acquire(cfg);
        LeaseService.AcquireOrRenew(cfg, "pull", log);

        var ses = SessionEnsure.ForCommand(cfg, log, o, "pull");
        var (app, meta) = (ses.App, ses.Meta);
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

        var modules = new List<ModuleResult>();
        var componentBaseNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Directory.CreateDirectory(cfg.SrcRoot);

        var sync = SyncCache.Load(SyncCache.PathFor(cfg));

        // 미푸시 편집 가드(설계 §4 pull): 지문이 마지막 push/pull 어느 쪽과도 다르면
        // 동기화된 적 없는 src 편집 — 파일 기록 시작 전에 차단한다.
        string? blockReason = PullGuard.FindBlockReason(
            Fingerprint.Compute(cfg.SrcRoot), sync.LastPushFingerprint, sync.LastPullFingerprint);
        if (blockReason != null)
        {
            if (!opts.Force)
                throw new CliError($"pull 차단: {blockReason}",
                    "vbaloop push로 먼저 반영, 의도된 덮어쓰기면 vbaloop pull --force");
            o.HumanLines.Add("경고: 동기화되지 않은 src 편집을 --force로 덮어씀");
            log.Write("warn", "pull_forced", new { reason = blockReason });
        }

        List<VbaProjectOps.ComponentInfo> components = VbaProjectOps.ListComponents(proj);
        foreach (var ci in components)
        {
            componentBaseNames.Add(ci.Name);

            // 벤더링 관리 모듈(LogShim)은 정본이 init이지 워크북이 아니다 — pull이 워크북
            // 사본으로 덮으면 가짜 드리프트를 만든다. src 무결성은
            // init --check가 담당. --force도 예외 없음(관리 사본은 사용자 편집 대상이 아님).
            if (VendoredModules.IsVendored(ci.Name))
            {
                log.Write("info", "pull_vendored_skipped", new { module = ci.Name });
                modules.Add(new ModuleResult
                {
                    Name = ci.Name,
                    Action = "skipped",
                    Reason = "벤더링 관리 모듈 — 정본은 init(pull 회수 대상 아님, 갱신은 vbaloop init)",
                });
                continue;
            }

            try
            {
                switch (ci.Type)
                {
                    case VbaProjectOps.TypeStdModule:
                    case VbaProjectOps.TypeClassModule:
                    case VbaProjectOps.TypeDocument:
                        modules.Add(PullCodeComponent(cfg, ci, o, log));
                        break;
                    case VbaProjectOps.TypeMsForm:
                        modules.Add(PullForm(cfg, ci, log, sync));
                        break;
                    default:
                        modules.Add(new ModuleResult
                        {
                            Name = ci.Name,
                            Action = "skipped",
                            Reason = $"미지원 컴포넌트 타입 {ci.Type}",
                        });
                        break;
                }
            }
            catch (CliError e)
            {
                modules.Add(new ModuleResult { Name = ci.Name, Action = "failed", Reason = e.Message });
                log.Write("error", "pull_module_failed", new { module = ci.Name, error = e.Message });
            }
        }

        // 워크북에서 사라진 모듈의 src 파일은 삭제하지 않고 경고 보고(설계 §4 pull)
        var orphans = new List<string>();
        foreach (var file in Directory.EnumerateFiles(cfg.SrcRoot))
        {
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext is not (".bas" or ".cls" or ".frm")) continue;
            var baseName = Path.GetFileNameWithoutExtension(file);
            if (!componentBaseNames.Contains(baseName))
                orphans.Add(Path.GetFileName(file));
        }

        int failed = modules.Count(m => m.Action == "failed");
        var summaryCounts = modules.GroupBy(m => m.Action)
            .ToDictionary(g => g.Key, g => g.Count());

        o.HumanLines.Add($"pull 완료: {string.Join(", ", summaryCounts.Select(kv => $"{kv.Key} {kv.Value}"))}");
        foreach (var orphan in orphans)
            o.HumanLines.Add($"경고: 워크북에 없는 src 파일(orphan): {orphan} — 삭제하지 않음");

        // 사후 지문 기록: pull이 만든 디스크 상태가 새 동기화 기준선 — 부분 실패여도
        // 기록한다(사후 상태는 도구 산출물이지 사용자 편집이 아니므로 재pull이 막히면 안 됨).
        sync.LastPullFingerprint = Fingerprint.Compute(cfg.SrcRoot);
        sync.LastPullAt = DateTime.UtcNow.ToString("o");
        sync.Save(SyncCache.PathFor(cfg)); // 폼 해시 캐시 갱신분 포함 영속화

        log.Write("info", "pull_done", new { counts = summaryCounts, orphans });

        o.Payload = new
        {
            modules = modules.Select(m => new { name = m.Name, action = m.Action, reason = m.Reason }),
            orphans,
        };
        if (failed > 0)
            o.Fail($"pull 부분 실패: 모듈 {failed}개 실패 — payload.modules에서 상세 확인");
        return o;
    }

    private static ModuleResult PullCodeComponent(ProjectConfig cfg, VbaProjectOps.ComponentInfo ci, Outcome o, ToolLog log)
    {
        string ext = ci.Type == VbaProjectOps.TypeStdModule ? ".bas" : ".cls";
        string path = Path.Combine(cfg.SrcRoot, ci.Name + ext);

        string body = VbaProjectOps.ReadModuleBody(ci.Comp.CodeModule);
        IReadOnlyList<string> bodyLines = body.Length > 0 ? CanonicalText.SplitLines(body) : Array.Empty<string>();

        IReadOnlyList<string> headerLines;
        bool existed = File.Exists(path);
        if (existed)
        {
            // 기존 src 파일의 헤더 엔벨로프 보존, 본문만 교체(설계 §2 파일 포맷)
            var split = VbaHeader.Split(
                CanonicalText.SplitLines(CanonicalText.ReadFileStripBom(path)), path, strict: false);
            headerLines = split.Header;
            if (!headerLines.Any(l => l.TrimStart().StartsWith("Attribute VB_Name", StringComparison.Ordinal)))
                headerLines = Synthesize(ci, o);
            // VBE 재케이싱이 src에 영속되지 않도록 기존 src 철자 보존 —
            // 토큰 등가관계 + LCS 정렬(라인 삽입/삭제에 견고), 진짜 변경만 워크북 채택. --force도 동일
            // (케이싱은 비의미적이므로 언제나 src 철자로 수렴시키는 게 옳다).
            bodyLines = VbaCasing.ReconcileBlock(bodyLines, split.Body);
        }
        else
        {
            headerLines = Synthesize(ci, o);
        }

        var newLines = headerLines.Concat(bodyLines).ToList();
        return WriteIfChanged(ci.Name, path, newLines, existed);
    }

    private static IReadOnlyList<string> Synthesize(VbaProjectOps.ComponentInfo ci, Outcome o)
    {
        var kind = ci.Type switch
        {
            VbaProjectOps.TypeStdModule => VbaComponentKind.StdModule,
            VbaProjectOps.TypeDocument => VbaComponentKind.Document,
            _ => VbaComponentKind.ClassModule,
        };
        if (kind == VbaComponentKind.ClassModule)
            o.HumanLines.Add(
                $"안내: {ci.Name}은 표준 클래스 헤더로 합성됨 — VB_PredeclaredId 등 비표준 속성이 필요하면 최초 1회 수동 Export로 헤더를 시드할 것");
        return VbaHeader.SynthesizeHeader(kind, ci.Name);
    }

    private static string HashFile(string path) => HashBytes(File.ReadAllBytes(path));

    private static string HashBytes(byte[] bytes)
        => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));

    private static ModuleResult PullForm(
        ProjectConfig cfg, VbaProjectOps.ComponentInfo ci, ToolLog log, SyncCache sync)
    {
        // 폼 예외: 폼 레이아웃 블록과 .frx는 Export 없이 왕복 불가.
        // 인코딩은 GetACP 동적 조회 — 상수 하드코딩 금지.
        string tempDir = Path.Combine(Path.GetTempPath(), "vbaloop", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        try
        {
            string frmTemp = Path.Combine(tempDir, ci.Name + ".frm");
            ComRetry.Invoke(() => { ci.Comp.Export(frmTemp); });

            var acp = Encoding.GetEncoding(NativeMethods.GetACP());
            string frmText = acp.GetString(File.ReadAllBytes(frmTemp));

            string frmPath = Path.Combine(cfg.SrcRoot, ci.Name + ".frm");
            string frxPath = Path.Combine(cfg.SrcRoot, ci.Name + ".frx");
            string frxTemp = Path.Combine(tempDir, ci.Name + ".frx");
            bool frmExisted = File.Exists(frmPath);

            // 코드비하인드 본문의 VBE 재케이싱을 기존 src 철자로 흡수 —
            // 폼도 코드 모듈과 동일하게 커버(구 미봉책 결함 #3). 헤더/레이아웃 엔벨로프는 fresh
            // Export 그대로 두고 본문만 토큰 등가관계 + LCS 정렬로 수렴시킨 뒤, 트림/스트립보다
            // 먼저 수행해 LCS 인덱스를 흔들지 않는다.
            string[] wbLinesArr = CanonicalText.SplitLines(frmText);
            IReadOnlyList<string> wbLines = wbLinesArr;
            if (frmExisted)
            {
                string frmFile = ci.Name + ".frm";
                var wbSplit = VbaHeader.Split(wbLinesArr, frmFile, strict: false);
                var srcSplit = VbaHeader.Split(
                    CanonicalText.SplitLines(CanonicalText.ReadFileStripBom(frmPath)), frmFile, strict: false);
                wbLines = wbSplit.Header
                    .Concat(VbaCasing.ReconcileBlock(wbSplit.Body, srcSplit.Body)).ToList();
            }
            // Import가 코드비하인드 끝에 빈 줄을 추가하는 VBE 동작 상쇄 (왕복 수렴) +
            // VBE가 WithEvents 멤버에 삽입하는 Attribute .VB_VarHelpID 라인 제거(push 왕복 불가 방지)
            var frmLines = CanonicalText.StripMemberAttributes(
                CanonicalText.TrimTrailingEmpty(wbLines));

            byte[] newFrmBytes = CanonicalText.Utf8NoBom.GetBytes(CanonicalText.ToCanonical(frmLines));
            bool frmTextSame = frmExisted && File.ReadAllBytes(frmPath).AsSpan().SequenceEqual(newFrmBytes);

            // frx 비결정성 회피(실측): frm 텍스트가 그대로이고 src frx가 마지막 동기화 상태
            // 그대로이면(사용자가 의도적으로 교체하지 않음) fresh Export의 frx 바이트가 달라도
            // 재기록하지 않는다 — 매 pull마다 frx가 바뀌던 git 노이즈 제거.
            sync.Forms.TryGetValue(ci.Name, out var formSync);
            string? srcFrxHash = File.Exists(frxPath) ? HashFile(frxPath) : null;
            bool frxCacheMatches = formSync?.FrxHash != null && formSync.FrxHash == srcFrxHash;

            if (frmTextSame && frxCacheMatches)
                return new ModuleResult { Name = ci.Name, Action = "skipped", Reason = "변경 없음(frx 노이즈 생략)" };

            // 변경(또는 최초 시딩): frm/frx 기록 후 캐시를 방금 쓴 src 파일 기준으로 갱신
            if (!frmTextSame) File.WriteAllBytes(frmPath, newFrmBytes);
            if (File.Exists(frxTemp)) File.WriteAllBytes(frxPath, File.ReadAllBytes(frxTemp));

            sync.Forms[ci.Name] = new SyncCache.FormSync
            {
                FrmHash = HashBytes(newFrmBytes),
                FrxHash = File.Exists(frxPath) ? HashFile(frxPath) : null,
            };

            return new ModuleResult
            {
                Name = ci.Name,
                Action = frmExisted ? "replaced" : "added",
                Reason = frmTextSame ? "frx만 갱신(캐시 시딩)" : null,
            };
        }
        finally
        {
            try { Directory.Delete(tempDir, recursive: true); } catch { /* 임시 폴더 정리 실패 무시 */ }
        }
    }

    private static ModuleResult WriteIfChanged(string name, string path, IReadOnlyList<string> newLines, bool existed)
    {
        // 비교는 바이트 기준 — 텍스트 비교는 디코더가 BOM을 흡수해 BOM/개행 위반 파일이
        // 영영 재정규화되지 않는다(lint 실측으로 발견). 설계 §4: 최초 pull은 일괄 정규화.
        byte[] newBytes = CanonicalText.Utf8NoBom.GetBytes(CanonicalText.ToCanonical(newLines));
        if (existed)
        {
            byte[] oldBytes = File.ReadAllBytes(path);
            if (oldBytes.AsSpan().SequenceEqual(newBytes))
                return new ModuleResult { Name = name, Action = "skipped", Reason = "변경 없음" };
        }
        File.WriteAllBytes(path, newBytes);
        return new ModuleResult
        {
            Name = name,
            Action = existed ? "replaced" : "added",
        };
    }
}
