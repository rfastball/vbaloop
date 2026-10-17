namespace VbaLoop.Core;

public enum ArtifactStatus { Ok, Missing, Drift }
public enum ArtifactAction { Created, Overwritten, Unchanged }

/// <summary>
/// init(온보딩 부트스트래퍼)의 순수 로직 (v1.1). 관리 아티팩트 8종을
/// toml -> shim -> CLAUDE.md -> AGENTS.md -> 스킬 -> pitfall 참조 -> dist 참조 -> .gitignore 순으로 분석/수렴한다.
/// toml은 관리 스탬프([vbaloop] managed_by, 설계 §12)만 라인 단위로 갱신한다 — 다른 아티팩트를 쓰거나
/// 스탬프가 없을 때. toml이 목록 첫 항목이라 스탬프가 먼저 기록된다(중간 실패 시에도 스탬프가 파일보다
/// 낮게 남지 않아 더 오래된 exe의 다운그레이드를 막는 쪽으로 실패한다).
/// Analyze가 기록할 내용까지 전부 준비하므로(추론 실패 시 여기서 throw)
/// Apply는 부분 실행 없이 쓰기만 한다. Excel/세션 불필요.
/// </summary>
public static class Onboarding
{
    public sealed class Item
    {
        public required string Name { get; init; }          // 루트 기준 상대 표시 경로
        public required string Path { get; init; }          // 절대 경로
        public required ArtifactStatus Status { get; init; }
        /// <summary>Status != Ok일 때 기록할 전체 텍스트(UTF-8 BOM 없음).</summary>
        public string? NewContent { get; init; }
        /// <summary>사람용 부가 설명(예: toml 스탬프 변경 "managed_by 0.2.0 -> 0.3.0").</summary>
        public string? Detail { get; init; }
    }

    /// <summary>8종 아티팩트 상태 + 기록 내용 준비. tomlPath가 null이면 워크북 추론으로
    /// 스캐폴드를 준비하고, 추론 불가(0개/복수)면 CliError. 기존 toml 파싱 실패도 CliError
    /// (사용자 소유 파일은 절대 덮어쓰지 않으므로 수렴 불가 = 에러).
    /// stampVersion: 관리 파일을 쓸 때 기록할 exe 버전(null이면 스탬프를 다루지 않음). 호출자가
    /// ManagedStamp.AllowsWrite를 먼저 확인한다 — 여기서도 스탬프를 내리는 쓰기는 만들지 않는다.</summary>
    public static List<Item> Analyze(string root, string? tomlPath,
        IReadOnlyList<string> shimCanonicalLines, IReadOnlyList<string> docsBodyLines,
        IReadOnlyList<string> skillCanonicalLines, IReadOnlyList<string> pitfallsCanonicalLines,
        IReadOnlyList<string> distCanonicalLines, string? stampVersion = null)
    {
        var items = new List<Item>();
        string srcRoot;
        Item? scaffold = null;

        if (tomlPath != null)
        {
            var cfg = ProjectConfig.Load(tomlPath, root); // 파싱/스키마 검증, 실패 시 CliError 전파
            srcRoot = cfg.SrcRoot;
        }
        else
        {
            srcRoot = System.IO.Path.Combine(root, "src");
            var workbook = InferWorkbook(root);
            var content = BuildTomlScaffold(new DirectoryInfo(root).Name, workbook);
            scaffold = new Item
            {
                Name = "vbaloop.toml",
                Path = System.IO.Path.Combine(root, "vbaloop.toml"),
                Status = ArtifactStatus.Missing,
                NewContent = stampVersion != null ? ManagedStamp.Upsert(content, stampVersion) : content,
            };
        }

        var shimPath = System.IO.Path.Combine(srcRoot, "LogShim.bas");
        var shimStatus = !File.Exists(shimPath) ? ArtifactStatus.Missing
            : IsWholeFileDrifted(shimPath, shimCanonicalLines) ? ArtifactStatus.Drift
            : ArtifactStatus.Ok;
        items.Add(new Item
        {
            Name = System.IO.Path.GetRelativePath(root, shimPath),
            Path = shimPath,
            Status = shimStatus,
            NewContent = shimStatus == ArtifactStatus.Ok ? null : CanonicalText.ToCanonical(shimCanonicalLines),
        });

        foreach (var docName in new[] { "CLAUDE.md", "AGENTS.md" })
        {
            var docPath = System.IO.Path.Combine(root, docName);
            string? existing = File.Exists(docPath) ? CanonicalText.ReadFileStripBom(docPath) : null;
            var (newText, changed) = ManagedBlock.Upsert(existing, docsBodyLines);
            var status = existing == null ? ArtifactStatus.Missing
                : changed ? ArtifactStatus.Drift
                : ArtifactStatus.Ok;
            items.Add(new Item
            {
                Name = docName,
                Path = docPath,
                Status = status,
                NewContent = status == ArtifactStatus.Ok ? null : newText,
            });
        }

        // 워크플로우 스킬 (whole-file-owned, LogShim과 동일 패턴 — DESIGN §8 스킬 계층 결정)
        var skillPath = System.IO.Path.Combine(root, ".claude", "skills", "vbaloop", "SKILL.md");
        var skillStatus = !File.Exists(skillPath) ? ArtifactStatus.Missing
            : IsWholeFileDrifted(skillPath, skillCanonicalLines) ? ArtifactStatus.Drift
            : ArtifactStatus.Ok;
        items.Add(new Item
        {
            Name = System.IO.Path.GetRelativePath(root, skillPath),
            Path = skillPath,
            Status = skillStatus,
            NewContent = skillStatus == ArtifactStatus.Ok ? null : CanonicalText.ToCanonical(skillCanonicalLines),
        });

        // VBA pitfall 참조 (스킬 references/ 지연 로드 층, whole-file-owned).
        // 정본은 vbaloop/assets — 축적은 정본에, 전파는 각 프로젝트의 init --check 드리프트 수렴으로.
        var pitfallsPath = System.IO.Path.Combine(
            root, ".claude", "skills", "vbaloop", "references", "vba-pitfalls.md");
        var pitfallsStatus = !File.Exists(pitfallsPath) ? ArtifactStatus.Missing
            : IsWholeFileDrifted(pitfallsPath, pitfallsCanonicalLines) ? ArtifactStatus.Drift
            : ArtifactStatus.Ok;
        items.Add(new Item
        {
            Name = System.IO.Path.GetRelativePath(root, pitfallsPath),
            Path = pitfallsPath,
            Status = pitfallsStatus,
            NewContent = pitfallsStatus == ArtifactStatus.Ok ? null : CanonicalText.ToCanonical(pitfallsCanonicalLines),
        });

        // dist 참조 (스킬 references/ 지연 로드 층, whole-file-owned — pitfall과 동형).
        // 배포는 편집 루프와 별개 태스크라 SKILL.md에서 분리해 상시 로드하지 않는다(DESIGN §8 스킬 계층).
        var distPath = System.IO.Path.Combine(
            root, ".claude", "skills", "vbaloop", "references", "dist.md");
        var distStatus = !File.Exists(distPath) ? ArtifactStatus.Missing
            : IsWholeFileDrifted(distPath, distCanonicalLines) ? ArtifactStatus.Drift
            : ArtifactStatus.Ok;
        items.Add(new Item
        {
            Name = System.IO.Path.GetRelativePath(root, distPath),
            Path = distPath,
            Status = distStatus,
            NewContent = distStatus == ArtifactStatus.Ok ? null : CanonicalText.ToCanonical(distCanonicalLines),
        });

        // .gitignore 관리 블록 (사용자 소유 파일 부분 관리 — CLAUDE.md와 동형 upsert).
        // dist 아티팩트/도구 스크래치(dist/, _dbg/)를 커밋 대상에서 제외(설계 §9.7).
        var gitignorePath = System.IO.Path.Combine(root, ".gitignore");
        string? gitignoreExisting = File.Exists(gitignorePath)
            ? CanonicalText.ReadFileStripBom(gitignorePath) : null;
        var (gitignoreText, gitignoreChanged) = GitignoreBlock.Upsert(gitignoreExisting, GitignoreBlock.ManagedLines);
        var gitignoreStatus = gitignoreExisting == null ? ArtifactStatus.Missing
            : gitignoreChanged ? ArtifactStatus.Drift
            : ArtifactStatus.Ok;
        items.Add(new Item
        {
            Name = ".gitignore",
            Path = gitignorePath,
            Status = gitignoreStatus,
            NewContent = gitignoreStatus == ArtifactStatus.Ok ? null : gitignoreText,
        });

        items.Insert(0, scaffold ?? AnalyzeTomlStamp(tomlPath!, stampVersion,
            othersChange: items.Any(i => i.Status != ArtifactStatus.Ok)));
        return items;
    }

    /// <summary>기존 toml의 스탬프 항목: 스탬프가 없으면 기록, exe가 더 새롭고 다른 관리 파일을 쓸 때만 올린다
    /// (바뀐 게 없는데 exe 갱신만으로 toml을 건드리는 churn을 피한다). 스탬프는 절대 내리지 않는다.</summary>
    private static Item AnalyzeTomlStamp(string tomlPath, string? stampVersion, bool othersChange)
    {
        var ok = new Item { Name = "vbaloop.toml", Path = tomlPath, Status = ArtifactStatus.Ok };
        if (stampVersion == null) return ok;
        var text = File.ReadAllText(tomlPath);
        var existing = ManagedStamp.Read(text);
        var state = ManagedStamp.Decide(SemVer.ParseOrNull(stampVersion), existing);
        bool write = state == StampState.Unstamped || (state == StampState.ExeNewer && othersChange);
        if (!write) return ok;
        return new Item
        {
            Name = "vbaloop.toml",
            Path = tomlPath,
            Status = ArtifactStatus.Drift,
            NewContent = ManagedStamp.Upsert(text, stampVersion),
            Detail = existing == null
                ? $"managed_by {stampVersion} 기록"
                : $"managed_by {existing} -> {stampVersion}",
        };
    }

    /// <summary>Analyze가 준비한 내용을 기록. Ok는 건드리지 않는다.</summary>
    public static List<(Item Item, ArtifactAction Action)> Apply(List<Item> items)
    {
        var results = new List<(Item, ArtifactAction)>();
        foreach (var item in items)
        {
            if (item.Status == ArtifactStatus.Ok)
            {
                results.Add((item, ArtifactAction.Unchanged));
                continue;
            }
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(item.Path)!);
            File.WriteAllText(item.Path, item.NewContent!, CanonicalText.Utf8NoBom);
            results.Add((item, item.Status == ArtifactStatus.Missing
                ? ArtifactAction.Created : ArtifactAction.Overwritten));
        }
        return results;
    }

    /// <summary>루트 최상위의 *.xlsm 중 Excel 잠금 스텁(~$*)을 제외하고 정확히 1개를 요구.</summary>
    public static string InferWorkbook(string root)
    {
        var candidates = Directory.EnumerateFiles(root, "*.xlsm", SearchOption.TopDirectoryOnly)
            .Select(System.IO.Path.GetFileName)
            .Where(n => n != null && !n.StartsWith("~$"))
            .Select(n => n!)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToList();
        return candidates.Count switch
        {
            1 => candidates[0],
            0 => throw new CliError($"워크북(*.xlsm)을 찾지 못함: {root}",
                "vbaloop.toml 수동 생성 후 vbaloop init 재실행"),
            _ => throw new CliError($"워크북 후보가 복수임: {string.Join(", ", candidates)}",
                "vbaloop.toml 수동 생성 후 vbaloop init 재실행"),
        };
    }

    public static string BuildTomlScaffold(string projectName, string workbookFileName)
    {
        var lines = new[]
        {
            "[project]",
            $"name = \"{projectName}\"",
            "",
            "[excel]",
            $"workbook = \"{workbookFileName}\"",
            "visible = true",
            "session = \"ephemeral\"",
            "",
            "[src]",
            "root = \"src\"",
        };
        return CanonicalText.ToCanonical(lines);
    }

    /// <summary>정규화(BOM/개행/후행 공백/끝 빈 줄 무시) 후 라인 시퀀스 ordinal 비교.
    /// whole-file-owned 아티팩트(LogShim/스킬) 공용 드리프트 판정.</summary>
    public static bool IsWholeFileDrifted(string filePath, IReadOnlyList<string> canonicalLines)
    {
        var actual = NormalizeForCompare(CanonicalText.SplitLines(CanonicalText.ReadFileStripBom(filePath)));
        var expected = NormalizeForCompare(canonicalLines);
        return !actual.SequenceEqual(expected, StringComparer.Ordinal);
    }

    private static string[] NormalizeForCompare(IReadOnlyList<string> lines)
        => CanonicalText.TrimTrailingEmpty(lines.Select(l => l.TrimEnd()).ToArray());
}
