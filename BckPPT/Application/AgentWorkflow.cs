using System.IO;
using BckPPT.Agent;
using BckPPT.Agent.Cards;
using BckPPT.Llm;
using BckPPT.PowerPoint;
using BckPPT.Rag;
using BckPPT.Services;

namespace BckPPT.Application;

/// <summary>한 번의 요청 결과. 썸네일은 건드린 슬라이드만, 적용 전은 원본에서 같은 자리를 렌더한다.</summary>
public sealed record AgentTurnResult(
    AgentRunResult Run,
    IReadOnlyList<string> BeforeThumbnails,
    IReadOnlyList<string> AfterThumbnails,
    IReadOnlyList<string> Changes,
    IReadOnlyDictionary<string, (int Calls, LlmUsage Usage)>? UsageByModel = null);

/// <summary>
/// 채팅창과 도구 루프 사이. 대상 PPT마다 세션(작업 사본 + 대화)을 하나 두고, 같은 대상에 대한 후속 요청은
/// 그 세션에서 이어간다. 승인하면 결과 사본을 남기고 세션을 끝낸다. 대상이 바뀌거나 버리면 작업 사본을 지운다.
/// </summary>
public sealed class AgentWorkflow(
    ILlmClient llm, GitokyLlmOptions? options, IPptComService ppt, RenderValidator renderer, SourceLibrary library)
    : IDisposable
{
    private readonly List<string> _attached = new();
    private UsageMeter? _meter;
    private MeteredLlmClient? _metered;

    public AgentSession? Session { get; private set; }

    /// <summary>턴·토큰 상한 등. 큰 작업(카드 여러 장)은 기본 30턴으로 모자랄 수 있다.</summary>
    public AgentLoopOptions LoopOptions { get; set; } = new();
    public IReadOnlyList<string> AttachedSources => _attached;

    /// <summary>원문 파일이나 폴더를 색인하고 이번 작업의 검색 대상으로 붙인다. (파일 수, 쪽 수)를 돌려준다.</summary>
    public async Task<(int Files, int Pages)> AttachAsync(IEnumerable<string> paths, CancellationToken ct = default)
    {
        int files = 0, pages = 0;
        foreach (var path in paths)
        {
            ct.ThrowIfCancellationRequested();
            if (Directory.Exists(path))
            {
                var found = PreferTextVersions(Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories)
                    .Where(f => SourceLibrary.SupportedExtensions.Contains(Path.GetExtension(f)))
                    .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal)));
                foreach (var file in found) pages += await AttachFileAsync(file, ct);
                files += found.Count;
            }
            else
            {
                pages += await AttachFileAsync(path, ct);
                files++;
            }
        }
        return (files, pages);
    }

    private async Task<int> AttachFileAsync(string path, CancellationToken ct)
    {
        var full = Path.GetFullPath(path);
        int pages = await Task.Run(() => library.AddFile(full), ct);
        if (!_attached.Contains(full, StringComparer.OrdinalIgnoreCase)) _attached.Add(full);
        return pages;
    }

    public void DetachAll() => _attached.Clear();

    /// <summary>
    /// 카드 문안을 사용자 고급 모델(ChatGPT·Claude 등)에서 쓰게 한다. gitoky는 비용 관리가 어려워
    /// 비싼 모델 대신 사용자가 요청문을 복사해 자기 모델에 붙여넣고, 답을 붙여넣어 적용한다.
    /// </summary>
    public bool ExternalWriting { get; set; }

    /// <summary>길이 맞춤에 qwen을 쓸지. 끄면 넘친 항목을 코드로만 뺀다.</summary>
    public bool LlmFit { get; set; } = true;

    /// <summary>외부 작성 대기 중인 카드 작업(요청문 포함).</summary>
    public PendingCards? PendingCards => Session?.Context.PendingCards;

    /// <summary>사용자 모델의 답을 붙여넣어 카드 작업을 이어간다.</summary>
    public Task<AgentTurnResult> ApplyPendingCardsAsync(string answer, IProgress<AgentEvent>? progress, CancellationToken ct) =>
        ContinuePendingAsync((pipeline, pending) => pipeline.ApplyExternalAsync(pending, answer, ct), progress, ct);

    /// <summary>외부 작성 없이 기본 모델(qwen)로 문안을 써서 이어간다.</summary>
    public Task<AgentTurnResult> WritePendingWithDefaultAsync(IProgress<AgentEvent>? progress, CancellationToken ct) =>
        ContinuePendingAsync((pipeline, pending) => pipeline.WritePendingWithDefaultAsync(pending, ct), progress, ct);

    private async Task<AgentTurnResult> ContinuePendingAsync(
        Func<CardPipeline, PendingCards, Task<string>> step, IProgress<AgentEvent>? progress, CancellationToken ct)
    {
        var session = Session ?? throw new InvalidOperationException("이어갈 작업이 없습니다.");
        var pending = session.Context.PendingCards ?? throw new InvalidOperationException("붙여넣기를 기다리는 카드 작업이 없습니다.");
        var meterBefore = _meter!.Snapshot();
        session.Context.Status = progress is null ? null
            : new Progress<string>(text => progress.Report(new AgentEvent(AgentEventKind.Message, text)));
        session.Context.LlmFit = LlmFit;
        var report = await step(new CardPipeline(session.Context), pending);
        session.Context.PendingCards = null;   // 답을 읽지 못하면 위에서 예외가 나 대기 상태가 유지된다
        session.Log.Write("cards_applied", new { report = AgentSessionLog.Clip(report) });

        var issues = new List<BckPPT.PowerPoint.LintIssue>();
        foreach (var slide in session.Context.Origins.Touched)
            issues.AddRange(await Agent.Tools.PptAgentTools.LintAsync(session.Context, slide, null, ct));
        var usage = UsageMeter.Delta(meterBefore, _meter.Snapshot());
        var run = new AgentRunResult(AgentStopReason.Finished, report, 0,
            usage.Values.Aggregate(LlmUsage.Zero, (a, v) => a + v.Usage), issues);
        var (before, after) = await RenderAsync(session, ct);
        return new AgentTurnResult(run, before, after, session.Context.Changes.ToList(), usage);
    }

    private async Task<(IReadOnlyList<string> Before, IReadOnlyList<string> After)> RenderAsync(AgentSession session, CancellationToken ct)
    {
        var touched = session.Context.Origins.Touched.ToHashSet();
        var after = touched.Count == 0 ? [] : await renderer.RenderSlidesAsync(session.WorkingPath, touched, ct);
        var originals = touched.Select(session.Context.Origins.OriginalIndexOf).ToHashSet();
        var before = originals.Count == 0 ? [] : await renderer.RenderSlidesAsync(session.SourcePath, originals, ct);
        return (before, after);
    }

    /// <summary>
    /// 같은 폴더에 같은 이름의 원문이 여러 형식으로 있으면(보고서.pdf와 PDF에서 뽑은 보고서.md) 하나만 남긴다.
    /// md·txt(이미 뽑은 텍스트)를 pdf·pptx보다 먼저 고른다. 안 그러면 "문서마다 카드 하나"가 두 배로 생긴다.
    /// </summary>
    public static List<string> PreferTextVersions(IEnumerable<string> files)
    {
        static int Rank(string f) => Path.GetExtension(f).ToLowerInvariant() switch
        {
            ".md" or ".markdown" => 0, ".txt" => 1, ".pdf" => 2, _ => 3
        };
        return files
            .GroupBy(f => Path.Combine(Path.GetDirectoryName(f) ?? "", Path.GetFileNameWithoutExtension(f)), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.OrderBy(Rank).First())
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// 요청 문장에서 실제로 있는 원문 경로(폴더, 또는 pdf·pptx·md·txt 파일)를 찾는다.
    /// 따옴표로 감싼 경로와, 따옴표 없이 드라이브 문자로 시작해 줄 끝까지 이어지는 경로를 본다.
    /// </summary>
    public static IReadOnlyList<string> MentionedSourcePaths(string text)
    {
        var candidates = new List<string>();
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(text, "[\"“”'](?<p>[A-Za-z]:\\\\[^\"“”'\\r\\n]+)[\"“”']"))
            candidates.Add(m.Groups["p"].Value.Trim());
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(text, @"(?<![""“”'\w])(?<p>[A-Za-z]:\\[^""“”'\r\n]+)"))
        {
            // 따옴표 없는 경로는 공백을 품을 수 있어, 뒤에서부터 잘라 가며 실제로 있는 가장 긴 경로를 고른다.
            var cut = m.Groups["p"].Value.TrimEnd();
            while (cut.Length > 3)
            {
                if (Directory.Exists(cut) || File.Exists(cut)) { candidates.Add(cut); break; }
                int space = cut.LastIndexOf(' ');
                if (space <= 2) break;
                cut = cut[..space].TrimEnd();
            }
        }
        return candidates
            .Where(p => Directory.Exists(p) || (File.Exists(p) && SourceLibrary.SupportedExtensions.Contains(Path.GetExtension(p))
                                                 && !p.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    public async Task<AgentTurnResult> RunAsync(
        string sourcePath, string instruction, IProgress<AgentEvent>? progress, CancellationToken ct)
    {
        sourcePath = Path.GetFullPath(sourcePath);
        if (Session is not null && !string.Equals(Session.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase))
            Discard();
        if (Session is null)
        {
            // 세션의 모든 LLM 호출(루프·요약·문안·길이 맞춤·이미지)을 모델별로 센다.
            _meter = new UsageMeter();
            _metered = new MeteredLlmClient(llm, _meter, options?.Model ?? "auto");
            Session = AgentFactory.StartSession(_metered, options, ppt, sourcePath);
            Session.Context.Meter = _meter;
        }
        var meterBefore = _meter!.Snapshot();
        Session.Context.Status = progress is null ? null
            : new Progress<string>(text => progress.Report(new AgentEvent(AgentEventKind.Message, text)));

        Session.Context.ExternalWriter = ExternalWriting;
        Session.Context.LlmFit = LlmFit;
        Session.Context.PendingCards = null;   // 새 요청이 오면 붙여넣기 대기는 취소한다

        // 요청에 원문 파일·폴더 경로를 적으면 첨부한다. 버튼을 쓰지 않고 경로를 적어 "도구가 없다"로 끝난 일이 있었다.
        var mentioned = MentionedSourcePaths(instruction)
            .Where(p => !_attached.Contains(Path.GetFullPath(p), StringComparer.OrdinalIgnoreCase)).ToList();
        if (mentioned.Count > 0)
        {
            Session.Context.Status?.Report($"요청에 적힌 경로의 원문을 첨부하는 중: {string.Join(", ", mentioned.Select(Path.GetFileName))}");
            var (files, pages) = await AttachAsync(mentioned, ct);
            Session.Context.Status?.Report($"원문 {files}개 · {pages}쪽 첨부");
        }
        bool hasSources = _attached.Count > 0;
        var scoped = hasSources ? library.WithScope(_attached) : null;
        Session.Context.Sources = scoped;
        var request = scoped is null
            ? instruction
            : $"{instruction}\n\n(첨부 원문: {string.Join(", ", scoped.Documents().Select(d => $"{d.Alias}={d.Name}"))})";

        var run = await AgentFactory.CreateLoop(_metered!, options, hasSources, LoopOptions).RunAsync(Session, request, progress, ct);

        var (before, after) = await RenderAsync(Session, ct);
        return new AgentTurnResult(run, before, after, Session.Context.Changes.ToList(),
            UsageMeter.Delta(meterBefore, _meter.Snapshot()));
    }

    /// <summary>작업 사본을 원본 옆 결과 파일로 남기고 세션을 끝낸다. 결과 파일 경로를 돌려준다.</summary>
    public string Approve()
    {
        var session = Session ?? throw new InvalidOperationException("승인할 작업이 없습니다.");
        var result = session.Approve();
        Discard();
        return result;
    }

    /// <summary>작업 사본과 대화를 버린다. 원본과 이미 승인한 결과 파일은 그대로다.</summary>
    public void Discard()
    {
        Session?.Discard();
        Session?.Dispose();
        Session = null;
    }

    public void Dispose() => Discard();
}
