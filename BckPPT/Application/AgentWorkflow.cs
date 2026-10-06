using System.IO;
using BckPPT.Agent;
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
    IReadOnlyList<string> Changes);

/// <summary>
/// 채팅창과 도구 루프 사이. 대상 PPT마다 세션(작업 사본 + 대화)을 하나 두고, 같은 대상에 대한 후속 요청은
/// 그 세션에서 이어간다. 승인하면 결과 사본을 남기고 세션을 끝낸다. 대상이 바뀌거나 버리면 작업 사본을 지운다.
/// </summary>
public sealed class AgentWorkflow(
    ILlmClient llm, GitokyLlmOptions? options, IPptComService ppt, RenderValidator renderer, SourceLibrary library)
    : IDisposable
{
    private readonly List<string> _attached = new();

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
                var found = Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories)
                    .Where(f => SourceLibrary.SupportedExtensions.Contains(Path.GetExtension(f)))
                    .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal))
                    .ToList();
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

    public async Task<AgentTurnResult> RunAsync(
        string sourcePath, string instruction, IProgress<AgentEvent>? progress, CancellationToken ct)
    {
        sourcePath = Path.GetFullPath(sourcePath);
        if (Session is not null && !string.Equals(Session.SourcePath, sourcePath, StringComparison.OrdinalIgnoreCase))
            Discard();
        Session ??= AgentFactory.StartSession(llm, options, ppt, sourcePath);

        bool hasSources = _attached.Count > 0;
        Session.Context.Sources = hasSources ? library.WithScope(_attached) : null;
        var request = hasSources
            ? $"{instruction}\n\n(첨부 원문: {string.Join(", ", _attached.Select(Path.GetFileName))})"
            : instruction;

        var run = await AgentFactory.CreateLoop(llm, options, hasSources, LoopOptions).RunAsync(Session, request, progress, ct);

        var touched = Session.Context.Origins.Touched.ToHashSet();
        var after = touched.Count == 0 ? [] : await renderer.RenderSlidesAsync(Session.WorkingPath, touched, ct);
        var originals = touched.Select(Session.Context.Origins.OriginalIndexOf).ToHashSet();
        var before = originals.Count == 0 ? [] : await renderer.RenderSlidesAsync(sourcePath, originals, ct);
        return new AgentTurnResult(run, before, after, Session.Context.Changes.ToList());
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
