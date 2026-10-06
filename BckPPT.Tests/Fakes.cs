using BckPPT.Llm;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Tests;

/// <summary>녹화한 응답을 순서대로 돌려주는 LLM. 받은 요청도 남겨 검증에 쓴다.</summary>
public sealed class FakeLlmClient(params object[] script) : ILlmClient
{
    // LlmCompletion은 응답으로, Exception은 그 차례에 던진다.
    private readonly Queue<object> _script = new(script);
    public List<LlmRequest> Requests { get; } = new();

    public static LlmCompletion Calls(params (string Name, string Args)[] calls) =>
        new("", calls.Select((c, i) => new LlmToolCall($"call_{Guid.NewGuid():N}"[..12] + i, c.Name, c.Args)).ToList(),
            "tool_calls", new LlmUsage(100, 20));

    public static LlmCompletion Text(string content) => new(content, [], "stop", new LlmUsage(100, 20));

    public Task<LlmCompletion> CompleteAsync(LlmRequest request, CancellationToken cancellationToken = default)
    {
        // 나중에 대화가 바뀌어도 당시 요청을 볼 수 있게 메시지 목록을 복사해 둔다.
        Requests.Add(request with { Messages = request.Messages.ToList() });
        if (_script.Count == 0) throw new InvalidOperationException("대본이 끝났습니다.");
        return _script.Dequeue() switch
        {
            Exception ex => Task.FromException<LlmCompletion>(ex),
            var c => Task.FromResult((LlmCompletion)c)
        };
    }

    public IAsyncEnumerable<LlmStreamEvent> StreamAsync(LlmRequest request, CancellationToken cancellationToken = default) =>
        throw new NotSupportedException();

    public Task TestConnectionAsync(CancellationToken cancellationToken = default) => Task.CompletedTask;
}

public sealed class FakeShape(string name, float left, float top, float width, float height, int charsPerLine, params string[] paragraphs)
{
    private static int _nextId = 100;
    public int Id { get; private init; } = Interlocked.Increment(ref _nextId);
    public string Name { get; set; } = name;
    public bool Is(string address) => Name == address || address == $"@{Id}" || (Group != "" && address == $"{Group}/{Name}");
    public float Left { get; set; } = left;
    public float Top { get; set; } = top;
    public float Width { get; set; } = width;
    public float Height { get; set; } = height;
    public int CharsPerLine { get; } = charsPerLine;
    public List<string> Paragraphs { get; } = paragraphs.ToList();
    public TextFormatInfo Format { get; set; } = new("맑은 고딕", 12f, 0f, false, "#000000");
    /// <summary>속한 그룹 이름(최상위면 빈 값).</summary>
    public string Group { get; set; } = "";
    /// <summary>추가할 때 받은 요청(TextShapeRequest·LineRequest·TableRequest).</summary>
    public object? Spec { get; set; }

    public FakeShape Clone() => new(Name, Left, Top, Width, Height, CharsPerLine, Paragraphs.ToArray())
        { Format = Format, Id = Id, Group = Group, Spec = Spec };

    /// <summary>글자 수로 자동 줄바꿈을 흉내 낸다. \v는 Shift+Enter, 문단 끝은 Enter.</summary>
    public IReadOnlyList<TextLineInfo> Lines()
    {
        if (Paragraphs.All(p => p.Length == 0)) return [];
        var lines = new List<TextLineInfo>();
        for (int p = 0; p < Paragraphs.Count; p++)
        {
            var segments = Paragraphs[p].Split('\v');
            for (int s = 0; s < segments.Length; s++)
            {
                var chunks = segments[s].Chunk(CharsPerLine).Select(c => new string(c)).DefaultIfEmpty("").ToList();
                for (int c = 0; c < chunks.Count; c++)
                {
                    var type = c < chunks.Count - 1 ? LineBreakType.AutoWrap
                        : s < segments.Length - 1 ? LineBreakType.ShiftEnter
                        : p < Paragraphs.Count - 1 ? LineBreakType.Enter
                        : LineBreakType.None;
                    lines.Add(new TextLineInfo(lines.Count + 1, chunks[c], type, FontSize: Format.Size));
                }
            }
        }
        return lines;
    }
}

/// <summary>메모리 안의 슬라이드로 IPptComService 중 에이전트 도구가 쓰는 부분만 흉내 낸다.</summary>
public sealed class FakePptComService : IPptComService
{
    public List<List<FakeShape>> Slides { get; } = new();
    public List<string> Calls { get; } = new();

    private List<FakeShape> Slide(int index) => Slides[index - 1];

    /// <summary>원본 파일(판정 기준)로 읽힐 상태. 작업 사본과 다른 경로로 읽으면 이 상태를 돌려준다.</summary>
    public string? OriginalPath { get; private set; }
    private List<List<FakeShape>>? _original;

    public void SnapshotOriginal(string path)
    {
        OriginalPath = path;
        _original = Slides.Select(s => s.Select(x => x.Clone()).ToList()).ToList();
    }

    private List<FakeShape> Read(string path, int index) =>
        _original is not null && path == OriginalPath ? _original[index - 1] : Slides[index - 1];
    private FakeShape Shape(int slide, string address) =>
        Slide(slide).FirstOrDefault(s => s.Is(address))
        ?? throw new InvalidOperationException($"{slide}쪽에서 도형 '{address}'를 찾지 못했습니다.");

    public bool IsPowerPointRunning => true;

    public PresentationAnalysis AnalyzePresentation(string presentationPath,
        IProgress<SlideAnalysisProgress>? slideProgress = null, CancellationToken cancellationToken = default) =>
        new(Slides.Select((s, i) => new SlideListItem(i + 1, s.FirstOrDefault()?.Paragraphs.FirstOrDefault() ?? "", s.Count, 0)).ToList(), [], []);

    public int DuplicateSlide(string presentationPath, int sourceSlideIndex, int insertAfterSlideIndex)
    {
        Calls.Add($"duplicate {sourceSlideIndex} after {insertAfterSlideIndex}");
        Slides.Insert(insertAfterSlideIndex, Slide(sourceSlideIndex).Select(s => s.Clone()).ToList());
        return insertAfterSlideIndex + 1;
    }

    public ParagraphEditResult SetParagraphTexts(string presentationPath, IReadOnlyList<ParagraphEdit> edits, bool markRed)
    {
        // 실제 구현과 같은 순서: 교체는 앞에서부터, 삭제·삽입은 큰 인덱스부터(같은 인덱스는 나열 순서).
        var ordered = edits.Where(e => e.Kind == ParagraphEditKind.Replace).OrderBy(e => e.ParagraphIndex)
            .Concat(edits.Where(e => e.Kind != ParagraphEditKind.Replace).OrderByDescending(e => e.ParagraphIndex));
        foreach (var edit in ordered)
        {
            Calls.Add($"{edit.Kind} {edit.SlideIndex}/{edit.ShapeName}/{edit.ParagraphIndex}");
            var paragraphs = Shape(edit.SlideIndex, edit.ShapeName).Paragraphs;
            switch (edit.Kind)
            {
                case ParagraphEditKind.Replace: paragraphs[edit.ParagraphIndex - 1] = edit.NewText; break;
                case ParagraphEditKind.Delete: paragraphs.RemoveAt(edit.ParagraphIndex - 1); break;
                case ParagraphEditKind.Insert: paragraphs.Insert(edit.ParagraphIndex, edit.NewText); break;
            }
        }
        return new ParagraphEditResult([], []);
    }

    public PptSlideInfo? GetShapeInfo(string presentationPath, int slideIndex, string shapeName)
    {
        var shape = Read(presentationPath, slideIndex).FirstOrDefault(s => s.Is(shapeName));
        if (shape is null) return null;
        var lines = shape.Lines();
        float bound = lines.Count * 10f;
        return new PptSlideInfo(slideIndex, shapeName, string.Join("\r", shape.Paragraphs), "", bound, shape.Height,
            lines.Count, bound > shape.Height + 1f, shape.Format.Size);
    }

    public IReadOnlyList<TextLineInfo> GetTextLineBreaks(string presentationPath, int slideIndex, string shapeName) =>
        Read(presentationPath, slideIndex).FirstOrDefault(s => s.Is(shapeName))?.Lines() ?? [];

    public IReadOnlyList<ShapeEntry> ListShapes(string presentationPath, int slideIndex) =>
        Read(presentationPath, slideIndex).Select(s => new ShapeEntry(s.Group == "" ? s.Name : $"{s.Group}/{s.Name}", s.Name, s.Group, 1, s.Left, s.Top, s.Width, s.Height,
            s.Paragraphs.Any(p => p.Length > 0), string.Join(" ", s.Paragraphs), s.Id)).ToList();

    public TextFormatInfo GetTextFormat(string presentationPath, int slideIndex, string address) =>
        Read(presentationPath, slideIndex).First(s => s.Is(address)).Format;

    public int ReplaceTextInShape(string presentationPath, int slideIndex, string address, string find, string replace)
    {
        var paragraphs = Shape(slideIndex, address).Paragraphs;
        int count = paragraphs.Sum(p => (p.Length - p.Replace(find, "").Length) / Math.Max(find.Length, 1));
        for (int i = 0; i < paragraphs.Count; i++) paragraphs[i] = paragraphs[i].Replace(find, replace);
        return count;
    }

    public void RenameShape(string presentationPath, int slideIndex, string address, string newName) => Shape(slideIndex, address).Name = newName;

    public void DeleteShape(string presentationPath, int slideIndex, string address)
    {
        // 그룹 이름이면 그룹 통째로 지운다.
        if (!Slide(slideIndex).Any(s => s.Is(address)) && Slide(slideIndex).Any(s => s.Group == address))
        {
            Slide(slideIndex).RemoveAll(s => s.Group == address);
            return;
        }
        Slide(slideIndex).Remove(Shape(slideIndex, address));
    }

    public IReadOnlyList<string> AddTextShapes(string presentationPath, int slideIndex, IReadOnlyList<TextShapeRequest> requests) =>
        requests.Select(r => AddTextShape(presentationPath, slideIndex, r)).ToList();

    public string AddLine(string presentationPath, int slideIndex, LineRequest request)
    {
        var name = request.Name ?? $"Line {Slide(slideIndex).Count + 1}";
        Slide(slideIndex).Add(new FakeShape(name, Math.Min(request.X1, request.X2), Math.Min(request.Y1, request.Y2),
            Math.Abs(request.X2 - request.X1), Math.Abs(request.Y2 - request.Y1), 100) { Spec = request });
        return name;
    }

    public TableLayout AddTable(string presentationPath, int slideIndex, TableRequest request)
    {
        var name = request.Name ?? $"Table {Slide(slideIndex).Count + 1}";
        int cols = request.Cells.Max(r => r.Count);
        var heights = Enumerable.Range(0, request.Cells.Count)
            .Select(r => request.RowHeights is { } h && r < h.Count ? h[r] : request.RowHeight).ToList();
        var widths = Enumerable.Range(0, cols)
            .Select(c => request.ColumnWidths is { } w && c < w.Count ? w[c] : request.Width / cols).ToList();
        Slide(slideIndex).Add(new FakeShape(name, request.Left, request.Top, widths.Sum(), heights.Sum(), 100,
            request.Cells.Select(r => string.Join("|", r)).ToArray()) { Spec = request });
        var tops = heights.Select((_, i) => request.Top + heights.Take(i).Sum()).ToList();
        var lefts = widths.Select((_, i) => request.Left + widths.Take(i).Sum()).ToList();
        return new TableLayout(name, tops, heights, lefts, widths);
    }

    public string GroupShapes(string presentationPath, int slideIndex, IReadOnlyList<string> addresses, string? name)
    {
        name ??= $"Group {Slide(slideIndex).Count + 1}";
        foreach (var address in addresses) Shape(slideIndex, address).Group = name;
        return name;
    }

    public IReadOnlyList<string> UngroupShape(string presentationPath, int slideIndex, string address)
    {
        var members = Slide(slideIndex).Where(s => s.Group == address).ToList();
        foreach (var s in members) s.Group = "";
        return members.Select(s => s.Name).ToList();
    }

    public int ClearSlide(string presentationPath, int slideIndex, IReadOnlyCollection<string> keep) =>
        Slide(slideIndex).RemoveAll(s => !keep.Any(k => s.Is(k)) && !(s.Group != "" && keep.Contains(s.Group)));

    public void SetShapeGeometry(string presentationPath, int slideIndex, string address, ShapeGeometry geometry)
    {
        var shape = Shape(slideIndex, address);
        shape.Left = geometry.Left ?? shape.Left;
        shape.Top = geometry.Top ?? shape.Top;
        shape.Width = geometry.Width ?? shape.Width;
        shape.Height = geometry.Height ?? shape.Height;
    }

    public string AddTextShape(string presentationPath, int slideIndex, TextShapeRequest request)
    {
        var name = request.Name ?? $"TextBox {Slide(slideIndex).Count + 1}";
        Slide(slideIndex).Add(new FakeShape(name, request.Left, request.Top, request.Width, request.Height, 100, request.Text) { Spec = request });
        return name;
    }

    public Task<string?> GetSlideThumbnailAsync(string presentationPath, int slideIndex) => Task.FromResult<string?>(null);

    // ── 에이전트 도구가 쓰지 않는 멤버 ──
    public IReadOnlyList<PptPresentation> GetOpenPresentations() => [];
    public MarkdownExportResult ExportMarkdown(string presentationPath, string outputPath, CancellationToken cancellationToken = default) => throw new NotSupportedException();
    public Task<AutoFitResult> AutoFitAsync(string presentationPath, IProgress<string>? progress = null) => throw new NotSupportedException();
    public Task FixShapeAsync(string presentationPath, int slideIndex, string shapeName, IProgress<string>? progress = null) => throw new NotSupportedException();
    public Task NavigateToShapeAsync(string presentationPath, int slideIndex, string shapeName) => throw new NotSupportedException();
    public void NavigateToSlide(string presentationPath, int slideIndex) => throw new NotSupportedException();
    public void OpenPptxInPowerPoint(string filePath) => throw new NotSupportedException();
    public IDisposable KeepOpen(string presentationPath, bool readOnly) => new MemoryStream();
    public IReadOnlyList<(int Index, string Title)> ListSlides(string presentationPath) =>
        Slides.Select((s, i) => (i + 1, s.FirstOrDefault()?.Paragraphs.FirstOrDefault() ?? "")).ToList();
    public int ReplaceText(string presentationPath, string find, string replace, IReadOnlyList<(int slideIndex, string shapeName)>? targets = null) => throw new NotSupportedException();
    public int NormalizeFonts(string presentationPath) => throw new NotSupportedException();
    public void ReplaceImage(string presentationPath, int slideIndex, float targetLeft, float targetTop, string newImagePath) => throw new NotSupportedException();
    public void SetShapeText(string presentationPath, int slideIndex, string shapeName, string newText) => throw new NotSupportedException();
    public int ImportSlide(string sourcePath, int sourceSlideIndex, string targetPath, int insertAfterSlideIndex) => throw new NotSupportedException();
    public IReadOnlyList<(int SlideIndex, string ShapeName, string OldText, string NewText)> ApplyRules(string presentationPath, IReadOnlyList<ReplaceRule> rules) => throw new NotSupportedException();
    public void RestoreParagraphColors(string presentationPath, IReadOnlyList<ParagraphColorSnapshot> snapshots) => throw new NotSupportedException();
    public SlideLayout GetSlideLayout(string presentationPath, int slideIndex) => throw new NotSupportedException();
    public IReadOnlyList<RecentPptxFile> GetRecentFiles() => [];
    public int SetTextColor(string presentationPath, int slideIndex, string address, string find, string color) => throw new NotSupportedException();
    public void SetTextFormat(string presentationPath, int slideIndex, string address, TextFormatPatch patch) => throw new NotSupportedException();
}
