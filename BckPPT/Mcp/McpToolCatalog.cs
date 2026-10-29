using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using BckPPT.Agent.Tools;
using BckPPT.Llm;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Mcp;

/// <summary>MCP 도구 결과 한 조각. 글 또는 그림(base64 PNG).</summary>
public sealed record McpContent(string Type, string? Text = null, string? Data = null, string? MimeType = null)
{
    public static McpContent FromText(string text) => new("text", Text: text);
    public static McpContent FromImage(byte[] png) => new("image", Data: Convert.ToBase64String(png), MimeType: "image/png");
}

public sealed record McpToolResult(IReadOnlyList<McpContent> Content, bool IsError)
{
    public static McpToolResult Ok(string text) => new([McpContent.FromText(text)], false);
    public static McpToolResult Error(string text) => new([McpContent.FromText(text)], true);
}

public sealed record McpTool(string Name, string Description, JsonObject InputSchema,
    Func<JsonElement, CancellationToken, Task<McpToolResult>> Call);

/// <summary>
/// 외부 에이전트(Claude Code 등)에 내보내는 도구. 판단은 그 에이전트가 하고 BckPPT는 PowerPoint를 정확히 다룬다.
/// 내부 에이전트의 PPT 도구를 그대로 쓰되, 대상 문서를 고르는 <c>file</c> 인자를 붙인다.
/// 열린 창의 문서는 그 자리에서 고치고 저장하지 않는다(저장은 사용자가 한다). 창이 없는 파일은 고친 뒤 저장한다.
/// LLM이 필요한 도구(카드 채우기·원문 요약)와 내부 루프 전용 도구(finish)는 내보내지 않는다.
/// </summary>
public sealed class McpToolCatalog
{
    private static readonly HashSet<string> Excluded = new(StringComparer.Ordinal) { "finish", "fill_cards", "render_slide" };
    private static readonly string[] PresentationExtensions = [".pptx", ".pptm", ".ppt"];

    private const string FileDescription =
        "대상 PPT. 절대 경로나 PowerPoint에 열린 문서 이름 조각. 비우면 PowerPoint에서 지금 보고 있는 문서";

    private readonly IPptComService _ppt;
    private readonly Dictionary<string, McpTool> _tools = new(StringComparer.Ordinal);
    private readonly ToolRegistry _agentTools;

    public McpToolCatalog(IPptComService ppt)
    {
        _ppt = ppt;
        _agentTools = new ToolRegistry(PptAgentTools.Create(includeSources: false).Where(t => !Excluded.Contains(t.Name)));

        Add("list_presentations", "PowerPoint 창에 열린 문서 목록(이름·경로·쪽 수·보고 있는 문서 표시)",
            Schema(), (_, _) => Task.FromResult(ListPresentations()));
        Add("open_presentation", "PPT 파일을 PowerPoint 창에 연다",
            Schema(("path", "string", "PPT 절대 경로")), OpenPresentation, "path");
        Add("go_to_slide", "PowerPoint 창을 그 쪽으로 옮겨 사용자에게 보여 준다",
            WithFile(Schema(("slide", "integer", "쪽 번호(1부터)")), "slide"), GoToSlide);
        Add("list_fonts",
            "문서에 쓰인 글꼴 목록(글꼴마다 쓰인 횟수·쪽). font_name을 주면 그 글꼴을 쓰는 쪽·도형·글 앞부분을 모두 찾아 준다",
            WithFile(Schema(("font_name", "string", "찾을 글꼴 이름(정확히). 비우면 목록"))), ListFonts);
        Add("export_markdown", "문서를 슬라이드별 Markdown으로 내보낸다(쪽마다 PNG 미리보기를 _assets 폴더에 함께 저장)",
            WithFile(Schema(("output_path", "string", "저장할 .md 절대 경로. 비우면 PPT 옆에 같은 이름으로"))), ExportMarkdown);
        Add("render_slide", "슬라이드를 PNG로 렌더해 그림으로 돌려준다. 고친 뒤 넘침·겹침을 눈으로 확인할 때 쓴다",
            WithFile(Schema(("slide", "integer", "쪽 번호(1부터)")), "slide"), RenderSlide);

        foreach (var tool in _agentTools.Tools)
        {
            var schema = JsonNode.Parse(tool.ParametersSchema)!.AsObject();
            var name = tool.Name;
            _tools[name] = new McpTool(name, tool.Description + (tool.Writes ? " (문서를 고친다)" : ""),
                WithFile(schema), (args, ct) => CallAgentTool(name, tool.Writes, args, ct));
        }
    }

    public IReadOnlyCollection<McpTool> Tools => _tools.Values;

    public async Task<McpToolResult> CallAsync(string name, JsonElement arguments, CancellationToken cancellationToken)
    {
        if (!_tools.TryGetValue(name, out var tool))
            return McpToolResult.Error($"'{name}' 도구는 없습니다.");
        if (arguments.ValueKind != JsonValueKind.Object)
            return McpToolResult.Error("인자는 JSON 객체여야 합니다.");
        try { return await tool.Call(arguments, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) when (ex.GetBaseException() is PptUserInterventionException)
        {
            return McpToolResult.Error("사용자가 PowerPoint에서 도형이나 글을 선택해 작업을 멈췄습니다. 사용자에게 확인한 뒤 다시 부르세요.");
        }
        catch (AgentToolException ex) { return McpToolResult.Error(ex.Message); }
        catch (Exception ex) { return McpToolResult.Error($"{name} 실행 실패: {ex.GetBaseException().Message}"); }
    }

    // ── 대상 문서 ────────────────────────────────────────────

    /// <summary>file 인자를 문서 경로로 바꾼다. 비우면 보고 있는 문서, 경로면 그 파일, 아니면 열린 문서 이름 조각.</summary>
    public string ResolveFile(JsonElement args) => ResolveFile(ToolArgs.OptionalString(args, "file"));

    public string ResolveFile(string? file)
    {
        file = file?.Trim().Trim('"');
        if (string.IsNullOrEmpty(file))
            return _ppt.GetActivePresentation()?.Path is { Length: > 0 } active
                ? active
                : throw new AgentToolException("PowerPoint에서 보고 있는 문서가 없습니다. file에 PPT 경로를 주거나 open_presentation으로 여세요.");

        if (Path.IsPathFullyQualified(file))
        {
            var full = Path.GetFullPath(file);
            if (!PresentationExtensions.Contains(Path.GetExtension(full), StringComparer.OrdinalIgnoreCase))
                throw new AgentToolException("PPT 파일(.pptx·.pptm·.ppt)이 아닙니다.");
            return File.Exists(full) ? full : throw new AgentToolException($"파일이 없습니다: {full}");
        }

        var open = _ppt.ListOpenWindows().Where(o => o.IsSaved).ToList();
        var hits = open.Where(o => Path.GetFileName(o.Path).Contains(file, StringComparison.OrdinalIgnoreCase)).ToList();
        var exact = hits.Where(o => string.Equals(Path.GetFileNameWithoutExtension(o.Path), file, StringComparison.OrdinalIgnoreCase)).ToList();
        if (exact.Count == 1) return exact[0].Path;
        return hits.Count switch
        {
            1 => hits[0].Path,
            0 => throw new AgentToolException($"'{file}' 이름이 든 열린 문서가 없습니다. 열린 문서: " +
                                              (open.Count == 0 ? "없음" : string.Join(", ", open.Select(o => o.Name)))),
            _ => throw new AgentToolException($"'{file}' 이름이 든 문서가 여럿입니다: {string.Join(", ", hits.Select(o => o.Name))}. 더 구체적으로 주세요.")
        };
    }

    private async Task<McpToolResult> CallAgentTool(string name, bool writes, JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        // 창에 열린 문서는 사용자가 보고 있으니 저장하지 않고 고친다. 그동안 사용자가 도형을 고르면 멈춘다.
        using var watch = writes && _ppt.IsOpenInWindow(path) ? _ppt.WatchUserActivity(path) : null;
        var outcome = await _agentTools.ExecuteAsync(new AgentToolContext(_ppt, path),
            new LlmToolCall(Guid.NewGuid().ToString("N"), name, args.GetRawText()), ct);
        return new([McpContent.FromText(outcome.Content)], outcome.IsError);
    }

    // ── MCP 전용 도구 ────────────────────────────────────────

    private McpToolResult ListPresentations()
    {
        var open = _ppt.ListOpenWindows();
        if (open.Count == 0) return McpToolResult.Ok("PowerPoint에 열린 문서가 없습니다.");
        var sb = new StringBuilder($"열린 문서 {open.Count}개\n");
        foreach (var o in open)
            sb.AppendLine($"{(o.IsActive ? "▶" : " ")} {o.Name} · {o.SlideCount}쪽 · {(o.IsSaved ? o.Path : "저장 안 된 새 문서")}");
        if (_ppt.GetActivePresentation() is { CurrentSlide: int current } active)
            sb.AppendLine($"보고 있는 쪽: {active.Name} {current}쪽");
        return McpToolResult.Ok(sb.ToString().TrimEnd());
    }

    private Task<McpToolResult> OpenPresentation(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(ToolArgs.OptionalString(args, "path")
                               ?? throw new AgentToolException("인자 'path'가 필요합니다."));
        _ppt.OpenPptxInPowerPoint(path);
        return Task.FromResult(McpToolResult.Ok($"열었습니다: {path}"));
    }

    private Task<McpToolResult> GoToSlide(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        _ppt.NavigateToSlide(path, slide);
        return Task.FromResult(McpToolResult.Ok($"{Path.GetFileName(path)} {slide}쪽으로 옮겼습니다."));
    }

    private async Task<McpToolResult> ListFonts(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        var fontName = ToolArgs.OptionalString(args, "font_name")?.Trim();
        // 전체 분석(도형마다 줄 측정)은 큰 문서에서 COM 45초 제한을 넘는다. 글꼴만 쪽마다 따로 읽는다.
        var fonts = await Task.Run(() =>
        {
            using var keep = _ppt.IsOpenInWindow(path) ? null : _ppt.KeepOpen(path, readOnly: true);
            var all = new List<PptFontInfo>();
            int count = _ppt.ListSlides(path).Count;
            for (int slide = 1; slide <= count; slide++)
            {
                ct.ThrowIfCancellationRequested();
                all.AddRange(_ppt.ListSlideFonts(path, slide));
            }
            return all;
        }, ct);
        var sb = new StringBuilder();
        if (string.IsNullOrEmpty(fontName))
        {
            var groups = fonts.GroupBy(f => f.FontName).OrderByDescending(g => g.Count()).ToList();
            sb.AppendLine($"글꼴 {groups.Count}종");
            foreach (var g in groups)
            {
                var slides = g.Select(f => f.SlideIndex).Distinct().Order().ToList();
                sb.AppendLine($"- {g.Key}: {g.Count()}곳 · {slides.Count}쪽 ({string.Join(",", slides.Take(30))}{(slides.Count > 30 ? ",…" : "")})" +
                              (g.Any(f => !f.IsAllowed) ? " · 허용 목록 밖" : ""));
            }
            return McpToolResult.Ok(sb.ToString().TrimEnd());
        }

        var hits = fonts.Where(f => string.Equals(f.FontName, fontName, StringComparison.OrdinalIgnoreCase)).ToList();
        if (hits.Count == 0)
            return McpToolResult.Ok($"'{fontName}' 글꼴을 쓰는 곳이 없습니다. 쓰인 글꼴: {string.Join(", ", fonts.Select(f => f.FontName).Distinct())}");
        sb.AppendLine($"'{fontName}' {hits.Count}곳");
        foreach (var f in hits.OrderBy(f => f.SlideIndex))
            sb.AppendLine($"- {f.SlideIndex}쪽 [{f.ShapeName}] {f.FontSize}pt{(f.IsBold ? " 굵게" : "")} · {f.TextPreview}");
        return McpToolResult.Ok(sb.ToString().TrimEnd());
    }

    private async Task<McpToolResult> ExportMarkdown(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        var output = ToolArgs.OptionalString(args, "output_path") is { Length: > 0 } o ? o : Path.ChangeExtension(path, ".md");
        if (!Path.IsPathFullyQualified(output)) throw new AgentToolException("output_path는 절대 경로여야 합니다.");
        var result = await Task.Run(() => _ppt.ExportMarkdown(path, output, ct), ct);
        return McpToolResult.Ok($"{result.SlideCount}쪽을 내보냈습니다: {result.OutputPath}" +
                                (result.WarningCount > 0 ? $" (경고 {result.WarningCount}건)" : ""));
    }

    private async Task<McpToolResult> RenderSlide(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        var png = await _ppt.GetSlideThumbnailAsync(path, slide) ?? throw new AgentToolException($"{slide}쪽을 렌더하지 못했습니다.");
        try
        {
            return new([McpContent.FromText($"{Path.GetFileName(path)} {slide}쪽"), McpContent.FromImage(await File.ReadAllBytesAsync(png, ct))], false);
        }
        finally { try { File.Delete(png); } catch { } }
    }

    // ── 스키마 ──────────────────────────────────────────────

    private void Add(string name, string description, JsonObject schema,
        Func<JsonElement, CancellationToken, Task<McpToolResult>> call, params string[] required)
    {
        if (required.Length > 0) schema["required"] = new JsonArray(required.Select(r => (JsonNode)r).ToArray());
        _tools[name] = new McpTool(name, description, schema, call);
    }

    private static JsonObject Schema(params (string Name, string Type, string Description)[] properties)
    {
        var props = new JsonObject();
        foreach (var (name, type, description) in properties)
            props[name] = new JsonObject { ["type"] = type, ["description"] = description };
        return new JsonObject { ["type"] = "object", ["properties"] = props };
    }

    private static JsonObject WithFile(JsonObject schema, params string[] required)
    {
        if (schema["properties"] is not JsonObject props) schema["properties"] = props = new JsonObject();
        props["file"] = new JsonObject { ["type"] = "string", ["description"] = FileDescription };
        if (required.Length > 0) schema["required"] = new JsonArray(required.Select(r => (JsonNode)r).ToArray());
        return schema;
    }
}
