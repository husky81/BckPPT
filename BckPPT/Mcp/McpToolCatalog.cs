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
        Add("open_presentation",
            "PPT 파일을 PowerPoint 창에 열고 문서가 올라올 때까지 기다린다(큰 파일도 다음 도구가 바로 찾게). " +
            "메일·인터넷에서 받은 파일이 「제한된 보기」로 열리면 결과에 알린다 — 그 상태에서는 다른 도구가 문서를 못 찾으니, " +
            "사용자에게 확인받고 unblock=true로 다시 부르면 받은 파일 표시를 지우고 편집 모드로 다시 연다",
            Schema(("path", "string", "PPT 절대 경로"),
                ("unblock", "boolean", "true면 인터넷에서 받은 파일 표시(Zone.Identifier)를 지우고, 제한된 보기 창이 있으면 닫고 편집 모드로 연다"),
                ("wait_seconds", "integer", "문서가 올라오길 기다릴 최대 초(기본 60)")), OpenPresentation, "path");
        Add("go_to_slide", "PowerPoint 창을 그 쪽으로 옮겨 사용자에게 보여 준다",
            WithFile(Schema(("slide", "integer", "쪽 번호(1부터)")), "slide"), GoToSlide);
        Add("list_fonts",
            "문서에 쓰인 글꼴 목록(글꼴마다 쓰인 횟수·쪽). font_name을 주면 그 글꼴을 쓰는 쪽·도형·글 앞부분을 모두 찾아 준다",
            WithFile(Schema(("font_name", "string", "찾을 글꼴 이름(정확히). 비우면 목록"))), ListFonts);
        Add("export_markdown", "문서를 슬라이드별 Markdown으로 내보낸다(쪽마다 PNG 미리보기를 _assets 폴더에 함께 저장)",
            WithFile(Schema(("output_path", "string", "저장할 .md 절대 경로. 비우면 PPT 옆에 같은 이름으로"))), ExportMarkdown);
        Add("render_slide", "슬라이드를 PNG로 렌더해 그림으로 돌려준다. 고친 뒤 넘침·겹침을 눈으로 확인할 때 쓴다",
            WithFile(Schema(("slide", "integer", "쪽 번호(1부터)")), "slide"), RenderSlide);
        Add("render_slides",
            "여러 쪽을 한 번에 렌더한다. slides(쪽 번호 배열) 또는 from~to. sheet=true면 여러 쪽을 한 장에 모은 축소판(쪽 번호 표시)으로 돌려준다. " +
            $"따로 받는 그림은 {MaxSeparateRenders}쪽까지, 넘으면 축소판으로 모은다(최대 {MaxSheetRenders}쪽)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slides"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "integer" }, ["description"] = "쪽 번호들(1부터)" },
                    ["from"] = new JsonObject { ["type"] = "integer", ["description"] = "slides 대신 범위의 첫 쪽" },
                    ["to"] = new JsonObject { ["type"] = "integer", ["description"] = "범위의 끝 쪽" },
                    ["sheet"] = new JsonObject { ["type"] = "boolean", ["description"] = "true면 한 장에 모은 축소판" },
                    ["columns"] = new JsonObject { ["type"] = "integer", ["description"] = "축소판 한 줄의 쪽 수(기본 4)" },
                },
            }), RenderSlides);

        var cropSchema = new JsonObject
        {
            ["type"] = "object",
            ["description"] = "원본 그림 기준으로 먼저 잘라 낼 비율(0~1). 핵심 부분만 크게 보이게 할 때. 예: {\"left\":0.2,\"right\":0.1}",
            ["properties"] = new JsonObject
            {
                ["left"] = new JsonObject { ["type"] = "number" }, ["top"] = new JsonObject { ["type"] = "number" },
                ["right"] = new JsonObject { ["type"] = "number" }, ["bottom"] = new JsonObject { ["type"] = "number" },
            },
        };
        JsonObject FitSchema(string note) => new()
        {
            ["type"] = "string", ["enum"] = new JsonArray("cover", "contain", "stretch"),
            ["description"] = "cover: 칸을 채우고 넘는 쪽을 가운데 기준으로 자른다. contain: 비율 유지, 칸 안 가운데. stretch: 칸 크기로 늘린다. " + note,
        };
        Add("add_picture",
            "그림 파일을 쪽에 넣는다(문서에 포함). 칸(left·top·width·height, pt)에 fit으로 맞춘다. 열린 창의 문서면 그 자리에 넣고 저장은 사용자가 한다. " +
            "결과에 도형 주소·실제 자리·크기·자른 양·원본 픽셀을 준다 (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["path"] = new JsonObject { ["type"] = "string", ["description"] = "그림 파일 절대 경로(png·jpg·gif·bmp·tif·emf·wmf·svg)" },
                    ["left"] = new JsonObject { ["type"] = "number" }, ["top"] = new JsonObject { ["type"] = "number" },
                    ["width"] = new JsonObject { ["type"] = "number" }, ["height"] = new JsonObject { ["type"] = "number" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["crop"] = cropSchema.DeepClone(),
                    ["name"] = new JsonObject { ["type"] = "string", ["description"] = "도형 이름" },
                },
            }, "slide", "path", "left", "top", "width", "height"), AddPicture);
        Add("add_pictures_in_box",
            "칸 하나를 그림 수만큼 나눠(가로로 긴 칸은 좌우, 세로로 긴 칸은 위아래) gap 간격으로 나란히 넣는다. " +
            "box는 {left,top,width,height} 또는 기존 도형 주소(그 도형 자리·크기를 쓴다). replace_placeholder=true면 그 도형(예: 「[그림 삽입]」 글상자)을, " +
            "도형 주소면 그 도형을 지우고 그 Z순서 자리에 넣는다 (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["paths"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "그림 파일 절대 경로들(넣는 순서: 왼쪽→오른쪽 또는 위→아래)" },
                    ["box"] = new JsonObject { ["description"] = "칸: {left,top,width,height}(pt) 또는 list_shapes의 도형 주소" },
                    ["gap"] = new JsonObject { ["type"] = "number", ["description"] = "그림 사이 간격 pt(기본 3)" },
                    ["direction"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("auto", "horizontal", "vertical"), ["description"] = "나누는 방향(기본 auto)" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["crops"] = new JsonObject { ["type"] = "array", ["items"] = cropSchema.DeepClone(), ["description"] = "그림마다 crop(paths 순서). 없으면 null" },
                    ["names"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "그림마다 도형 이름" },
                    ["replace_placeholder"] = new JsonObject
                    {
                        ["description"] = "true면 box 도형을, 도형 주소(문자열)면 그 도형을 지우고 그 Z순서 자리에 넣는다. " +
                                          "「[그림 삽입]」 글상자가 글에 맞춰 줄어 칸보다 작으면 box에는 칸 도형, 여기에는 글상자 주소를 준다",
                    },
                },
            }, "slide", "paths", "box"), AddPicturesInBox);
        Add("replace_picture",
            "도형(보통 그림)을 다른 그림 파일로 바꾼다. 같은 자리·크기·Z순서를 지키고 이름도 이어받는다. fit 기본 cover (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["shape"] = new JsonObject { ["type"] = "string", ["description"] = "바꿀 도형 주소(list_shapes의 값)" },
                    ["path"] = new JsonObject { ["type"] = "string", ["description"] = "새 그림 파일 절대 경로" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["crop"] = cropSchema.DeepClone(),
                    ["name"] = new JsonObject { ["type"] = "string", ["description"] = "새 이름. 비우면 원래 이름" },
                },
            }, "slide", "shape", "path"), ReplacePicture);
        Add("import_slides",
            "다른 PPT 파일의 쪽을 이 문서 after 쪽 뒤에 붙인다(클립보드를 쓰지 않는다). 붙인 쪽은 이 문서의 디자인(테마)을 따른다. " +
            "source가 창에 열려 있어도 디스크에 저장된 판을 읽으니, 고친 내용이 있으면 사용자에게 먼저 저장해 달라고 한다 (문서를 고친다)",
            WithFile(Schema(
                ("source", "string", "가져올 PPT. 절대 경로나 열린 문서 이름 조각"),
                ("from", "integer", "가져올 첫 쪽(1부터)"),
                ("to", "integer", "가져올 끝 쪽. 비우면 from 한 쪽"),
                ("after", "integer", "이 문서에서 이 쪽 뒤에 붙인다(0이면 맨 앞)")), "source", "from", "after"), ImportSlides);
        Add("delete_slides",
            "쪽들을 지운다. 지운 쪽의 번호·제목을 돌려준다. 지우기 전에 list_slides로 번호를 확인하고 사용자에게 어느 쪽을 지울지 확인받는다. " +
            "창에 열리지 않은 파일은 바로 저장되므로 confirm=true가 있어야 지운다 (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slides"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "integer" }, ["description"] = "지울 쪽 번호들(1부터, 지우기 전 번호)" },
                    ["confirm"] = new JsonObject { ["type"] = "boolean", ["description"] = "창에 열리지 않은 파일을 지울 때 사용자 확인을 받았으면 true" },
                },
            }, "slides"), DeleteSlides);
        Add("move_slide", "쪽을 옮긴다. to는 옮긴 뒤 그 쪽의 번호 (문서를 고친다)",
            WithFile(Schema(("slide", "integer", "옮길 쪽(1부터)"), ("to", "integer", "옮긴 뒤 번호")), "slide", "to"), MoveSlide);

        foreach (var tool in _agentTools.Tools)
        {
            var schema = JsonNode.Parse(tool.ParametersSchema)!.AsObject();
            var name = tool.Name;
            _tools[name] = new McpTool(name, tool.Description + (tool.Writes ? " (문서를 고친다)" : ""),
                WithFile(schema), (args, ct) => CallAgentTool(name, tool.Writes, args, ct));
        }
    }

    public IReadOnlyCollection<McpTool> Tools => _tools.Values;

    /// <summary>도구를 부르고 진단 로그에 호출(도구·파일·시간·오류)과 그 뒤 PowerPoint 상태를 남긴다.</summary>
    public async Task<McpToolResult> CallAsync(string name, JsonElement arguments, CancellationToken cancellationToken)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        string file = arguments.ValueKind == JsonValueKind.Object ? ToolArgs.OptionalString(arguments, "file") ?? "(보고 있는 문서)" : "";
        McpToolResult? result = null;
        try { return result = await CallCoreAsync(name, arguments, cancellationToken); }
        finally
        {
            string outcome = result is null ? "취소" : result.IsError ? "오류: " + Truncate(result.Content.FirstOrDefault()?.Text) : "ok";
            PptDiagnosticLog.Write("mcp", $"{name} file={file} {watch.ElapsedMilliseconds}ms {outcome}");
            string state;
            try { state = await Task.Run(_ppt.DescribeState); }
            catch (Exception ex) { state = "상태 읽기 실패: " + ex.GetBaseException().Message; }
            PptDiagnosticLog.Write("state", state);
        }

        static string Truncate(string? text) => text is null ? "" : text.Length <= 200 ? text : text[..200] + "…";
    }

    private async Task<McpToolResult> CallCoreAsync(string name, JsonElement arguments, CancellationToken cancellationToken)
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

    /// <summary>open_presentation 직후에는 문서가 아직 안 보일 수 있어 잠깐 다시 찾는다.</summary>
    private DateTime _lastOpen = DateTime.MinValue;
    public static TimeSpan RecentOpenWindow { get; set; } = TimeSpan.FromSeconds(120);
    public static TimeSpan RetryAfterOpen { get; set; } = TimeSpan.FromSeconds(10);

    public string ResolveFile(string? file)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        while (true)
        {
            try { return ResolveFileOnce(file); }
            catch (AgentToolException) when (DateTime.UtcNow - _lastOpen < RecentOpenWindow && watch.Elapsed < RetryAfterOpen)
            {
                Thread.Sleep(500);
            }
            catch (AgentToolException ex) when (ProtectedViewHint(file) is { } hint)
            {
                throw new AgentToolException(ex.Message + " " + hint);
            }
        }
    }

    private string? ProtectedViewHint(string? file)
    {
        IReadOnlyList<ProtectedViewInfo> views;
        try { views = _ppt.ListProtectedViews(); } catch { return null; }
        file = file?.Trim().Trim('"');
        var hits = views.Where(v => string.IsNullOrEmpty(file) || v.Name.Contains(Path.GetFileName(file), StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(v.Path, file, StringComparison.OrdinalIgnoreCase)).ToList();
        return hits.Count == 0 ? null
            : $"「제한된 보기」로 열린 창이 있습니다({string.Join(", ", hits.Select(v => v.Name))}). 이 창의 문서는 도구가 읽지 못합니다. " +
              "사용자에게 확인받고 open_presentation(path, unblock=true)로 편집 모드로 다시 여세요.";
    }

    private string ResolveFileOnce(string? file)
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
        if (!writes)
        {
            var read = await _agentTools.ExecuteAsync(new AgentToolContext(_ppt, path),
                new LlmToolCall(Guid.NewGuid().ToString("N"), name, args.GetRawText()), ct);
            return new([McpContent.FromText(read.Content)], read.IsError);
        }
        return await WriteAsync(path, async () =>
        {
            var outcome = await _agentTools.ExecuteAsync(new AgentToolContext(_ppt, path),
                new LlmToolCall(Guid.NewGuid().ToString("N"), name, args.GetRawText()), ct);
            return new McpToolResult([McpContent.FromText(outcome.Content)], outcome.IsError);
        });
    }

    /// <summary>
    /// 쓰기 도구 공통. 창에 열린 문서는 사용자가 보고 있으니 저장하지 않고 고친다(그동안 사용자가 도형을 고르면 멈춘다).
    /// 읽기 전용으로 열린 창이면 고친 내용을 그 이름으로 저장할 수 없으니 결과에 알린다.
    /// CellBox 동기화 폴더의 큰 PPT가 읽기 전용으로 열려 고친 내용을 저장하지 못한 일이 있었다(2026-11-03).
    /// </summary>
    private async Task<McpToolResult> WriteAsync(string path, Func<Task<McpToolResult>> write)
    {
        bool inWindow = _ppt.IsOpenInWindow(path);
        using var watch = inWindow ? _ppt.WatchUserActivity(path) : null;
        var result = await write();
        if (result.IsError || !inWindow || !IsReadOnlyWindow(path)) return result;
        return result with
        {
            Content = [.. result.Content, McpContent.FromText(
                "주의: 이 문서는 PowerPoint에 읽기 전용으로 열려 있습니다. 고친 내용은 창에 보이지만 같은 이름으로 저장되지 않으니, " +
                "사용자에게 '다른 이름으로 저장'하거나 파일을 쓰기 가능하게 다시 열어 달라고 알리세요.")],
        };
    }

    private bool IsReadOnlyWindow(string path) =>
        _ppt.ListOpenWindows().Any(o => o.IsReadOnly && string.Equals(o.Path, path, StringComparison.OrdinalIgnoreCase));

    // ── MCP 전용 도구 ────────────────────────────────────────

    private McpToolResult ListPresentations()
    {
        var open = _ppt.ListOpenWindows();
        var protectedViews = _ppt.ListProtectedViews();
        if (open.Count == 0 && protectedViews.Count == 0) return McpToolResult.Ok("PowerPoint에 열린 문서가 없습니다.");
        var sb = new StringBuilder($"열린 문서 {open.Count}개\n");
        foreach (var pv in protectedViews)
            sb.AppendLine($"  {pv.Name} · 「제한된 보기」 — 도구가 읽지 못함. open_presentation(path, unblock=true)로 편집 모드로 다시 연다" +
                          (pv.Path.Length > 0 ? $" · {pv.Path}" : ""));
        foreach (var o in open)
            sb.AppendLine($"{(o.IsActive ? "▶" : " ")} {o.Name} · {o.SlideCount}쪽 · {(o.IsSaved ? o.Path : "저장 안 된 새 문서")}" +
                          (o.IsReadOnly ? " · 읽기 전용(고쳐도 같은 이름으로 저장 안 됨)" : ""));
        if (_ppt.GetActivePresentation() is { CurrentSlide: int current } active)
            sb.AppendLine($"보고 있는 쪽: {active.Name} {current}쪽");
        return McpToolResult.Ok(sb.ToString().TrimEnd());
    }

    /// <summary>
    /// 연 뒤 문서가 올라올 때까지 기다린다. 26MB PPT를 연 직후 다음 도구가 「열린 문서: 없음」을 받았고(셸로 열면 바로 돌아온다),
    /// 메일로 받은 파일은 「제한된 보기」로 열려 다른 도구가 문서를 찾지 못했다(2026-11-17).
    /// </summary>
    private async Task<McpToolResult> OpenPresentation(JsonElement args, CancellationToken ct)
    {
        var raw = ToolArgs.OptionalString(args, "path")?.Trim().Trim('"') ?? throw new AgentToolException("인자 'path'가 필요합니다.");
        if (!Path.IsPathFullyQualified(raw)) throw new AgentToolException("path는 PPT 절대 경로여야 합니다.");
        var path = ResolveFileOnce(raw);
        bool unblock = ToolArgs.OptionalBool(args, "unblock") == true;
        var timeout = TimeSpan.FromSeconds(Math.Clamp(ToolArgs.OptionalInt(args, "wait_seconds") ?? 60, 1, 300));
        _lastOpen = DateTime.UtcNow;
        var r = await Task.Run(() => _ppt.OpenAndWait(path, unblock, timeout, ct), ct);
        _lastOpen = DateTime.UtcNow;
        string mark = r.Unblocked ? " 인터넷에서 받은 파일 표시를 지웠습니다." : "";
        return r.State switch
        {
            OpenedState.Editable => McpToolResult.Ok($"열었습니다: {path} · {r.SlideCount}쪽 · 편집 가능 ({r.Waited.TotalSeconds:0.#}초).{mark}"),
            OpenedState.ReadOnly => McpToolResult.Ok($"열었습니다: {path} · {r.SlideCount}쪽 · 읽기 전용 ({r.Waited.TotalSeconds:0.#}초).{mark} " +
                                                     "고친 내용을 같은 이름으로 저장할 수 없으니, 고칠 일이 있으면 사용자에게 다른 곳에서 열었는지 확인하세요."),
            OpenedState.ProtectedView => McpToolResult.Ok(
                $"「제한된 보기」로 열렸습니다: {path}" + (r.HadDownloadMark ? " (메일·인터넷에서 받은 파일 표시가 있음)" : "") + ". " +
                "이 상태에서는 다른 도구가 이 문서를 찾지 못합니다. 사용자에게 확인받은 뒤 unblock=true로 다시 부르면 표시를 지우고 편집 모드로 다시 엽니다."),
            _ => McpToolResult.Ok($"여는 중입니다: {path} — {r.Waited.TotalSeconds:0}초 안에 문서가 올라오지 않았습니다. " +
                                  "잠시 뒤 list_presentations로 확인하세요(PowerPoint에 대화상자가 떠 있을 수도 있습니다)."),
        };
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

    private Task<McpToolResult> ImportSlides(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        var source = ResolveFile(ToolArgs.OptionalString(args, "source") ?? throw new AgentToolException("인자 'source'가 필요합니다."));
        if (string.Equals(source, path, StringComparison.OrdinalIgnoreCase))
            throw new AgentToolException("같은 문서의 쪽은 duplicate_slide로 복제하세요.");
        int from = ToolArgs.Int(args, "from");
        int to = ToolArgs.OptionalInt(args, "to") ?? from;
        int after = ToolArgs.Int(args, "after");
        if (from < 1 || to < from) throw new AgentToolException("from은 1 이상, to는 from 이상이어야 합니다.");
        int sourceCount = _ppt.ListSlides(source).Count;
        if (to > sourceCount) throw new AgentToolException($"{Path.GetFileName(source)}는 {sourceCount}쪽까지 있습니다.");
        return WriteAsync(path, async () =>
        {
            var (first, count) = await Task.Run(() => _ppt.ImportSlides(source, from, to, path, after), ct);
            return McpToolResult.Ok($"{Path.GetFileName(source)} {from}~{to}쪽을 {first}~{first + count - 1}쪽으로 붙였습니다. " +
                                    $"뒤 쪽 번호는 {count}씩 밀렸습니다. 디자인은 이 문서 것을 따르니 render_slide로 확인하세요.");
        });
    }

    private Task<McpToolResult> DeleteSlides(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        var slides = args.TryGetProperty("slides", out var s) && s.ValueKind == JsonValueKind.Array
            ? s.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Number).Select(e => e.GetInt32()).ToList()
            : throw new AgentToolException("인자 'slides'(쪽 번호 배열)가 필요합니다.");
        if (!_ppt.IsOpenInWindow(path) && ToolArgs.OptionalBool(args, "confirm") != true)
            throw new AgentToolException("창에 열리지 않은 파일은 지우는 즉시 저장됩니다. 사용자에게 지울 쪽을 확인받은 뒤 confirm=true로 다시 부르세요.");
        return WriteAsync(path, async () =>
        {
            var deleted = await Task.Run(() => _ppt.DeleteSlides(path, slides), ct);
            return McpToolResult.Ok($"{deleted.Count}쪽을 지웠습니다(지우기 전 번호): " +
                                    string.Join(", ", deleted.Select(d => $"{d.Index}쪽 「{(d.Title.Length > 0 ? d.Title : "제목 없음")}」")));
        });
    }

    private Task<McpToolResult> MoveSlide(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide"), to = ToolArgs.Int(args, "to");
        return WriteAsync(path, async () =>
        {
            await Task.Run(() => _ppt.MoveSlide(path, slide, to), ct);
            return McpToolResult.Ok($"{slide}쪽을 {to}쪽 자리로 옮겼습니다.");
        });
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

    // ── 여러 쪽 렌더 ─────────────────────────────────────────

    public const int MaxSeparateRenders = 12, MaxSheetRenders = 60;

    private async Task<McpToolResult> RenderSlides(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        List<int> slides;
        if (args.TryGetProperty("slides", out var s) && s.ValueKind == JsonValueKind.Array)
            slides = s.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Number).Select(e => e.GetInt32()).Distinct().ToList();
        else if (ToolArgs.OptionalInt(args, "from") is int from)
            slides = Enumerable.Range(from, Math.Max(1, (ToolArgs.OptionalInt(args, "to") ?? from) - from + 1)).ToList();
        else throw new AgentToolException("slides(쪽 번호 배열)나 from~to를 주세요.");
        if (slides.Count == 0) throw new AgentToolException("렌더할 쪽이 없습니다.");
        int count = _ppt.ListSlides(path).Count;
        if (slides.Where(n => n < 1 || n > count).ToList() is { Count: > 0 } bad)
            throw new AgentToolException($"{Path.GetFileName(path)}는 1~{count}쪽입니다({string.Join(",", bad)}쪽 없음).");
        bool sheet = ToolArgs.OptionalBool(args, "sheet") ?? slides.Count > MaxSeparateRenders;
        if (slides.Count > (sheet ? MaxSheetRenders : MaxSeparateRenders))
            throw new AgentToolException(sheet ? $"축소판은 {MaxSheetRenders}쪽까지입니다." : $"따로 받는 그림은 {MaxSeparateRenders}쪽까지입니다. sheet=true로 모으세요.");

        var pngs = new List<(int Slide, string? Png)>();
        try
        {
            using (var keep = _ppt.IsOpenInWindow(path) ? null : _ppt.KeepOpen(path, readOnly: true))
                foreach (var n in slides)
                {
                    ct.ThrowIfCancellationRequested();
                    pngs.Add((n, await _ppt.GetSlideThumbnailAsync(path, n)));
                }
            var failed = pngs.Where(p => p.Png is null).Select(p => p.Slide).ToList();
            var ok = pngs.Where(p => p.Png is not null).ToList();
            if (ok.Count == 0) throw new AgentToolException("렌더하지 못했습니다.");
            string note = failed.Count > 0 ? $" (렌더 실패: {string.Join(",", failed)}쪽)" : "";
            if (sheet)
            {
                int columns = Math.Clamp(ToolArgs.OptionalInt(args, "columns") ?? 4, 1, 8);
                var bytes = ContactSheet.Compose(ok.Select(p => (p.Slide, p.Png!)).ToList(), columns);
                return new([McpContent.FromText($"{Path.GetFileName(path)} {string.Join(",", ok.Select(p => p.Slide))}쪽 축소판{note}"), McpContent.FromImage(bytes)], false);
            }
            var content = new List<McpContent> { McpContent.FromText($"{Path.GetFileName(path)} {ok.Count}쪽{note}") };
            foreach (var (n, png) in ok)
            {
                content.Add(McpContent.FromText($"{n}쪽"));
                content.Add(McpContent.FromImage(await File.ReadAllBytesAsync(png!, ct)));
            }
            return new(content, false);
        }
        finally
        {
            foreach (var (_, png) in pngs) if (png is not null) try { File.Delete(png); } catch { }
        }
    }

    // ── 그림 ────────────────────────────────────────────────

    private static readonly string[] ImageExtensions = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg"];

    private static string ImagePath(string? path)
    {
        path = path?.Trim().Trim('"');
        if (string.IsNullOrEmpty(path)) throw new AgentToolException("그림 파일 경로가 필요합니다.");
        if (!Path.IsPathFullyQualified(path)) throw new AgentToolException($"그림 경로는 절대 경로여야 합니다: {path}");
        var full = Path.GetFullPath(path);
        if (!ImageExtensions.Contains(Path.GetExtension(full), StringComparer.OrdinalIgnoreCase))
            throw new AgentToolException($"그림 파일이 아닙니다({string.Join("·", ImageExtensions)}): {full}");
        return File.Exists(full) ? full : throw new AgentToolException($"그림 파일이 없습니다: {full}");
    }

    private static PictureFit ParseFit(string? fit) => fit?.Trim().ToLowerInvariant() switch
    {
        null or "" or "cover" => PictureFit.Cover,
        "contain" => PictureFit.Contain,
        "stretch" => PictureFit.Stretch,
        _ => throw new AgentToolException("fit은 cover·contain·stretch 중 하나입니다."),
    };

    private static CropFractions? ParseCrop(JsonElement crop)
    {
        if (crop.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined) return null;
        if (crop.ValueKind != JsonValueKind.Object) throw new AgentToolException("crop은 {left,top,right,bottom}(0~1) 객체입니다.");
        float F(string n) => ToolArgs.OptionalFloat(crop, n) ?? 0f;
        var c = new CropFractions(F("left"), F("top"), F("right"), F("bottom"));
        if (new[] { c.Left, c.Top, c.Right, c.Bottom }.Any(v => v is < 0 or >= 1))
            throw new AgentToolException("crop 값은 0 이상 1 미만의 비율입니다(원본 기준, 0.1 = 10%).");
        if (c.Left + c.Right >= 0.99f || c.Top + c.Bottom >= 0.99f)
            throw new AgentToolException("crop이 그림을 다 잘라 냅니다(left+right, top+bottom은 1보다 작아야 합니다).");
        return c;
    }

    private static CropFractions? ParseCrop(JsonElement args, string name) =>
        args.TryGetProperty(name, out var c) ? ParseCrop(c) : null;

    private static string Describe(PlacedPicture p) =>
        $"  {p.Address} | {p.Left:0.##},{p.Top:0.##} {p.Width:0.##}×{p.Height:0.##}" +
        (p.Crop.IsNone ? " | 자르지 않음" : $" | 잘림 {p.Crop}") +
        (p.PixelWidth is int w && p.PixelHeight is int h ? $" | 원본 {w}×{h}px" : "");

    private Task<McpToolResult> AddPicture(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        var request = new PictureRequest(ImagePath(ToolArgs.OptionalString(args, "path")),
            ToolArgs.OptionalFloat(args, "left") ?? throw new AgentToolException("인자 'left'가 필요합니다."),
            ToolArgs.OptionalFloat(args, "top") ?? throw new AgentToolException("인자 'top'이 필요합니다."),
            ToolArgs.OptionalFloat(args, "width") ?? throw new AgentToolException("인자 'width'가 필요합니다."),
            ToolArgs.OptionalFloat(args, "height") ?? throw new AgentToolException("인자 'height'가 필요합니다."),
            ParseFit(ToolArgs.OptionalString(args, "fit")), ParseCrop(args, "crop"), ToolArgs.OptionalString(args, "name"));
        if (request.Width <= 0 || request.Height <= 0) throw new AgentToolException("width·height는 0보다 커야 합니다.");
        return WriteAsync(path, async () =>
        {
            var placed = await Task.Run(() => _ppt.PlacePictures(path, slide, [request]), ct);
            return McpToolResult.Ok($"{slide}쪽에 그림을 넣었습니다({FitName(request.Fit)}).\n" + string.Join("\n", placed.Select(Describe)));
        });
    }

    private async Task<McpToolResult> AddPicturesInBox(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        var files = ToolArgs.StringArray(args, "paths").Select(ImagePath).ToList();
        if (files.Count == 0) throw new AgentToolException("paths에 그림이 한 장 이상 있어야 합니다.");
        if (!args.TryGetProperty("box", out var boxArg)) throw new AgentToolException("인자 'box'가 필요합니다.");
        var (box, address) = await ResolveBoxAsync(path, slide, boxArg, ct);
        string? remove = null;
        if (args.TryGetProperty("replace_placeholder", out var rp))
        {
            if (rp.ValueKind == JsonValueKind.String && rp.GetString()!.Trim() is { Length: > 0 } other && !bool.TryParse(other, out _))
                remove = (await FindShapeAsync(path, slide, other, ct)).Address;
            else if (ToolArgs.OptionalBool(args, "replace_placeholder") == true)
                remove = address ?? throw new AgentToolException("replace_placeholder=true는 box가 도형 주소일 때만 씁니다. 지울 도형 주소를 replace_placeholder에 직접 주세요.");
        }

        float gap = ToolArgs.OptionalFloat(args, "gap") ?? 3f;
        bool? horizontal = ToolArgs.OptionalString(args, "direction")?.Trim().ToLowerInvariant() switch
        {
            null or "" or "auto" => null, "horizontal" => true, "vertical" => false,
            _ => throw new AgentToolException("direction은 auto·horizontal·vertical 중 하나입니다."),
        };
        IReadOnlyList<BoxF> cells;
        try { cells = PictureFitter.Split(box, files.Count, gap, horizontal); }
        catch (ArgumentException ex) { throw new AgentToolException(ex.Message); }
        var fit = ParseFit(ToolArgs.OptionalString(args, "fit"));
        var crops = args.TryGetProperty("crops", out var c) && c.ValueKind == JsonValueKind.Array ? c.EnumerateArray().Select(ParseCrop).ToList() : [];
        var names = args.TryGetProperty("names", out var n) && n.ValueKind == JsonValueKind.Array
            ? n.EnumerateArray().Select(e => e.ValueKind == JsonValueKind.String ? e.GetString() : null).ToList() : [];
        var requests = files.Select((f, i) => new PictureRequest(f, cells[i].Left, cells[i].Top, cells[i].Width, cells[i].Height, fit,
            i < crops.Count ? crops[i] : null, i < names.Count && !string.IsNullOrWhiteSpace(names[i]) ? names[i] : null)).ToList();

        return await WriteAsync(path, async () =>
        {
            var placed = await Task.Run(() => _ppt.PlacePictures(path, slide, requests, remove), ct);
            return McpToolResult.Ok(
                $"{slide}쪽 칸 {box.Left:0.##},{box.Top:0.##} {box.Width:0.##}×{box.Height:0.##}에 {placed.Count}장을 넣었습니다({FitName(fit)}, 간격 {gap:0.##}pt)" +
                (remove is not null ? $". 「{remove}」는 지웠습니다" : "") + ".\n" + string.Join("\n", placed.Select(Describe)));
        });
    }

    private async Task<McpToolResult> ReplacePicture(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape");
        string file = ImagePath(ToolArgs.OptionalString(args, "path"));
        var entry = await FindShapeAsync(path, slide, shape, ct);
        var request = new PictureRequest(file, entry.Left, entry.Top, entry.Width, entry.Height,
            ParseFit(ToolArgs.OptionalString(args, "fit")), ParseCrop(args, "crop"), ToolArgs.OptionalString(args, "name"));
        return await WriteAsync(path, async () =>
        {
            var placed = await Task.Run(() => _ppt.PlacePictures(path, slide, [request], shape, keepReplacedName: true), ct);
            return McpToolResult.Ok($"{slide}쪽 「{shape}」를 바꿨습니다(같은 자리·크기·Z순서, {FitName(request.Fit)})" +
                                    (entry.Picture is null ? ". 원래 도형은 그림이 아니었습니다" : "") + ".\n" +
                                    string.Join("\n", placed.Select(Describe)));
        });
    }

    private static string FitName(PictureFit fit) => fit switch
    {
        PictureFit.Contain => "contain: 비율 유지",
        PictureFit.Stretch => "stretch: 칸에 늘림",
        _ => "cover: 칸 채우고 넘는 쪽 자름",
    };

    private async Task<ShapeEntry> FindShapeAsync(string path, int slide, string address, CancellationToken ct)
    {
        var shapes = await Task.Run(() => _ppt.ListShapes(path, slide), ct);
        return shapes.FirstOrDefault(s => s.Address == address) ?? shapes.FirstOrDefault(s => s.Name == address)
            ?? throw new AgentToolException($"{slide}쪽에서 도형 '{address}'를 찾지 못했습니다. list_shapes로 주소를 확인하세요.");
    }

    /// <summary>box 인자: {left,top,width,height} 또는 도형 주소. 주소면 그 도형 자리와 주소를 돌려준다.</summary>
    private async Task<(BoxF Box, string? Address)> ResolveBoxAsync(string path, int slide, JsonElement box, CancellationToken ct)
    {
        if (box.ValueKind == JsonValueKind.String)
        {
            string address = box.GetString()!.Trim();
            var e = await FindShapeAsync(path, slide, address, ct);
            return (new BoxF(e.Left, e.Top, e.Width, e.Height), e.Address);
        }
        if (box.ValueKind != JsonValueKind.Object) throw new AgentToolException("box는 {left,top,width,height} 또는 도형 주소입니다.");
        float Need(string n) => ToolArgs.OptionalFloat(box, n) ?? throw new AgentToolException($"box.{n}가 필요합니다.");
        var b = new BoxF(Need("left"), Need("top"), Need("width"), Need("height"));
        if (b.Width <= 0 || b.Height <= 0) throw new AgentToolException("box의 width·height는 0보다 커야 합니다.");
        return (b, null);
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
