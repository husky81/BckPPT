using System.IO;
using System.Text;
using System.Text.Json;
using BckPPT.Models;
using BckPPT.PowerPoint;
using BckPPT.Services;

namespace BckPPT.Agent.Tools;

/// <summary>
/// <see cref="IPptComService"/>를 얇게 감싼 도구 모음. 슬라이드 번호는 1부터, 도형은 ShapeAddress 규약
/// ("이름", "이름#2", "그룹/이름", "표 [R1C2]")으로 가리킨다. 모든 쓰기는 작업 사본에 한다.
/// </summary>
public static class PptAgentTools
{
    private const string SlideShape = """
        "slide": {"type": "integer", "description": "슬라이드 번호(1부터)"},
        "shape": {"type": "string", "description": "도형 주소. list_shapes가 준 값을 그대로 쓴다"}
        """;

    public static IReadOnlyList<AgentTool> Create(bool includeSources)
    {
        var tools = new List<AgentTool>
        {
            new("list_slides", "작업 문서의 슬라이드 번호와 제목 목록", Schema(""), false, ListSlides),
            new("list_shapes", "슬라이드의 도형 목록(그룹 안까지). 주소·위치·크기(pt)·텍스트 앞부분",
                Schema("""{"slide": {"type": "integer"}}""", "slide"), false, ListShapes),
            new("read_text", "도형의 문단 목록과 상자 높이·글자 크기. 문단 번호는 1부터",
                Schema("{" + SlideShape + "}", "slide", "shape"), false, ReadText),
            new("line_breaks",
                "PowerPoint가 실제로 나눈 줄. 각 줄 끝 유형: auto(자동 줄바꿈), enter(문단 끝), shift(Shift+Enter). 문구 길이를 맞출 때 쓰는 핵심 측정 도구",
                Schema("{" + SlideShape + "}", "slide", "shape"), false, LineBreaks),
            new("duplicate_slide", "슬라이드를 복제해 after 번 슬라이드 뒤에 넣는다(after=0이면 맨 앞). 새 슬라이드 번호를 돌려주고 뒤 번호는 1씩 밀린다",
                Schema("""{"source": {"type": "integer"}, "after": {"type": "integer"}}""", "source", "after"), true, DuplicateSlide),
            new("set_paragraphs",
                "도형의 문단 전체를 paragraphs로 바꾼다. 겹치는 문단은 서식을 유지하고 남으면 지우고 모자라면 덧붙인다. 문단 안 줄바꿈(Shift+Enter)은 \\n. 적용 후 줄 측정 결과를 돌려준다",
                Schema("{" + SlideShape + """, "paragraphs": {"type": "array", "items": {"type": "string"}}}""",
                    "slide", "shape", "paragraphs"), true, SetParagraphs),
            new("replace_text", "도형 안의 문자열을 바꾼다. 글자 크기가 섞인 제목처럼 서식을 지켜야 할 때 쓴다. 적용 후 줄 측정 결과를 돌려준다",
                Schema("{" + SlideShape + """, "find": {"type": "string"}, "replace": {"type": "string"}}""",
                    "slide", "shape", "find", "replace"), true, ReplaceText),
            new("rename_shape", "도형 이름을 바꾼다. 같은 이름이 여러 개면 C{열}_{역할}처럼 구분되는 이름을 붙인다",
                Schema("{" + SlideShape + """, "new_name": {"type": "string"}}""", "slide", "shape", "new_name"), true, RenameShape),
            new("set_geometry", "도형 위치·크기(pt)를 바꾼다. 준 값만 바뀐다",
                Schema("{" + SlideShape + """, "left": {"type": "number"}, "top": {"type": "number"}, "width": {"type": "number"}, "height": {"type": "number"}}""",
                    "slide", "shape"), true, SetGeometry),
            new("delete_shape", "도형을 지운다",
                Schema("{" + SlideShape + "}", "slide", "shape"), true, DeleteShape),
            new("add_text_shape", "텍스트 상자(범례·배지)를 추가한다. auto_width면 글자 폭에 맞추고 right_edge가 있으면 오른쪽 끝을 그 위치에 맞춘다",
                Schema("""
                    {"slide": {"type": "integer"}, "text": {"type": "string"},
                     "left": {"type": "number"}, "top": {"type": "number"}, "width": {"type": "number"}, "height": {"type": "number"},
                     "name": {"type": "string"}, "rounded": {"type": "boolean"}, "fill_color": {"type": "string", "description": "#RRGGBB, 없으면 투명"},
                     "font_size": {"type": "number"}, "bold": {"type": "boolean"}, "font_color": {"type": "string"},
                     "auto_width": {"type": "boolean"}, "right_edge": {"type": "number"}}
                    """, "slide", "text", "left", "top", "width", "height"), true, AddTextShape),
            new("set_text_color", "도형 텍스트에서 find가 나오는 곳의 글자색을 바꾼다(#RRGGBB)",
                Schema("{" + SlideShape + """, "find": {"type": "string"}, "color": {"type": "string"}}""",
                    "slide", "shape", "find", "color"), true, SetTextColor),
            new("set_text_format", "도형 전체 텍스트 서식. 준 값만 바뀐다. spacing은 자간(pt)",
                Schema("{" + SlideShape + """, "size": {"type": "number"}, "bold": {"type": "boolean"}, "color": {"type": "string"}, "spacing": {"type": "number"}, "font_name": {"type": "string"}}""",
                    "slide", "shape"), true, SetTextFormat),
            new("lint_slide",
                "레이아웃 규칙 검사: 기준 슬라이드 대비 줄 수, 고아 단어, 넘침, 제목 균형, 같은 역할 상자 서식 차이, 추가 도형 겹침. baseline_slide를 주지 않으면 원본 파일의 같은 슬라이드(복제본이면 복제 원본)를 기준으로 쓰고, 기준과 같은(손대지 않은) 상자는 판정하지 않는다",
                Schema("""{"slide": {"type": "integer"}, "baseline_slide": {"type": "integer"}}""", "slide"), false, LintSlide),
            new("render_slide", "슬라이드를 이미지로 렌더해 눈으로 확인한 설명을 돌려준다. question에 확인할 점을 적는다",
                Schema("""{"slide": {"type": "integer"}, "question": {"type": "string"}}""", "slide"), false, RenderSlide),
            new("finish", "작업을 마친다. summary에 무엇을 바꿨는지 사용자에게 보고할 내용을 적는다",
                Schema("""{"summary": {"type": "string"}}""", "summary"), false, Finish),
        };
        if (includeSources)
        {
            tools.Add(new("source_outline", "첨부한 원문의 목차(절 제목과 시작 쪽). document를 비우면 문서 목록. 원문을 읽기 전에 먼저 본다",
                Schema("""{"document": {"type": "string"}}"""), false, SourceOutline));
            tools.Add(new("search_source", "첨부한 원문 문서에서 관련 쪽을 찾는다. 결과는 문서명·쪽·〈절 제목〉·발췌",
                Schema("""{"query": {"type": "string"}, "limit": {"type": "integer"}}""", "query"), false, SearchSource));
            tools.Add(new("read_source", "원문 문서의 쪽 범위를 읽는다",
                Schema("""{"document": {"type": "string"}, "from_page": {"type": "integer"}, "to_page": {"type": "integer"}}""",
                    "document", "from_page"), false, ReadSource));
        }
        return tools;
    }

    private static string Schema(string properties, params string[] required)
    {
        var props = properties.Length == 0 ? "{}" : properties;
        var req = string.Join(", ", required.Select(r => $"\"{r}\""));
        return $$"""{"type": "object", "properties": {{props}}, "required": [{{req}}]}""";
    }

    // ── 읽기 ───────────────────────────────────────────────────

    private static async Task<string> ListSlides(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        var slides = await Task.Run(() => ctx.Ppt.ListSlides(ctx.WorkingPath), ct);
        var sb = new StringBuilder($"슬라이드 {slides.Count}장\n");
        foreach (var (index, title) in slides)
        {
            sb.Append($"{index}. {(string.IsNullOrWhiteSpace(title) ? "(제목 없음)" : title)}");
            if (ctx.Origins.BaselineOf(index) is int origin) sb.Append($" [← {origin}쪽 복제]");
            sb.Append('\n');
        }
        return sb.ToString();
    }

    private static async Task<string> ListShapes(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        var shapes = await Task.Run(() => ctx.Ppt.ListShapes(ctx.WorkingPath, slide), ct);
        var sb = new StringBuilder(
            $"{slide}쪽 도형 {shapes.Count}개 (주소 | left,top width×height | 텍스트 앞 60자, ¶ 문단 경계, ↵ Shift+Enter)\n" +
            "미리보기는 원문이 아니다. 문구를 바꾸기 전에 read_text로 원문을 확인한다.\n");
        foreach (var s in shapes)
        {
            sb.Append($"{s.Address} | {s.Left:0.#},{s.Top:0.#} {s.Width:0.#}×{s.Height:0.#}");
            if (s.HasText) sb.Append($" | {s.TextPreview}");
            sb.Append('\n');
        }
        return sb.ToString();
    }

    private static async Task<string> ReadText(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape");
        var info = await Task.Run(() => ctx.Ppt.GetShapeInfo(ctx.WorkingPath, slide, shape), ct)
            ?? throw new AgentToolException($"{slide}쪽에서 텍스트 도형 '{shape}'를 찾지 못했습니다. list_shapes로 주소를 확인하세요.");
        var sb = new StringBuilder(
            $"{shape}: 글자 {info.FontSize:0.#}pt, 상자 높이 {info.BoxHeight:0.#}pt, 텍스트 높이 {info.BoundHeight:0.#}pt, {info.LineCount}줄{(info.IsOverflow ? " (넘침)" : "")}\n");
        var paragraphs = info.FullText.TrimEnd('\r').Split('\r');
        for (int i = 0; i < paragraphs.Length; i++)
            sb.Append($"[{i + 1}] {paragraphs[i].Replace("\v", "\\n")}\n");
        return sb.ToString();
    }

    private static async Task<string> LineBreaks(AgentToolContext ctx, JsonElement args, CancellationToken ct) =>
        await MeasureAsync(ctx, ToolArgs.Int(args, "slide"), ToolArgs.String(args, "shape"), ct);

    /// <summary>줄 측정 결과. 문단 번호와 줄 끝 유형을 함께 보여 모델이 어느 문단을 줄일지 바로 알게 한다.</summary>
    internal static async Task<string> MeasureAsync(AgentToolContext ctx, int slide, string shape, CancellationToken ct)
    {
        var lines = await Task.Run(() => ctx.Ppt.GetTextLineBreaks(ctx.WorkingPath, slide, shape), ct);
        if (lines.Count == 0) return $"{shape}: 텍스트 없음(또는 도형을 찾지 못함)";
        var budget = await BaselineLineCountAsync(ctx, slide, shape, ct);
        var verdict = budget is not int b ? ""
            : lines.Count <= b ? $" (기준 {b}줄 이내)"
            : $" (기준 {b}줄{LayoutLint.ExcessHint(lines, b)})";
        var sb = new StringBuilder($"{shape}: {lines.Count}줄{verdict}\n");
        int paragraph = 1;
        foreach (var line in lines)
        {
            string end = line.BreakType switch
            {
                LineBreakType.AutoWrap => "auto",
                LineBreakType.Enter => "enter",
                LineBreakType.ShiftEnter => "shift",
                _ => "end"
            };
            sb.Append($"{line.LineNumber}. (문단{paragraph}) {line.Text} [{end}]\n");
            if (line.BreakType == LineBreakType.Enter) paragraph++;
        }
        return sb.ToString();
    }

    /// <summary>원본 파일에서 같은 상자(복제본이면 복제 원본의 상자)의 줄 수. 원본에 없는 상자면 null.</summary>
    private static async Task<int?> BaselineLineCountAsync(AgentToolContext ctx, int slide, string shape, CancellationToken ct)
    {
        if (ctx.BaselinePath is not { } original) return null;
        try
        {
            var lines = await Task.Run(() => ctx.Ppt.GetTextLineBreaks(original, ctx.Origins.OriginalIndexOf(slide), shape), ct);
            return lines.Count > 0 ? lines.Count : null;
        }
        catch (OperationCanceledException) { throw; }
        catch { return null; }
    }

    // ── 쓰기 ───────────────────────────────────────────────────

    private static async Task<string> DuplicateSlide(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int source = ToolArgs.Int(args, "source"), after = ToolArgs.Int(args, "after");
        int created = await Task.Run(() => ctx.Ppt.DuplicateSlide(ctx.WorkingPath, source, after), ct);
        ctx.Origins.RecordDuplicate(source, created);
        ctx.Changes.Add($"{source}쪽을 복제해 {created}쪽으로 추가");
        return $"{source}쪽을 복제해 {created}쪽에 넣었습니다. {created}쪽 이후 기존 슬라이드 번호는 1씩 밀렸습니다.";
    }

    private static async Task<string> SetParagraphs(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape");
        var paragraphs = ParagraphFitPlanner.SplitPilcrows(ToolArgs.StringArray(args, "paragraphs"));
        if (paragraphs.Count == 0) throw new AgentToolException("paragraphs가 비었습니다. 지우려면 delete_shape를 쓰세요.");

        var info = await Task.Run(() => ctx.Ppt.GetShapeInfo(ctx.WorkingPath, slide, shape), ct)
            ?? throw new AgentToolException($"{slide}쪽에서 텍스트 도형 '{shape}'를 찾지 못했습니다.");
        var edits = ParagraphFitPlanner.Plan(slide, shape, ParagraphFitPlanner.CountParagraphs(info.FullText), paragraphs);
        var result = await Task.Run(() => ctx.Ppt.SetParagraphTexts(ctx.WorkingPath, edits, markRed: false), ct);
        ctx.Origins.MarkTouched(slide);
        ctx.Changes.Add($"{slide}쪽 {shape} 문단 {paragraphs.Count}개로 교체");
        var failures = result.Failures.Count == 0 ? "" : "실패:\n" + string.Join("\n", result.Failures) + "\n";
        return failures + await MeasureAsync(ctx, slide, shape, ct);
    }

    private static async Task<string> ReplaceText(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape"), find = ToolArgs.String(args, "find");
        string replace = ToolArgs.OptionalString(args, "replace") ?? "";
        if (find.Contains('\n') || find.Contains('\r') || find.Contains('¶'))
            throw new AgentToolException("replace_text는 문단 경계를 넘어 바꿀 수 없습니다. 여러 문단을 고치려면 set_paragraphs로 문단 전체를 다시 넣으세요.");
        int count = await Task.Run(() => ctx.Ppt.ReplaceTextInShape(
            ctx.WorkingPath, slide, shape, find, ParagraphFitPlanner.NormalizeParagraph(replace)), ct);
        if (count == 0) throw new AgentToolException($"'{find}'를 {shape}에서 찾지 못했습니다. read_text로 현재 문구를 확인하세요.");
        ctx.Origins.MarkTouched(slide);
        ctx.Changes.Add($"{slide}쪽 {shape} '{find}' → '{replace}' {count}곳");
        return $"{count}곳 바꿈\n" + await MeasureAsync(ctx, slide, shape, ct);
    }

    private static async Task<string> RenameShape(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape"), name = ToolArgs.String(args, "new_name");
        if (shape == name) return $"이미 '{name}'입니다.";
        await Task.Run(() => ctx.Ppt.RenameShape(ctx.WorkingPath, slide, shape, name), ct);
        ctx.Origins.MarkTouched(slide);
        ctx.Changes.Add($"{slide}쪽 {shape} 이름 → {name}");
        return $"{shape} → {name}";
    }

    private static async Task<string> SetGeometry(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape");
        var geometry = new ShapeGeometry(ToolArgs.OptionalFloat(args, "left"), ToolArgs.OptionalFloat(args, "top"),
            ToolArgs.OptionalFloat(args, "width"), ToolArgs.OptionalFloat(args, "height"));
        if (geometry == new ShapeGeometry()) throw new AgentToolException("left/top/width/height 중 하나는 있어야 합니다.");
        await Task.Run(() => ctx.Ppt.SetShapeGeometry(ctx.WorkingPath, slide, shape, geometry), ct);
        ctx.Origins.MarkTouched(slide);
        ctx.Changes.Add($"{slide}쪽 {shape} 위치·크기 변경");
        return $"{shape} 배치를 바꿨습니다.";
    }

    private static async Task<string> DeleteShape(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape");
        await Task.Run(() => ctx.Ppt.DeleteShape(ctx.WorkingPath, slide, shape), ct);
        ctx.Origins.MarkTouched(slide);
        ctx.AddedShapes.Remove((slide, shape));
        ctx.Changes.Add($"{slide}쪽 {shape} 삭제");
        return $"{shape}를 지웠습니다. 같은 이름 도형의 #n 번호가 바뀌었을 수 있으니 list_shapes로 다시 확인하세요.";
    }

    private static async Task<string> AddTextShape(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        var request = new TextShapeRequest(
            ToolArgs.String(args, "text"),
            ToolArgs.OptionalFloat(args, "left") ?? 0f, ToolArgs.OptionalFloat(args, "top") ?? 0f,
            ToolArgs.OptionalFloat(args, "width") ?? 100f, ToolArgs.OptionalFloat(args, "height") ?? 20f,
            ToolArgs.OptionalString(args, "name"),
            ToolArgs.OptionalBool(args, "rounded") == true ? TextShapeKind.RoundedRectangle : TextShapeKind.Rectangle,
            ToolArgs.OptionalString(args, "fill_color"),
            FontSize: ToolArgs.OptionalFloat(args, "font_size") ?? 10f,
            Bold: ToolArgs.OptionalBool(args, "bold") ?? false,
            FontColor: ToolArgs.OptionalString(args, "font_color") ?? "#000000",
            AutoWidth: ToolArgs.OptionalBool(args, "auto_width") ?? false,
            RightEdge: ToolArgs.OptionalFloat(args, "right_edge"));
        string name = await Task.Run(() => ctx.Ppt.AddTextShape(ctx.WorkingPath, slide, request), ct);
        ctx.Origins.MarkTouched(slide);
        ctx.AddedShapes.Add((slide, name));
        ctx.Changes.Add($"{slide}쪽 텍스트 상자 '{name}' 추가");
        return $"'{name}'을 추가했습니다.";
    }

    private static async Task<string> SetTextColor(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape"), find = ToolArgs.String(args, "find"), color = ToolArgs.String(args, "color");
        int count = await Task.Run(() => ctx.Ppt.SetTextColor(ctx.WorkingPath, slide, shape, find, color), ct);
        ctx.Origins.MarkTouched(slide);
        ctx.Changes.Add($"{slide}쪽 {shape} '{find}' 글자색 {color}");
        return $"{count}곳 색을 바꿨습니다.";
    }

    private static async Task<string> SetTextFormat(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape");
        var patch = new TextFormatPatch(ToolArgs.OptionalFloat(args, "size"), ToolArgs.OptionalBool(args, "bold"),
            ToolArgs.OptionalString(args, "color"), ToolArgs.OptionalFloat(args, "spacing"), ToolArgs.OptionalString(args, "font_name"));
        await Task.Run(() => ctx.Ppt.SetTextFormat(ctx.WorkingPath, slide, shape, patch), ct);
        ctx.Origins.MarkTouched(slide);
        ctx.Changes.Add($"{slide}쪽 {shape} 서식 변경");
        return $"{shape} 서식을 바꿨습니다.\n" + await MeasureAsync(ctx, slide, shape, ct);
    }

    // ── 검증 ───────────────────────────────────────────────────

    private static async Task<string> LintSlide(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        var issues = await LintAsync(ctx, slide, ToolArgs.OptionalInt(args, "baseline_slide"), ct);
        return $"{slide}쪽 " + LayoutLint.Describe(issues);
    }

    /// <summary>슬라이드를 측정해 판정한다. 기준을 주지 않으면 세션에서 기록한 복제 원본을 쓴다.</summary>
    public static async Task<IReadOnlyList<LintIssue>> LintAsync(
        AgentToolContext ctx, int slide, int? baselineSlide, CancellationToken ct)
    {
        var added = ctx.AddedShapes.Where(a => a.Slide == slide).Select(a => a.Address).ToHashSet();
        var target = await Task.Run(() => SlideLintCollector.Collect(ctx.Ppt, ctx.WorkingPath, slide, added), ct);
        LintSlide? baseline = null;
        if (baselineSlide is int explicitSlide && explicitSlide != slide)
            baseline = await Task.Run(() => SlideLintCollector.Collect(ctx.Ppt, ctx.WorkingPath, explicitSlide), ct);
        else if (ctx.BaselinePath is { } original)
        {
            int originalSlide = ctx.Origins.OriginalIndexOf(slide);
            baseline = await Task.Run(() => SlideLintCollector.Collect(ctx.Ppt, original, originalSlide), ct);
        }
        else if (ctx.Origins.BaselineOf(slide) is int origin)
            baseline = await Task.Run(() => SlideLintCollector.Collect(ctx.Ppt, ctx.WorkingPath, origin), ct);
        return LayoutLint.Check(target, baseline, LintSettings.FromRules(PptRules.Current));
    }

    private static async Task<string> RenderSlide(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        string question = ToolArgs.OptionalString(args, "question")
            ?? "이 슬라이드에서 글자가 상자를 넘치거나, 도형이 겹치거나, 정렬이 어긋난 곳을 구체적으로 지적해줘.";
        var png = await ctx.Ppt.GetSlideThumbnailAsync(ctx.WorkingPath, slide)
            ?? throw new AgentToolException($"{slide}쪽을 렌더하지 못했습니다.");
        try
        {
            if (ctx.DescribeImage is null) return $"{slide}쪽 렌더 완료(이미지 확인 모델 없음). lint_slide 결과로 판단하세요.";
            return $"{slide}쪽 렌더 확인: " + await ctx.DescribeImage(png, question, ct);
        }
        finally { try { File.Delete(png); } catch { } }
    }

    private static Task<string> Finish(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        ctx.FinishSummary = ToolArgs.String(args, "summary");
        return Task.FromResult("종료합니다.");
    }

    // ── 원문 ───────────────────────────────────────────────────

    private static Task<string> SearchSource(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        var library = ctx.Sources ?? throw new AgentToolException("첨부된 원문이 없습니다.");
        // 1건만 보고 엉뚱한 쪽을 읽는 일이 있어(실측) 최소 3건은 보여준다.
        var hits = library.Search(ToolArgs.String(args, "query"), Math.Clamp(ToolArgs.OptionalInt(args, "limit") ?? 5, 3, 10));
        if (hits.Count == 0) return Task.FromResult("찾은 쪽이 없습니다. 다른 낱말로 검색하세요.");
        return Task.FromResult(string.Join("\n\n", hits.Select(h => $"[{h.Document} p.{h.Page}] {h.Snippet}")));
    }

    private static Task<string> SourceOutline(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        var library = ctx.Sources ?? throw new AgentToolException("첨부된 원문이 없습니다.");
        return Task.FromResult(library.Outline(ToolArgs.OptionalString(args, "document")));
    }

    private static Task<string> ReadSource(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        var library = ctx.Sources ?? throw new AgentToolException("첨부된 원문이 없습니다.");
        int from = ToolArgs.Int(args, "from_page");
        int to = ToolArgs.OptionalInt(args, "to_page") ?? from;
        return Task.FromResult(library.Read(ToolArgs.String(args, "document"), from, Math.Max(from, to)));
    }
}
