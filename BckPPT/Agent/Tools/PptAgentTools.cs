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
            new("rename_shape", "도형 이름을 바꾼다. 여러 개면 rename_shapes로 한 번에 바꾼다",
                Schema("{" + SlideShape + """, "new_name": {"type": "string"}}""", "slide", "shape", "new_name"), true, RenameShape),
            new("rename_shapes", "한 슬라이드의 도형 이름을 한 번에 바꾼다. 카드마다 C{카드}_{역할} 이름을 붙일 때 쓴다(한 번 호출로 슬라이드 전체)",
                Schema("""
                    {"slide": {"type": "integer"},
                     "items": {"type": "array", "items": {"type": "object",
                       "properties": {"shape": {"type": "string"}, "new_name": {"type": "string"}}, "required": ["shape", "new_name"]}}}
                    """, "slide", "items"), true, RenameShapes),
            new("set_texts", "여러 상자의 문단을 한 번에 바꾸고 상자마다 줄 측정 결과를 돌려준다. 카드 하나의 제목·개요·문제점·연구개발 내용을 한 번에 채울 때 쓴다",
                Schema("""
                    {"slide": {"type": "integer"},
                     "items": {"type": "array", "items": {"type": "object",
                       "properties": {"shape": {"type": "string"}, "paragraphs": {"type": "array", "items": {"type": "string"}}},
                       "required": ["shape", "paragraphs"]}}}
                    """, "slide", "items"), true, SetTexts),
            new("remove_blank_lines",
                "슬라이드들의 글 상자에서 빈 줄(빈 문단)과 상자 끝에 남은 Enter를 지운다. '공백줄·빈 줄·빈 칸 삭제' 요청에 쓴다. 상자를 몰라도 슬라이드 번호만 주면 된다",
                Schema("""{"from_slide": {"type": "integer"}, "to_slide": {"type": "integer", "description": "없으면 from_slide 한 쪽"}}""", "from_slide"),
                true, RemoveBlankLines),
            new("set_geometry", "도형 위치·크기(pt)를 바꾼다. 준 값만 바뀐다",
                Schema("{" + SlideShape + """, "left": {"type": "number"}, "top": {"type": "number"}, "width": {"type": "number"}, "height": {"type": "number"}}""",
                    "slide", "shape"), true, SetGeometry),
            new("delete_shape", "도형을 지운다. with_contents=true면 그 도형 영역 안에 든 도형(카드 안의 글 상자 등)도 함께 지운다. 카드 하나를 통째로 지울 때 쓴다",
                Schema("{" + SlideShape + """, "with_contents": {"type": "boolean"}}""", "slide", "shape"), true, DeleteShape),
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
            // 원문이 없어도 보이게 둔다. 없으면 도구가 첨부를 안내한다(원문 첨부 전 요청에 "도구가 없다"고 끝난 일이 있었다).
            new("fill_cards",
                "카드형 템플릿 슬라이드를 복제해 보고서들을 카드로 정리한다. 요청 해석·복제·남는 카드 삭제·패널 폭·이름·원문 요점·문안 작성·길이 맞춤을 한 번에 한다. " +
                "이런 요청은 다른 도구를 부르기 전에 이 도구를 먼저 한 번 부른다(슬라이드를 미리 복제하지 않는다). " +
                "request에는 사용자 요청을 줄이지 말고 그대로 넣는다",
                Schema("""{"request": {"type": "string", "description": "사용자 요청 전문"}}"""), true, FillCards),
        };
        if (includeSources)
        {
            tools.Add(new("brief_source",
                "원문 한 문서에서 요약문·개요, 문제점·필요성, 연구개발 내용·목표 절을 찾아 읽고 요점만 정리해 돌려준다(쪽 번호 포함). " +
                "카드를 채울 때 문서마다 이것을 먼저 부른다. document는 D1 같은 별칭. focus로 원하는 항목을 바꿀 수 있다",
                Schema("""{"document": {"type": "string"}, "focus": {"type": "string"}}""", "document"), false, BriefSource));
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
        return FormatShapeList(slide, shapes);
    }

    /// <summary>
    /// 도형을 카드(텍스트 없는 큰 사각형)별로 묶어 위에서 아래 순서로 보여준다. 같은 이름이 63개씩 겹치는
    /// 템플릿에서 평면 목록만 주면 모델이 소제목("개요")을 본문으로 알고 바꿨다(2026-10-06 실측).
    /// 짧은 한 줄 텍스트는 (소제목)으로 표시한다.
    /// </summary>
    public static string FormatShapeList(int slide, IReadOnlyList<ShapeEntry> shapes)
    {
        float widest = shapes.Count == 0 ? 0 : shapes.Max(s => s.Width);
        // 카드: 텍스트가 없고 세로로 긴 사각형. 슬라이드 전체를 덮는 패널은 카드가 아니라 바깥 틀이다.
        var cards = shapes
            .Where(s => !s.HasText && s.Height >= 150 && s.Width >= 80 && s.Width < widest * 0.6)
            .OrderBy(s => s.Left).ThenBy(s => s.Top)
            .ToList();
        ShapeEntry? CardOf(ShapeEntry s)
        {
            float cx = s.Left + s.Width / 2, cy = s.Top + s.Height / 2;
            return cards.Where(c => !ReferenceEquals(c, s) && cx >= c.Left && cx <= c.Left + c.Width &&
                                    cy >= c.Top && cy <= c.Top + c.Height)
                        .OrderBy(c => c.Width * c.Height).FirstOrDefault();
        }

        var sb = new StringBuilder(
            $"{slide}쪽 도형 {shapes.Count}개 (주소 (이름) | left,top width×height | 텍스트 앞 60자, ¶ 문단 경계, ↵ Shift+Enter)\n" +
            "이름이 겹치는 도형은 @숫자 주소로 가리킨다. 이 주소는 다른 도형 이름을 바꿔도 변하지 않는다.\n" +
            "미리보기는 원문이 아니다. 문구를 바꾸기 전에 read_text로 원문을 확인한다. (소제목) 상자는 보통 그대로 둔다.\n");
        void Line(ShapeEntry s)
        {
            sb.Append($"  {s.Address}");
            if (s.Address != s.Name) sb.Append($" ({s.Name})");
            sb.Append($" | {s.Left:0.#},{s.Top:0.#} {s.Width:0.#}×{s.Height:0.#}");
            if (s.HasText)
            {
                bool label = s.Height <= 22 && s.TextPreview.Length <= 10 && !s.TextPreview.Contains('¶');
                sb.Append(label ? " | (소제목) " : " | ").Append(s.TextPreview);
            }
            sb.Append('\n');
        }

        var grouped = shapes.Where(s => !cards.Contains(s)).GroupBy(CardOf).ToDictionary(g => g.Key ?? Outside, g => g.ToList());
        for (int i = 0; i < cards.Count; i++)
        {
            var card = cards[i];
            sb.Append($"[카드 {i + 1}] {card.Address}{(card.Address != card.Name ? $" ({card.Name})" : "")} | {card.Left:0.#},{card.Top:0.#} {card.Width:0.#}×{card.Height:0.#}\n");
            foreach (var s in grouped.GetValueOrDefault(card, []).OrderBy(s => s.Top).ThenBy(s => s.Left)) Line(s);
        }
        if (grouped.GetValueOrDefault(Outside) is { Count: > 0 } rest)
        {
            sb.Append(cards.Count > 0 ? "[카드 밖]\n" : "");
            foreach (var s in rest.OrderBy(s => s.Top).ThenBy(s => s.Left)) Line(s);
        }
        return sb.ToString();
    }

    private static readonly ShapeEntry Outside = new("", "", "", 0, 0, 0, 0, 0, false, "");

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
            // 이름을 바꾼 상자는 원본에 그 이름이 없으므로 Id로 찾는다(복제해도 Id는 같다).
            string baselineAddress = shape;
            if (!shape.StartsWith('@'))
            {
                var entry = (await Task.Run(() => ctx.Ppt.ListShapes(ctx.WorkingPath, slide), ct))
                    .FirstOrDefault(e => e.Address == shape || e.Name == shape);
                if (entry is { Id: > 0 }) baselineAddress = $"@{entry.Id}";
            }
            var lines = await Task.Run(() => ctx.Ppt.GetTextLineBreaks(original, ctx.Origins.OriginalIndexOf(slide), baselineAddress), ct);
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

    private static Task<string> SetParagraphs(AgentToolContext ctx, JsonElement args, CancellationToken ct) =>
        FillAsync(ctx, ToolArgs.Int(args, "slide"), ToolArgs.String(args, "shape"),
            ParagraphFitPlanner.SplitPilcrows(ToolArgs.StringArray(args, "paragraphs")), ct);

    /// <summary>
    /// qwen3는 한 턴에 도구를 하나씩만 불러 이름 바꾸기에만 80턴을 썼다(2026-10-06 실측). 묶음 도구로 턴을 줄인다.
    /// 상자 하나가 실패해도 나머지는 계속하고, 실패는 결과에 적는다.
    /// </summary>
    private static async Task<string> SetTexts(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        var sb = new StringBuilder();
        foreach (var item in ToolArgs.ObjectArray(args, "items"))
        {
            string shape = ToolArgs.String(item, "shape");
            try { sb.Append(await FillAsync(ctx, slide, shape, ParagraphFitPlanner.SplitPilcrows(ToolArgs.StringArray(item, "paragraphs")), ct)); }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) { sb.Append($"{shape}: 오류 — {ex.GetBaseException().Message}\n"); }
        }
        return sb.ToString();
    }

    private static async Task<string> RenameShapes(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int slide = ToolArgs.Int(args, "slide");
        var done = new List<string>();
        var failed = new List<string>();
        foreach (var item in ToolArgs.ObjectArray(args, "items"))
        {
            string shape = ToolArgs.String(item, "shape"), name = ToolArgs.String(item, "new_name");
            if (shape == name) continue;
            try
            {
                await Task.Run(() => ctx.Ppt.RenameShape(ctx.WorkingPath, slide, shape, name), ct);
                done.Add($"{shape}→{name}");
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) { failed.Add($"{shape}: {ex.GetBaseException().Message}"); }
        }
        if (done.Count > 0)
        {
            ctx.Origins.MarkTouched(slide);
            ctx.Changes.Add($"{slide}쪽 도형 {done.Count}개 이름 변경");
        }
        return $"{done.Count}개 이름을 바꿨습니다." + (failed.Count == 0 ? "" : "\n실패:\n" + string.Join("\n", failed));
    }

    /// <summary>
    /// 빈 문단과 상자 끝 Enter를 지운다. 끝 Enter는 화면에서 보이지 않고 COM 문단 수에도 안 잡혀,
    /// "38~41쪽 공백줄 삭제해줘"에 모델이 무엇을 지울지 몰라 되묻기만 했다(2026-10-07).
    /// </summary>
    private static async Task<string> RemoveBlankLines(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        int from = ToolArgs.Int(args, "from_slide"), to = ToolArgs.OptionalInt(args, "to_slide") ?? from;
        if (to < from) (from, to) = (to, from);
        var report = new StringBuilder();
        int total = 0;
        for (int slide = from; slide <= to; slide++)
        {
            var shapes = await Task.Run(() => ctx.Ppt.ListShapes(ctx.WorkingPath, slide), ct);
            var edits = new List<ParagraphEdit>();
            int count = 0;
            foreach (var shape in shapes.Where(s => s.HasText))
            {
                var info = await Task.Run(() => ctx.Ppt.GetShapeInfo(ctx.WorkingPath, slide, shape.Address), ct);
                if (info is null || string.IsNullOrWhiteSpace(info.FullText)) continue;
                var plan = BlankLinePlan(slide, shape.Address, info.FullText);
                if (plan.Removed == 0) continue;
                edits.AddRange(plan.Edits);
                count += plan.Removed;
            }
            if (edits.Count == 0) continue;
            var result = await Task.Run(() => ctx.Ppt.SetParagraphTexts(ctx.WorkingPath, edits, markRed: false), ct);
            ctx.Origins.MarkTouched(slide);
            ctx.Changes.Add($"{slide}쪽 빈 줄 {count}개 삭제");
            report.Append($"{slide}쪽: 빈 줄 {count}개 삭제\n");
            foreach (var failure in result.Failures) report.Append("  실패: " + failure + "\n");
            total += count;
        }
        return total == 0 ? $"{from}~{to}쪽에 지울 빈 줄이 없습니다." : report.ToString();
    }

    /// <summary>
    /// 상자 텍스트에서 지울 빈 문단 편집. 빈 문단은 Delete로, 끝 Enter는 첫 문단을 그대로 다시 써서
    /// SetParagraphTexts의 끝 Enter 정리를 태운다. 상자가 통째로 비면 건드리지 않는다.
    /// </summary>
    public static (IReadOnlyList<ParagraphEdit> Edits, int Removed) BlankLinePlan(int slide, string address, string fullText)
    {
        var parts = fullText.Split('\r');
        bool trailing = parts.Length > 1 && parts[^1].Length == 0;
        int countable = trailing ? parts.Length - 1 : parts.Length;     // COM이 세는 문단
        int trailingBlanks = 0;
        for (int i = countable - 1; i >= 0 && string.IsNullOrWhiteSpace(parts[i]); i--) trailingBlanks++;
        if (trailingBlanks == countable) return ([], 0);
        var edits = new List<ParagraphEdit>();
        for (int i = 0; i < countable; i++)
            if (string.IsNullOrWhiteSpace(parts[i]))
                edits.Add(new ParagraphEdit(slide, address, i + 1, "", ParagraphEditKind.Delete));
        int removed = edits.Count + (trailing ? 1 : 0);
        if (trailing && edits.Count == 0)
            edits.Add(new ParagraphEdit(slide, address, 1, parts[0]));
        return (edits, removed);
    }

    private static async Task<string> FillAsync(
        AgentToolContext ctx, int slide, string shape, IReadOnlyList<string> paragraphs, CancellationToken ct)
    {
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
        var targets = new List<string> { shape };
        if (ToolArgs.OptionalBool(args, "with_contents") == true)
        {
            // 카드 배경만 지우고 안의 글 상자를 남기는 일이 있었다(2026-10-06 실측). 중심이 영역 안에 있는 도형을 함께 지운다.
            var shapes = await Task.Run(() => ctx.Ppt.ListShapes(ctx.WorkingPath, slide), ct);
            var container = shapes.FirstOrDefault(s => s.Address == shape || s.Name == shape)
                ?? throw new AgentToolException($"{slide}쪽에서 도형 '{shape}'를 찾지 못했습니다.");
            targets.AddRange(InsideOf(container, shapes).Select(s => s.Id > 0 ? $"@{s.Id}" : s.Address));
        }
        // 이름 주소가 밀리지 않게 Id 주소로, 안쪽부터 지운다.
        foreach (var target in Enumerable.Reverse(targets))
            await Task.Run(() => ctx.Ppt.DeleteShape(ctx.WorkingPath, slide, target), ct);
        ctx.Origins.MarkTouched(slide);
        ctx.AddedShapes.Remove((slide, shape));
        ctx.Changes.Add(targets.Count == 1 ? $"{slide}쪽 {shape} 삭제" : $"{slide}쪽 {shape}와 안의 도형 {targets.Count - 1}개 삭제");
        return targets.Count == 1
            ? $"{shape}를 지웠습니다."
            : $"{shape}와 그 안의 도형 {targets.Count - 1}개를 지웠습니다.";
    }

    /// <summary>중심이 <paramref name="container"/> 영역 안에 있는 다른 도형.</summary>
    public static IReadOnlyList<ShapeEntry> InsideOf(ShapeEntry container, IReadOnlyList<ShapeEntry> shapes) =>
        shapes.Where(s => !ReferenceEquals(s, container) && s.Address != container.Address)
              .Where(s =>
              {
                  float cx = s.Left + s.Width / 2, cy = s.Top + s.Height / 2;
                  return cx >= container.Left && cx <= container.Left + container.Width &&
                         cy >= container.Top && cy <= container.Top + container.Height &&
                         s.Width * s.Height < container.Width * container.Height;
              })
              .ToList();

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

    private static async Task<string> FillCards(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        // 모델이 요약해 넘긴 문장보다 사용자 원문이 정확하다. 원문이 있으면 그것으로 사양을 뽑는다.
        string request = ctx.UserRequest is { Length: > 0 } original ? original
            : ToolArgs.OptionalString(args, "request") ?? throw new AgentToolException("request(사용자 요청)가 필요합니다.");
        if (ctx.Sources is null || ctx.Sources.Outline(null).StartsWith("첨부된 원문이 없습니다", StringComparison.Ordinal))
        {
            ctx.FinishSummary = "카드에 채울 원문이 없습니다. [원문 첨부]로 보고서 파일이나 폴더를 붙이거나, 요청에 폴더 경로를 적어 다시 보내 주세요.";
            return ctx.FinishSummary;
        }
        var report = await new Cards.CardPipeline(ctx).RunFromRequestAsync(request, ct);
        // 파이프라인 뒤에 qwen 루프가 남은 위반을 고치게 두면 이름·패널·문구를 망쳤다(실행 11, 35만 토큰).
        // 결과를 그대로 사용자에게 보여주고 끝낸다. 추가 수정은 사용자가 짚어서 요청한다.
        ctx.FinishSummary = ctx.PendingCards is not null
            ? "작성 요청문을 준비했습니다. [복사]해서 사용하는 고급 모델(ChatGPT·Claude 등)에 붙여넣고, 받은 답을 붙여넣어 [적용]해 주세요.\n" + report
            : "카드 작성을 마쳤습니다. 남은 위반이 있으면 상자를 짚어 후속 요청해 주세요.\n" + report;
        return report;
    }

    /// <summary>요점 정리에 넘길 절을 찾는 질의. 국가 R&D 기획보고서의 공통 절 이름이다.</summary>
    public static readonly IReadOnlyList<string> BriefQueries =
        ["요약문", "연구개발 개요", "목적 필요성", "현황 문제점 한계", "연구개발 내용", "최종목표", "세부 연구내용"];

    public const string DefaultBriefFocus =
        "개요 1문장(과제가 무엇을 개발하는지), 문제점 2~3개(현행 한계), 연구개발 내용 5~6개(개발 항목)";

    private static async Task<string> BriefSource(AgentToolContext ctx, JsonElement args, CancellationToken ct)
    {
        var library = ctx.Sources ?? throw new AgentToolException("첨부된 원문이 없습니다.");
        string document = ToolArgs.String(args, "document");
        string focus = ToolArgs.OptionalString(args, "focus") ?? DefaultBriefFocus;
        var excerpts = await Task.Run(() => library.Excerpts(document, BriefQueries, pagesPerQuery: 2, maxChars: 14000), ct);
        if (excerpts.Length == 0) return "관련 절을 찾지 못했습니다. source_outline으로 목차를 보세요.";
        if (ctx.Summarize is null) return excerpts;

        string instruction =
            "아래는 국가 R&D 기획보고서의 발췌다. 발표 카드에 넣을 내용을 정리하라.\n" +
            $"항목: {focus}\n" +
            "규칙: 발췌에 있는 사실만 쓴다. 각 항목은 개조식 한 줄(30~45자), '~개발', '~구축', '~한계'처럼 명사형으로 끝낸다. " +
            "항목마다 근거 쪽을 (p.N)으로 붙인다. 머리말·설명 없이 항목만 쓴다.";
        var summary = await ctx.Summarize(instruction, excerpts, ct);
        return $"{excerpts.Split('\n')[0]} 요점\n{summary}";
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
