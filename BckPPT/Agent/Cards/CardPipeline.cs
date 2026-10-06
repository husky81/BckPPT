using System.Text;
using System.Text.Json;
using BckPPT.Agent.Tools;
using BckPPT.Models;
using BckPPT.PowerPoint;

namespace BckPPT.Agent.Cards;

/// <summary>
/// 카드 채우기 파이프라인. 긴 도구 루프 대신 정해진 단계로 돈다.
/// <list type="number">
/// <item>템플릿 카드 구조와 상자별 글자 예산을 코드로 잰다(LLM 없음)</item>
/// <item>원문마다 요점 정리(qwen, 문서당 1회)</item>
/// <item>카드 문안 작성(작성 모델, 전체 1회)</item>
/// <item>복제·카드 삭제·패널 폭·이름·채우기·태그·제목·머리말·범례(코드)</item>
/// <item>넘친 상자만 길이 맞춤(qwen, 상자당 단발 호출, 최대 3바퀴)</item>
/// </list>
/// 루프는 매 턴 대화 전체를 다시 보내 2쪽에 100만 토큰 이상 들었다(2026-10-06 실측). 이 흐름은 필요한 만큼만 보낸다.
/// </summary>
public sealed class CardPipeline(AgentToolContext ctx)
{
    public const int MaxFitRounds = 2;

    /// <summary>사용자 요청 문장에서 사양을 뽑아 실행한다. 사양 추출은 작성 모델 단발 호출(약 3천 토큰)이다.</summary>
    public async Task<string> RunFromRequestAsync(string userRequest, CancellationToken ct)
    {
        var ask = ctx.Write ?? ctx.Summarize ?? throw new AgentToolException("사양을 뽑을 LLM이 없습니다.");
        Status("요청에서 카드 사양을 뽑는 중");
        string documents = ctx.Sources?.Outline(null) ?? "(첨부 원문 없음)";
        string request = CardJobExtractor.Request(userRequest, documents);
        CardJob? job = null;
        string? problem = null;
        for (int attempt = 1; attempt <= 2 && job is null; attempt++)
        {
            var answer = await ask(CardJobExtractor.System, problem is null ? request : $"{request}\n\n앞의 답 문제: {problem}. 고쳐서 다시 답한다.", ct);
            try
            {
                var parsed = CardJobExtractor.Parse(answer);
                var missing = CardJobExtractor.MissingNumbers(userRequest, parsed);
                if (missing.Count > 0 && attempt < 2) { problem = $"요청의 카드 {string.Join(", ", missing)}번이 빠졌다"; continue; }
                job = parsed;
            }
            catch (Exception ex) when (ex is FormatException or JsonException && attempt < 2) { problem = ex.Message; }
        }
        if (job is null) throw new AgentToolException($"요청에서 카드 사양을 뽑지 못했습니다: {problem}");
        return $"사양: {CardJobExtractor.Describe(job)}\n" + await RunAsync(job, ct);
    }

    public async Task<string> RunAsync(CardJob job, CancellationToken ct)
    {
        var report = new StringBuilder();
        var ppt = ctx.Ppt;
        string path = ctx.WorkingPath;

        // 1. 템플릿 구조와 예산
        Status("템플릿 카드 구조를 읽는 중");
        _templateSlide = job.TemplateSlide;
        var templateShapes = await Task.Run(() => ppt.ListShapes(path, job.TemplateSlide), ct);
        var template = CardLayoutAnalyzer.Analyze(templateShapes);
        if (template.Cards.Count == 0)
            throw new AgentToolException($"{job.TemplateSlide}쪽에서 카드를 찾지 못했습니다. 카드는 텍스트 없는 세로로 긴 사각형입니다.");
        var model = template.Cards[0];
        if (model.Sections.Count == 0) throw new AgentToolException("템플릿 카드에서 소제목·본문 쌍을 찾지 못했습니다.");
        var titleBudget = model.Title is null
            ? new TextBudget("title", "제목", 2, 18, 1)
            : TextBudget.Measure("title", "제목", await Lines(job.TemplateSlide, model.Title, ct));
        var budgets = new List<TextBudget>();
        foreach (var section in model.Sections)
            budgets.Add(TextBudget.Measure(section.Role, section.Label, await Lines(job.TemplateSlide, section.Body, ct)));
        // 판정은 같은 자리 템플릿 카드와 비교하므로(Id가 같다) 길이 맞춤도 자리별 예산을 쓴다.
        // 1번 카드 예산만 쓰니 3번 자리 문제점(템플릿 3줄)에 4줄이 들어가 위반으로 남았다(실행 11).
        var positionBudgets = new List<(TextBudget Title, IReadOnlyList<TextBudget> Sections)>();
        foreach (var card in template.Cards)
        {
            var title = card.Title is null ? titleBudget : TextBudget.Measure("title", "제목", await Lines(job.TemplateSlide, card.Title, ct));
            var sections = new List<TextBudget>();
            foreach (var section in card.Sections)
                sections.Add(TextBudget.Measure(section.Role, section.Label, await Lines(job.TemplateSlide, section.Body, ct)));
            positionBudgets.Add((title, sections));
        }
        var sample = await SampleAsync(job.TemplateSlide, model, ct);
        int maxCards = template.Cards.Count;
        var tooMany = job.Slides.Select((s, i) => (s, i)).Where(x => x.s.Cards.Count > maxCards).ToList();
        if (tooMany.Count > 0) throw new AgentToolException($"템플릿 카드는 {maxCards}개입니다. 슬라이드 {tooMany[0].i + 1}에 카드 {tooMany[0].s.Cards.Count}개를 넣을 수 없습니다.");
        report.Append($"템플릿 {job.TemplateSlide}쪽: 카드 {maxCards}개, 예산 {titleBudget.Describe().Replace("제목: ", "제목 ")}; ")
              .Append(string.Join("; ", budgets.Select(b => b.Describe()))).Append('\n');

        // 2. 원문 요점(문서마다 한 번, 같은 문서는 재사용)
        var allCards = job.Slides.SelectMany(s => s.Cards).ToList();
        var briefs = await BriefAllAsync(allCards, budgets, ct);

        // 3. 문안 작성(한 번)
        Status($"카드 {allCards.Count}장 문안을 쓰는 중");
        var drafts = await WriteAsync(allCards, briefs, titleBudget, budgets, sample, report, ct);

        // 4. 배치
        var filled = new List<(int Slide, string Address, TextBudget Budget, IReadOnlyList<string> Items)>();
        int insertAfter = job.InsertAfter;
        for (int s = 0; s < job.Slides.Count; s++)
        {
            var spec = job.Slides[s];
            Status($"{s + 1}/{job.Slides.Count}번째 슬라이드를 만드는 중");
            int slide = await Task.Run(() => ppt.DuplicateSlide(path, job.TemplateSlide, insertAfter), ct);
            ctx.Origins.RecordDuplicate(job.TemplateSlide, slide);
            if (job.TemplateSlide > insertAfter) job = job with { TemplateSlide = job.TemplateSlide + 1 };
            _templateSlide = job.TemplateSlide;
            insertAfter = slide;

            var layout = CardLayoutAnalyzer.Analyze(await Task.Run(() => ppt.ListShapes(path, slide), ct));
            await TrimCardsAsync(slide, layout, spec.Cards.Count, ct);
            if (job.Header is { Length: > 0 } header && layout.Header is { } headerShape)
                await FillAsync(slide, CardLayoutAnalyzer.Addr(headerShape), [header], ct);
            if (spec.Title is { Length: > 0 } title) await RetitleAsync(slide, layout, title, ct);
            if (job.Legend is { Length: > 0 } legend) await AddLegendAsync(slide, template, layout, legend, ct);

            for (int c = 0; c < spec.Cards.Count; c++)
            {
                var card = layout.Cards[c];
                var cardSpec = spec.Cards[c];
                var draft = drafts.FirstOrDefault(d => d.Number == cardSpec.Number) ?? drafts.ElementAtOrDefault(allCards.IndexOf(cardSpec));
                var (positionTitle, positionSections) = positionBudgets[Math.Min(c, positionBudgets.Count - 1)];
                filled.AddRange(await FillCardAsync(slide, c + 1, card, cardSpec, draft, positionTitle,
                    positionSections.Count > 0 ? positionSections : budgets, ct));
            }
            await UnifySpacingAsync(slide, template, ct);
            ctx.Changes.Add($"{slide}쪽 카드 {string.Join("·", spec.Cards.Select(x => x.Number))} 작성");
            report.Append($"{slide}쪽: 카드 {string.Join("·", spec.Cards.Select(x => x.Number))}\n");
        }

        // 5. 길이 맞춤(넘친 상자만)
        await FitAsync(filled, report, ct);

        // 6. 판정
        foreach (var slide in filled.Select(f => f.Slide).Distinct())
        {
            var issues = await PptAgentTools.LintAsync(ctx, slide, null, ct);
            report.Append($"{slide}쪽 {LayoutLint.Describe(issues)}\n");
        }
        if (ctx.Meter is { } meter) report.Append("LLM 사용: ").Append(Llm.UsageMeter.Describe(meter.Snapshot())).Append('\n');
        return report.ToString();
    }

    // ── 단계별 ─────────────────────────────────────────────────

    private async Task<Dictionary<string, string>> BriefAllAsync(
        IReadOnlyList<CardSpec> cards, IReadOnlyList<TextBudget> budgets, CancellationToken ct)
    {
        var briefs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var documents = cards.Select(c => c.Document).Where(d => !string.IsNullOrWhiteSpace(d)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (documents.Count == 0 || ctx.Sources is null) return briefs;
        string focus = string.Join(", ", budgets.Select(b => $"{b.Label} {Math.Max(b.Paragraphs, 2)}~{b.Paragraphs + 2}개"));
        using var gate = new SemaphoreSlim(3);
        var tasks = documents.Select(async document =>
        {
            await gate.WaitAsync(ct);
            try
            {
                Status($"원문 {document} 요점 정리 중");
                var excerpts = await Task.Run(() => ctx.Sources.Excerpts(document!, PptAgentTools.BriefQueries, 2, 14000), ct);
                if (excerpts.Length == 0 || ctx.Summarize is null) return (document!, excerpts);
                string instruction =
                    "아래는 국가 R&D 기획보고서의 발췌다. 발표 카드에 넣을 내용을 정리하라.\n" +
                    $"항목: {focus}\n" +
                    "규칙: 발췌에 있는 사실만 쓴다. 각 항목은 개조식 한 줄, 명사형으로 끝낸다. 항목마다 근거 쪽을 (p.N)으로 붙인다. 머리말 없이 항목만 쓴다.";
                return (document!, await ctx.Summarize(instruction, excerpts, ct));
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex) { return (document!, $"(요점 정리 실패: {ex.Message})"); }
            finally { gate.Release(); }
        });
        foreach (var (document, brief) in await Task.WhenAll(tasks)) briefs[document] = brief;
        return briefs;
    }

    private async Task<IReadOnlyList<CardDraft>> WriteAsync(
        IReadOnlyList<CardSpec> cards, IReadOnlyDictionary<string, string> briefs, TextBudget title,
        IReadOnlyList<TextBudget> budgets, CardLayoutSample? sample, StringBuilder report, CancellationToken ct)
    {
        var write = ctx.Write ?? ctx.Summarize ?? throw new AgentToolException("문안을 쓸 LLM이 없습니다.");
        var request = CardFillPrompts.WriterRequest(
            cards.Select(c => (c, c.Document is { } d && briefs.TryGetValue(d, out var b) ? b : "")).ToList(),
            title, budgets, sample);
        var roles = budgets.Select(b => b.Role).ToList();
        for (int attempt = 1; ; attempt++)
        {
            var answer = await write(CardFillPrompts.WriterSystem, request, ct);
            try { return CardFillPrompts.ParseDrafts(answer, roles); }
            catch (Exception ex) when (ex is FormatException or System.Text.Json.JsonException && attempt < 2)
            {
                report.Append($"문안 JSON 해석 실패, 다시 요청: {ex.Message}\n");
                request += "\n\n앞의 답은 JSON으로 읽을 수 없었다. 형식대로 JSON 하나만 답한다.";
            }
        }
    }

    private async Task<IReadOnlyList<(int, string, TextBudget, IReadOnlyList<string>)>> FillCardAsync(
        int slide, int position, CardLayout card, CardSpec spec, CardDraft? draft,
        TextBudget titleBudget, IReadOnlyList<TextBudget> budgets, CancellationToken ct)
    {
        var filled = new List<(int, string, TextBudget, IReadOnlyList<string>)>();
        string prefix = $"C{position}";
        var renames = new List<(ShapeEntry Shape, string Name)> { (card.Container, $"{prefix}_card") };
        if (card.Number is { } number)
        {
            await FillAsync(slide, CardLayoutAnalyzer.Addr(number), [spec.Number], ct);
            renames.Add((number, $"{prefix}_num"));
        }
        if (card.Tag is { } tag)
        {
            if (spec.Tag is { Length: > 0 } tagText)
            {
                await FillAsync(slide, CardLayoutAnalyzer.Addr(tag), [tagText], ct);
                // 태그는 한 줄이어야 한다. 카드 오른쪽 끝까지 넓힌다(템플릿 태그 상자는 "성능 평가"용으로 좁다).
                float width = card.Right - tag.Left - 6;
                if (width > tag.Width)
                    await Task.Run(() => ctx.Ppt.SetShapeGeometry(ctx.WorkingPath, slide, CardLayoutAnalyzer.Addr(tag), new ShapeGeometry(Width: width)), ct);
            }
            renames.Add((tag, $"{prefix}_tag"));
        }
        if (card.Title is { } title)
        {
            string text = draft?.Title is { Length: > 0 } drafted ? drafted : spec.Title;
            await FillAsync(slide, CardLayoutAnalyzer.Addr(title), [text], ct);
            filled.Add((slide, CardLayoutAnalyzer.Addr(title), titleBudget, [text]));
            renames.Add((title, $"{prefix}_title"));
        }
        foreach (var section in card.Sections)
        {
            var budget = budgets.FirstOrDefault(b => b.Role == section.Role) ?? budgets[0];
            IReadOnlyList<string> items = draft?.Sections.GetValueOrDefault(section.Role) is { Count: > 0 } drafted ? drafted : [""];
            await FillAsync(slide, CardLayoutAnalyzer.Addr(section.Body), items, ct);
            filled.Add((slide, CardLayoutAnalyzer.Addr(section.Body), budget, items));
            renames.Add((section.LabelShape, $"{prefix}_{section.Role}_label"));
            renames.Add((section.Body, $"{prefix}_{section.Role}"));
        }
        // 이름은 끝에 붙인다. 채우는 동안에는 Id 주소를 쓰므로 이름이 겹쳐도 상관없다.
        foreach (var (shape, name) in renames)
        {
            try { await Task.Run(() => ctx.Ppt.RenameShape(ctx.WorkingPath, slide, CardLayoutAnalyzer.Addr(shape), name), ct); }
            catch (OperationCanceledException) { throw; }
            catch { /* 이름은 편의일 뿐이라 실패해도 계속한다. */ }
        }
        return filled;
    }

    /// <summary>쓰지 않는 카드를 오른쪽부터 안의 도형까지 지우고, 패널과 머리말 폭을 남은 카드에 맞춘다.</summary>
    private async Task TrimCardsAsync(int slide, SlideCards layout, int keep, CancellationToken ct)
    {
        if (keep >= layout.Cards.Count) return;
        foreach (var card in layout.Cards.Skip(keep).Reverse())
        {
            foreach (var shape in card.Members.Append(card.Container))
                await Task.Run(() => ctx.Ppt.DeleteShape(ctx.WorkingPath, slide, CardLayoutAnalyzer.Addr(shape)), ct);
        }
        if (layout.Panel is { } panel && keep > 0)
        {
            float margin = layout.Cards[0].Container.Left - panel.Left;
            float width = layout.Cards[keep - 1].Right + margin - panel.Left;
            await Task.Run(() => ctx.Ppt.SetShapeGeometry(ctx.WorkingPath, slide, CardLayoutAnalyzer.Addr(panel), new ShapeGeometry(Width: width)), ct);
            if (layout.Header is { } header && header.Left + header.Width > panel.Left + width)
                await Task.Run(() => ctx.Ppt.SetShapeGeometry(ctx.WorkingPath, slide, CardLayoutAnalyzer.Addr(header),
                    new ShapeGeometry(Width: panel.Left + width - header.Left)), ct);
        }
        ctx.Changes.Add($"{slide}쪽 카드 {layout.Cards.Count - keep}개 삭제, 패널 폭 조정");
    }

    /// <summary>
    /// 슬라이드 제목을 바꾼다. "4. 지정공모형 후보 아이템_우선순위…"처럼 앞부분과 뒷부분 글자 크기가 다른 제목이 많아,
    /// 같은 앞부분은 두고 달라지는 뒷부분만 바꿔 서식을 지킨다.
    /// </summary>
    private async Task RetitleAsync(int slide, SlideCards layout, string title, CancellationToken ct)
    {
        var shapes = await Task.Run(() => ctx.Ppt.ListShapes(ctx.WorkingPath, slide), ct);
        float limit = (layout.Header?.Top ?? layout.Panel?.Top ?? 80) - 5;
        var titleShape = shapes.Where(s => s.HasText && s.Top < limit).OrderByDescending(s => s.Width).FirstOrDefault();
        if (titleShape is null) return;
        string address = CardLayoutAnalyzer.Addr(titleShape);
        var info = await Task.Run(() => ctx.Ppt.GetShapeInfo(ctx.WorkingPath, slide, address), ct);
        string current = info?.FullText.TrimEnd('\r') ?? "";
        if (current == title) return;
        int common = 0;
        while (common < current.Length && common < title.Length && current[common] == title[common]) common++;
        if (common == current.Length) common = Math.Max(0, common - 1);   // 뒤에 덧붙이는 경우 마지막 글자부터 바꾼다
        string find = current[common..], replace = title[common..];
        if (find.Length == 0 || await Task.Run(() => ctx.Ppt.ReplaceTextInShape(ctx.WorkingPath, slide, address, find, replace), ct) == 0)
            await FillAsync(slide, address, [title], ct);
        ctx.Changes.Add($"{slide}쪽 제목 변경");
    }

    private async Task AddLegendAsync(int slide, SlideCards template, SlideCards layout, string legend, CancellationToken ct)
    {
        // 템플릿 패널의 오른쪽 끝에 맞춘다(카드를 줄인 쪽도 범례 위치는 같게).
        var panel = template.Panel ?? layout.Panel;
        if (panel is null) return;
        float height = layout.Header?.Height ?? 20;
        // 머리말과 같은 줄에 두되 패널과 겹치지 않게 패널 바로 위에서 끝낸다(겹침 판정 대상).
        float top = Math.Min(layout.Header?.Top ?? float.MaxValue, panel.Top - height - 1);
        var request = new TextShapeRequest(legend, panel.Left, top, 200, height, "범례", TextShapeKind.RoundedRectangle,
            FillColor: "#3B3B3B", FontSize: 9, Bold: true, FontColor: "#FFFFFF", AutoWidth: true, RightEdge: panel.Left + panel.Width);
        string name = await Task.Run(() => ctx.Ppt.AddTextShape(ctx.WorkingPath, slide, request), ct);
        ctx.AddedShapes.Add((slide, name));
    }

    private async Task FitAsync(
        List<(int Slide, string Address, TextBudget Budget, IReadOnlyList<string> Items)> filled, StringBuilder report, CancellationToken ct)
    {
        if (ctx.Summarize is null) return;
        int fixedCount = 0;
        for (int round = 1; round <= MaxFitRounds; round++)
        {
            var over = new List<int>();
            var measured = new Dictionary<int, IReadOnlyList<TextLineInfo>>();
            for (int i = 0; i < filled.Count; i++)
            {
                var f = filled[i];
                if (f.Items.All(string.IsNullOrWhiteSpace)) continue;
                var lines = await Task.Run(() => ctx.Ppt.GetTextLineBreaks(ctx.WorkingPath, f.Slide, f.Address), ct);
                if (Violates(lines, f.Budget)) { over.Add(i); measured[i] = lines; }
            }
            if (over.Count == 0) break;
            Status($"길이 맞춤 {round}바퀴: 상자 {over.Count}개");
            foreach (var i in over)
            {
                var f = filled[i];
                try
                {
                    var answer = await ctx.Summarize(CardFillPrompts.FitSystem,
                        CardFillPrompts.FitRequest(f.Budget, f.Items, measured[i]), ct);
                    var items = CardFillPrompts.ParseItems(answer);
                    if (items.Count == 0) continue;
                    await FillAsync(f.Slide, f.Address, items, ct);
                    filled[i] = f with { Items = items };
                    fixedCount++;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) { report.Append($"길이 맞춤 실패 {f.Slide}쪽 {f.Address}: {ex.Message}\n"); }
            }
        }
        report.Append($"길이 맞춤 {fixedCount}회\n");

        // 그래도 줄 수를 넘으면 마지막 항목부터 코드로 뺀다. qwen이 "31자 줄여라"를 3바퀴 동안 못 맞춘 일이 있었다(실행 10).
        for (int i = 0; i < filled.Count; i++)
        {
            var f = filled[i];
            var items = f.Items.ToList();
            var dropped = new List<string>();
            var lines = await Task.Run(() => ctx.Ppt.GetTextLineBreaks(ctx.WorkingPath, f.Slide, f.Address), ct);
            while (lines.Count > f.Budget.Lines && items.Count > 1)
            {
                dropped.Insert(0, items[^1]);
                items.RemoveAt(items.Count - 1);
                await FillAsync(f.Slide, f.Address, items, ct);
                lines = await Task.Run(() => ctx.Ppt.GetTextLineBreaks(ctx.WorkingPath, f.Slide, f.Address), ct);
            }
            if (dropped.Count == 0) continue;
            filled[i] = f with { Items = items };
            report.Append($"{f.Slide}쪽 {f.Budget.Label}: 줄 수를 맞추려고 뺀 항목 — {string.Join(" / ", dropped)}\n");
        }
    }

    public static bool Violates(IReadOnlyList<TextLineInfo> lines, TextBudget budget) =>
        lines.Count > budget.Lines ||
        LayoutLint.Segments(lines).Any(s => s.Count >= 2 && s[^1].Text.Trim().Length is > 0 and <= 4);

    /// <summary>
    /// 같은 역할 본문 상자의 자간을 템플릿 다수 값으로 맞춘다. 착수보고 33쪽 1번 카드 개요만 자간 -1.5였고,
    /// 채운 뒤에도 그대로 남아 서식 이탈로 지적됐다(2026-10-06 사례와 실행 11).
    /// </summary>
    private async Task UnifySpacingAsync(int slide, SlideCards template, CancellationToken ct)
    {
        var layout = CardLayoutAnalyzer.Analyze(await Task.Run(() => ctx.Ppt.ListShapes(ctx.WorkingPath, slide), ct));
        foreach (var role in template.Cards.SelectMany(c => c.Sections).Select(s => s.Role).Distinct())
        {
            var templateSpacing = new List<float>();
            foreach (var body in template.Cards.SelectMany(c => c.Sections).Where(s => s.Role == role).Select(s => s.Body))
            {
                try { templateSpacing.Add((await Task.Run(() => ctx.Ppt.GetTextFormat(ctx.WorkingPath, _templateSlide, CardLayoutAnalyzer.Addr(body)), ct)).Spacing); }
                catch (OperationCanceledException) { throw; }
                catch { }
            }
            var groups = templateSpacing.GroupBy(x => x).OrderByDescending(g => g.Count()).ToList();
            if (groups.Count < 2 || groups[0].Count() == groups[1].Count()) continue;   // 과반이 없으면 손대지 않는다
            float majority = groups[0].Key;
            foreach (var body in layout.Cards.SelectMany(c => c.Sections).Where(s => s.Role == role).Select(s => s.Body))
            {
                string address = CardLayoutAnalyzer.Addr(body);
                var format = await Task.Run(() => ctx.Ppt.GetTextFormat(ctx.WorkingPath, slide, address), ct);
                if (Math.Abs(format.Spacing - majority) > 0.05f)
                    await Task.Run(() => ctx.Ppt.SetTextFormat(ctx.WorkingPath, slide, address, new TextFormatPatch(Spacing: majority)), ct);
            }
        }
    }

    /// <summary>작업 사본에서 템플릿 슬라이드의 현재 번호. 앞에 쪽을 끼우면 밀린다.</summary>
    private int _templateSlide;

    // ── 도우미 ─────────────────────────────────────────────────

    private Task<IReadOnlyList<TextLineInfo>> Lines(int slide, ShapeEntry shape, CancellationToken ct) =>
        Task.Run(() => ctx.Ppt.GetTextLineBreaks(ctx.WorkingPath, slide, CardLayoutAnalyzer.Addr(shape)), ct);

    private async Task FillAsync(int slide, string address, IReadOnlyList<string> paragraphs, CancellationToken ct)
    {
        var info = await Task.Run(() => ctx.Ppt.GetShapeInfo(ctx.WorkingPath, slide, address), ct)
            ?? throw new AgentToolException($"{slide}쪽에서 텍스트 도형 '{address}'를 찾지 못했습니다.");
        var edits = ParagraphFitPlanner.Plan(slide, address, ParagraphFitPlanner.CountParagraphs(info.FullText),
            paragraphs.Count == 0 ? [""] : paragraphs);
        await Task.Run(() => ctx.Ppt.SetParagraphTexts(ctx.WorkingPath, edits, markRed: false), ct);
        ctx.Origins.MarkTouched(slide);
    }

    private async Task<CardLayoutSample?> SampleAsync(int slide, CardLayout card, CancellationToken ct)
    {
        async Task<IReadOnlyList<string>> Text(ShapeEntry s) =>
            ((await Task.Run(() => ctx.Ppt.GetShapeInfo(ctx.WorkingPath, slide, CardLayoutAnalyzer.Addr(s)), ct))?.FullText ?? "")
            .TrimEnd('\r').Split('\r').Select(t => t.Replace('\v', ' ').Trim()).Where(t => t.Length > 0).ToList();
        try
        {
            string title = card.Title is null ? "" : string.Join(" ", await Text(card.Title));
            var sections = new List<(string, IReadOnlyList<string>)>();
            foreach (var section in card.Sections) sections.Add((section.Label, await Text(section.Body)));
            return new CardLayoutSample(title, sections);
        }
        catch (OperationCanceledException) { throw; }
        catch { return null; }
    }

    private void Status(string message) => ctx.Status?.Report(message);
}
