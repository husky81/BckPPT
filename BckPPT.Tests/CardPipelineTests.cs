using BckPPT.Agent.Cards;
using BckPPT.Agent.Tools;
using BckPPT.Models;

namespace BckPPT.Tests;

public class CardFillPromptTests
{
    private static TextLineInfo L(int n, string text, LineBreakType type) => new(n, text, type);

    [Fact]
    public void 예산은_템플릿의_줄_수와_꽉_찬_줄_길이로_정한다()
    {
        var budget = TextBudget.Measure("rnd", "연구개발 내용",
        [
            L(1, "최적 해체·철거 시점 및 공법 선정 의사결", LineBreakType.AutoWrap),
            L(2, "정 모델 개발", LineBreakType.Enter),
            L(3, "전용 철거 장비 개발", LineBreakType.None),
        ]);
        Assert.Equal((3, 23, 2), (budget.Lines, budget.CharsPerLine, budget.Paragraphs));
    }

    [Fact]
    public void 문안_JSON을_역할별로_읽고_글머리를_뗀다()
    {
        const string answer = """
            <think>생각</think>
            ```json
            {"cards": [{"number": "1", "title": "노후 도로시설물 철거 의사결정",
              "overview": ["- 의사결정 시스템 구축."], "rnd": ["1. 선정기법 개발", "• 우선순위 모델 개발"]}]}
            ```
            """;
        var draft = Assert.Single(CardFillPrompts.ParseDrafts(answer, ["overview", "problem", "rnd"]));
        Assert.Equal("1", draft.Number);
        Assert.Equal(["의사결정 시스템 구축"], draft.Sections["overview"]);
        Assert.Equal(["선정기법 개발", "우선순위 모델 개발"], draft.Sections["rnd"]);
        Assert.False(draft.Sections.ContainsKey("problem"));
    }

    [Fact]
    public void 길이_맞춤_답은_배열이나_감싼_배열을_받는다()
    {
        Assert.Equal(["가", "나"], CardFillPrompts.ParseItems("""["가", "나"]"""));
        Assert.Equal(["가"], CardFillPrompts.ParseItems("""{"items": ["- 가."]}"""));
    }

    [Fact]
    public void 길이_맞춤_요청에_측정과_꼬리를_담는다()
    {
        var budget = new TextBudget("problem", "문제점", 2, 20, 2);
        var request = CardFillPrompts.FitRequest(budget, ["가", "나"],
        [
            L(1, "철거 교체 의사결정 체계 부재로 사후적", LineBreakType.AutoWrap),
            L(2, "관리", LineBreakType.Enter),
            L(3, "범용 장비와 분절적 절차", LineBreakType.None),
        ]);
        Assert.Contains("지금 3줄", request);
        Assert.Contains("'관리'", request);
        Assert.Contains("1줄을 줄여야", request);
    }

    [Fact]
    public void 줄_수_초과나_꼬리면_위반이다()
    {
        var budget = new TextBudget("t", "t", 2, 20, 1);
        Assert.True(CardPipeline.Violates([L(1, "a", LineBreakType.Enter), L(2, "b", LineBreakType.Enter), L(3, "c", LineBreakType.None)], budget));
        Assert.True(CardPipeline.Violates([L(1, "긴 문장 앞부분", LineBreakType.AutoWrap), L(2, "개발", LineBreakType.None)], budget));
        Assert.False(CardPipeline.Violates([L(1, "한 줄", LineBreakType.Enter), L(2, "두 줄", LineBreakType.None)], budget));
    }
}

/// <summary>가짜 PowerPoint 위에서 파이프라인 전체를 돌린다. LLM 자리에는 정해진 답을 준다.</summary>
public sealed class CardPipelineTests : IDisposable
{
    private readonly FakePptComService _ppt = new();
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "bckppt-card-tests-" + Guid.NewGuid().ToString("N")[..8]);

    private static IEnumerable<FakeShape> Card(float dx, string number)
    {
        // 카드 배경·머리 띠는 텍스트가 없다. 본문 상자는 한 줄 20자.
        yield return new FakeShape("Card", 42 + dx, 112, 216, 390, 20);
        yield return new FakeShape("Head", 42 + dx, 112, 216, 69, 20);
        yield return new FakeShape("Num", 51 + dx, 117, 27, 27, 5, number);
        yield return new FakeShape("Tag", 81 + dx, 120, 50, 20, 6, "성능 평가");
        yield return new FakeShape("Title", 52 + dx, 144, 206, 35, 20, "옛 과제 제목");
        yield return new FakeShape("Label", 50 + dx, 184, 101, 14, 20, "개요");
        yield return new FakeShape("Body", 51 + dx, 198, 204, 42, 20, "옛 개요 한 줄", "옛 개요 둘째");
        yield return new FakeShape("Label", 50 + dx, 261, 130, 14, 20, "연구개발 내용");
        yield return new FakeShape("Body", 52 + dx, 348, 204, 99, 20, "옛 내용 하나", "옛 내용 둘", "옛 내용 셋");
    }

    public CardPipelineTests()
    {
        Directory.CreateDirectory(_dir);
        var slide = new List<FakeShape>
        {
            new("Panel", 32, 93, 896, 422, 20),
            new("Header", 32, 77, 421, 22, 40, "내역 2. 효율적 시설물 안전관리"),
            new("SlideTitle", 29, 11, 835, 46, 80, "4. 지정공모형 후보 아이템_우선순위 도출 결과(6/7)"),
        };
        slide.AddRange(Card(0, "3"));
        slide.AddRange(Card(220, "4"));
        _ppt.Slides.Add([new FakeShape("Cover", 0, 0, 100, 20, 50, "표지")]);
        _ppt.Slides.Add(slide);
    }

    private AgentToolContext Context(Func<string, string, string> write, Func<string, string, string>? fit = null)
    {
        var context = new AgentToolContext(_ppt, Path.Combine(_dir, "work.pptx"))
        {
            Write = (system, content, _) => Task.FromResult(write(system, content)),
            Summarize = (system, content, _) => Task.FromResult(fit?.Invoke(system, content) ?? "[]"),
        };
        return context;
    }

    [Fact]
    public async Task 복제하고_남는_카드를_지우고_채운다()
    {
        string? writerRequest = null;
        var context = Context((_, request) =>
        {
            writerRequest = request;
            return """
                {"cards": [{"number": "1", "title": "노후 도로시설물 철거 의사결정",
                  "overview": ["의사결정 시스템 구축"], "rnd": ["선정기법 개발", "우선순위 모델 개발"]}]}
                """;
        });
        var job = new CardJob(2, 2, [new CardSlideSpec([new CardSpec("1", "(초)노후 도로시설물", "시공 안전성 ●")],
            "4. 지정공모형 후보 아이템_기존 기획과제 연계(2/2)")], Header: "기존 기획과제 中 직접 연계 과제");

        var report = await new CardPipeline(context).RunAsync(job, CancellationToken.None);

        Assert.Equal(3, _ppt.Slides.Count);
        var created = _ppt.Slides[2];
        Assert.Equal(["1"], created.Single(s => s.Name == "C1_num").Paragraphs);
        Assert.Equal(["시공 안전성 ●"], created.Single(s => s.Name == "C1_tag").Paragraphs);
        Assert.True(created.Single(s => s.Name == "C1_tag").Width > 50);                 // 태그 상자를 넓혔다
        Assert.Equal(["(초)노후 도로시설물"], created.Single(s => s.Name == "C1_title").Paragraphs);   // 제목은 요청한 과제명 그대로
        Assert.Equal(["선정기법 개발", "우선순위 모델 개발"], created.Single(s => s.Name == "C1_rnd").Paragraphs);
        Assert.Equal(1, created.Count(s => s.Name == "C1_card"));
        Assert.DoesNotContain(created, s => s.Paragraphs.Contains("4"));                 // 둘째 카드는 통째로 지웠다
        Assert.Equal(236f, created.Single(s => s.Name == "Panel").Width);                // 원래 결과본 36쪽 패널 폭과 같다
        Assert.Equal(["기존 기획과제 中 직접 연계 과제"], created.Single(s => s.Name == "Header").Paragraphs);
        Assert.Equal(["4. 지정공모형 후보 아이템_기존 기획과제 연계(2/2)"], created.Single(s => s.Name == "SlideTitle").Paragraphs);
        Assert.Equal(["4. 지정공모형 후보 아이템_우선순위 도출 결과(6/7)"], _ppt.Slides[1].Single(s => s.Name == "SlideTitle").Paragraphs);
        Assert.Contains("개요: 전체 2줄 이내", writerRequest);
        Assert.Contains("옛 개요 한 줄", writerRequest);                                   // 템플릿 예시
        Assert.Contains("3쪽: 카드 1", report);
    }

    [Fact]
    public async Task 넘친_상자만_길이를_맞춘다()
    {
        var fitRequests = new List<string>();
        var context = Context(
            (_, _) => """
                {"cards": [{"number": "1", "title": "짧은 제목",
                  "overview": ["철거·교체 대상 선정과 우선순위 평가를 위한 의사결정 시스템을 구축하고 공법 라이브러리를 실증하는 기술"], "rnd": ["선정기법 개발"]}]}
                """,
            (_, request) => { fitRequests.Add(request); return """["의사결정 시스템 구축"]"""; });
        var job = new CardJob(2, 2, [new CardSlideSpec([new CardSpec("1", "과제")])]);

        await new CardPipeline(context).RunAsync(job, CancellationToken.None);

        var request = Assert.Single(fitRequests);                                           // 개요 상자 하나만
        Assert.Contains("상자: 개요", request);
        Assert.Equal(["의사결정 시스템 구축"], _ppt.Slides[2].Single(s => s.Name == "C1_overview").Paragraphs);
    }

    [Fact]
    public async Task 외부_작성은_요청문을_만들고_붙여넣은_답으로_적용한다()
    {
        int summarizeCalls = 0;
        var context = new AgentToolContext(_ppt, Path.Combine(_dir, "work.pptx"))
        {
            // 사양 추출만 작성 모델을 쓴다. 요점 정리·문안은 사용자 모델이 한다.
            Write = (system, content, _) => Task.FromResult(
                """{"template_slide": 2, "after": 2, "slides": [{"cards": [{"number": "1", "title": "노후 도로시설물"}]}]}"""),
            Summarize = (_, _, _) => { summarizeCalls++; return Task.FromResult("[]"); },
            ExternalWriter = true,
        };
        var pipeline = new CardPipeline(context);

        var report = await pipeline.RunFromRequestAsync("2쪽 카드로 2쪽 뒤에 1번 노후 도로시설물 카드를 만들어줘", CancellationToken.None);

        var pending = Assert.IsType<PendingCards>(context.PendingCards);
        Assert.Contains("작성 요청문을 준비했습니다", report);
        Assert.Contains("개요: 전체 2줄 이내", pending.Prompt);
        Assert.Contains("JSON 하나만", pending.Prompt);
        Assert.Equal(2, _ppt.Slides.Count);          // 답을 받기 전에는 슬라이드를 건드리지 않는다
        Assert.Equal(0, summarizeCalls);

        await Assert.ThrowsAsync<AgentToolException>(() => pipeline.ApplyExternalAsync(pending, "죄송합니다만…", CancellationToken.None));
        Assert.Equal(2, _ppt.Slides.Count);

        const string fence = "```";
        await pipeline.ApplyExternalAsync(pending, fence + "json\n" +
            """{"cards": [{"number": "1", "title": "노후 도로시설물 철거", "overview": ["의사결정 시스템 구축"], "rnd": ["선정기법 개발"]}]}"""
            + "\n" + fence, CancellationToken.None);
        Assert.Equal(3, _ppt.Slides.Count);
        Assert.Equal(["의사결정 시스템 구축"], _ppt.Slides[2].Single(s => s.Name == "C1_overview").Paragraphs);
    }

    [Fact]
    public async Task 템플릿보다_카드가_많으면_여러_쪽으로_나눈다()
    {
        var context = Context((_, _) => """{"cards": [{"number": "1", "title": "가"}, {"number": "2", "title": "나"}, {"number": "3", "title": "다"}]}""");
        var cards = Enumerable.Range(1, 3).Select(i => new CardSpec($"{i}", $"과제{i}")).ToList();

        var report = await new CardPipeline(context).RunAsync(new CardJob(2, 2, [new CardSlideSpec(cards)]), CancellationToken.None);

        Assert.Equal(4, _ppt.Slides.Count);                                   // 템플릿 카드 2장 → 2+1로 두 쪽
        Assert.Contains("3쪽: 카드 1·2", report);
        Assert.Contains("4쪽: 카드 3", report);
        Assert.Equal(["과제3"], _ppt.Slides[3].Single(s => s.Name == "C1_title").Paragraphs);
    }

    [Fact]
    public void 문안_답에_cards가_없으면_형식_오류다()
    {
        Assert.Throws<FormatException>(() => CardFillPrompts.ParseDrafts("{}", ["overview"]));
    }

    public void Dispose()
    {
        try { Directory.Delete(_dir, true); } catch { }
    }
}

public class CardJobExtractorTests
{
    private const string Request =
        "33쪽 카드 레이아웃을 복제해 34쪽 뒤에 2쪽을 만들어줘. 35쪽: 1번 노후 도로시설물, 6번 AI 기반 TBM. 36쪽: 11번 디지털 트윈";

    [Fact]
    public void 사양_JSON을_읽는다()
    {
        const string answer = """
            {"template_slide": "33", "after": 34, "slides": [
              {"title": "연계(1/2)", "cards": [{"number": 1, "title": "노후 도로시설물", "tag": "시공 안전성 ●", "document": "D1"},
                                              {"number": "6", "title": "AI 기반 TBM", "document": "D2"}]},
              {"cards": [{"number": "11", "title": "디지털 트윈"}]}],
             "header": "기존 기획과제 中 직접 연계 과제"}
            """;
        var job = CardJobExtractor.Parse(answer);
        Assert.Equal((33, 34, 2), (job.TemplateSlide, job.InsertAfter, job.Slides.Count));
        Assert.Equal(["1", "6"], job.Slides[0].Cards.Select(c => c.Number));
        Assert.Equal("D1", job.Slides[0].Cards[0].Document);
        Assert.Null(job.Slides[1].Title);
        Assert.Equal("기존 기획과제 中 직접 연계 과제", job.Header);
        Assert.Empty(CardJobExtractor.MissingNumbers(Request, job));
    }

    [Fact]
    public void 요청의_카드_번호가_빠지면_알려준다()
    {
        // 실행 10: 카드 5장을 쪽마다 1장으로 줄였다.
        var job = CardJobExtractor.Parse("""{"template_slide": 33, "after": 34, "slides": [{"cards": [{"number": "1", "title": "x"}]}, {"cards": [{"number": "11", "title": "y"}]}]}""");
        Assert.Equal(["6"], CardJobExtractor.MissingNumbers(Request, job));
    }

    [Fact]
    public void 쪽_번호가_없으면_거부한다()
    {
        Assert.Throws<FormatException>(() => CardJobExtractor.Parse("""{"slides": [{"cards": [{"number": "1", "title": "x"}]}]}"""));
    }
}

public class CardTitleTests
{
    [Theory]
    [InlineData("건설로보틱스 시공자동화 기술 개발 기획", "건설로보틱스 시공자동화 기술 개발")]
    [InlineData("디지털 트윈 기반 관리 기술개발 기획 최종보고서", "디지털 트윈 기반 관리 기술개발")]
    [InlineData("(초)노후 도로시설물 철거·교체 의사결정지원 기술개발", "(초)노후 도로시설물 철거·교체 의사결정지원 기술개발")]
    public void 과제명_끝의_문서_종류를_뗀다(string title, string expected) =>
        Assert.Equal(expected, CardFillPrompts.CleanTitle(title));

    [Fact]
    public void 가운데에_가장_가까운_띄어쓰기에서_나눈다()
    {
        Assert.Equal("AI 기반 TBM 시공·안전\v통합관리 시스템 기술 개발", CardFillPrompts.BalancedBreak("AI 기반 TBM 시공·안전 통합관리 시스템 기술 개발"));
        Assert.Null(CardFillPrompts.BalancedBreak("띄어쓰기없음"));
    }
}
