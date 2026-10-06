using System.IO;
using BckPPT.Agent;
using BckPPT.Agent.Tools;
using BckPPT.Llm;

namespace BckPPT.Tests;

public sealed class AgentLoopTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "bckppt-agent-tests-" + Guid.NewGuid().ToString("N")[..8]);
    private readonly string _source;
    private readonly FakePptComService _ppt = new();

    public AgentLoopTests()
    {
        Directory.CreateDirectory(_dir);
        _source = Path.Combine(_dir, "착수보고.pptx");
        File.WriteAllText(_source, "원본");
        // 1쪽: 카드 템플릿. 본문 상자는 한 줄에 10자.
        _ppt.Slides.Add(
        [
            new FakeShape("C1_title", 0, 0, 200, 40, 20, "과제명"),
            new FakeShape("C1_rnd", 0, 50, 200, 100, 10, "기존 내용 1", "기존 내용 2"),
        ]);
        _ppt.Slides.Add([new FakeShape("Footer", 0, 0, 200, 20, 50, "끝")]);
        _ppt.SnapshotOriginal(_source);
    }

    private AgentSession StartSession() =>
        AgentSession.Start(_ppt, _source, AgentPrompts.System,
            workDirectory: Path.Combine(_dir, "work"), logDirectory: Path.Combine(_dir, "logs"));

    private static AgentLoop Loop(ILlmClient llm, AgentLoopOptions? options = null) =>
        new(llm, AgentFactory.CreateRegistry(includeSources: false), options ?? new AgentLoopOptions());

    [Fact]
    public async Task 복제하고_문단을_바꾼_뒤_끝낸다()
    {
        var llm = new FakeLlmClient(
            FakeLlmClient.Calls(("duplicate_slide", """{"source": 1, "after": 1}""")),
            FakeLlmClient.Calls(("set_paragraphs", """{"slide": 2, "shape": "C1_rnd", "paragraphs": ["새 내용 하나", "둘"]}""")),
            FakeLlmClient.Calls(("finish", """{"summary": "2쪽에 카드를 채웠습니다."}""")));
        using var session = StartSession();

        var result = await Loop(llm).RunAsync(session, "1쪽 카드로 새 쪽을 만들어 채워줘");

        Assert.Equal(AgentStopReason.Finished, result.Reason);
        Assert.Equal("2쪽에 카드를 채웠습니다.", result.Summary);
        Assert.Equal(3, result.Turns);
        Assert.Empty(result.RemainingIssues);
        Assert.Equal(3, _ppt.Slides.Count);
        Assert.Equal(["새 내용 하나", "둘"], _ppt.Slides[1][1].Paragraphs);
        Assert.Equal(["기존 내용 1", "기존 내용 2"], _ppt.Slides[0][1].Paragraphs);   // 템플릿은 그대로
        Assert.Equal(1, session.Context.Origins.BaselineOf(2));

        // 도구 결과가 대화에 붙고, 문단 교체 결과에는 줄 측정이 들어 있다.
        var toolMessages = session.Messages.Where(m => m.Role == "tool").ToList();
        Assert.Equal(3, toolMessages.Count);
        Assert.Contains("C1_rnd: 2줄 (기준 2줄 이내)", toolMessages[1].Content);
        Assert.Equal(new LlmUsage(300, 60), result.Usage);
        Assert.All(llm.Requests, r => Assert.NotNull(r.Tools));
    }

    [Fact]
    public async Task 작업은_사본에만_하고_원본_파일은_그대로다()
    {
        var llm = new FakeLlmClient(FakeLlmClient.Text("무엇을 할까요?"));
        using var session = StartSession();
        await Loop(llm).RunAsync(session, "안녕");

        Assert.NotEqual(_source, session.WorkingPath);
        Assert.True(File.Exists(session.WorkingPath));
        var result = session.Approve();
        Assert.StartsWith(Path.Combine(_dir, "착수보고_BckPPT_"), result);
        Assert.Equal("원본", File.ReadAllText(_source));
        session.Discard();
        Assert.False(File.Exists(session.WorkingPath));
        Assert.True(File.Exists(result));
    }

    [Fact]
    public async Task 위반이_남은_채_끝내려_하면_돌려보낸다()
    {
        // 한 줄 10자 상자에 꼬리가 넘어가는 문구 → 고아 단어. 고친 뒤에야 끝난다.
        var llm = new FakeLlmClient(
            FakeLlmClient.Calls(("set_paragraphs", """{"slide": 1, "shape": "C1_rnd", "paragraphs": ["품질 예측 기술 고도화 개발"]}""")),
            FakeLlmClient.Calls(("finish", """{"summary": "끝"}""")),
            FakeLlmClient.Calls(("set_paragraphs", """{"slide": 1, "shape": "C1_rnd", "paragraphs": ["예측 기술 개발"]}""")),
            FakeLlmClient.Calls(("finish", """{"summary": "고쳤습니다"}""")));
        using var session = StartSession();

        var result = await Loop(llm).RunAsync(session, "채워줘");

        Assert.Equal(AgentStopReason.Finished, result.Reason);
        Assert.Equal("고쳤습니다", result.Summary);
        var rejection = session.Messages.Where(m => m.Role == "tool").ElementAt(1).Content;
        Assert.Contains("아직 끝낼 수 없습니다", rejection);
        Assert.Contains("만 남음", rejection);
        Assert.Empty(result.RemainingIssues);
    }

    [Fact]
    public async Task 되돌려_보낼_횟수를_넘기면_남은_위반과_함께_끝낸다()
    {
        var llm = new FakeLlmClient(
            FakeLlmClient.Calls(("set_paragraphs", """{"slide": 1, "shape": "C1_rnd", "paragraphs": ["품질 예측 기술 고도화 개발"]}""")),
            FakeLlmClient.Calls(("finish", """{"summary": "끝"}""")),
            FakeLlmClient.Calls(("finish", """{"summary": "그래도 끝"}""")));
        using var session = StartSession();

        var result = await Loop(llm, new AgentLoopOptions { LintRejectionsBeforeFinish = 1 }).RunAsync(session, "채워줘");

        Assert.Equal(AgentStopReason.Finished, result.Reason);
        Assert.Contains(result.RemainingIssues, i => i.Rule == BckPPT.PowerPoint.LintRule.OrphanWord);
    }

    [Fact]
    public async Task 잘못된_인자는_오류로_돌려주고_계속한다()
    {
        var llm = new FakeLlmClient(
            FakeLlmClient.Calls(("read_text", """{"slide": 1, "shape": "없는 도형"}"""), ("no_such_tool", "{}")),
            FakeLlmClient.Text("도형을 찾지 못했습니다."));
        using var session = StartSession();
        var events = new List<AgentEvent>();

        var result = await Loop(llm).RunAsync(session, "읽어줘", new SyncProgress<AgentEvent>(events.Add));

        Assert.Equal(AgentStopReason.AnsweredWithoutTools, result.Reason);
        var errors = events.Where(e => e.Kind == AgentEventKind.ToolFinished && e.IsError).ToList();
        Assert.Equal(2, errors.Count);
        Assert.Contains("찾지 못했습니다", errors[0].Text);
        Assert.Contains("도구는 없습니다", errors[1].Text);
    }

    [Fact]
    public async Task 바꾼_뒤_말만_하고_멈추면_한_번_이어가게_한다()
    {
        var llm = new FakeLlmClient(
            FakeLlmClient.Calls(("set_paragraphs", """{"slide": 1, "shape": "C1_rnd", "paragraphs": ["새 내용"]}""")),
            FakeLlmClient.Text("이제 lint_slide로 확인하겠습니다."),
            FakeLlmClient.Calls(("finish", """{"summary": "끝"}""")));
        using var session = StartSession();

        var result = await Loop(llm).RunAsync(session, "채워줘");

        Assert.Equal(AgentStopReason.Finished, result.Reason);
        Assert.StartsWith(AgentLoop.ContinueNudge, llm.Requests[2].Messages[^1].Content);
    }

    [Fact]
    public async Task 이어가라고_한_뒤에도_말로_답하면_끝낸다()
    {
        var llm = new FakeLlmClient(
            FakeLlmClient.Calls(("set_paragraphs", """{"slide": 1, "shape": "C1_rnd", "paragraphs": ["새 내용"]}""")),
            FakeLlmClient.Text("어느 카드를 바꿀까요?"),
            FakeLlmClient.Text("어느 카드를 바꿀까요?"),
            FakeLlmClient.Text("어느 카드를 바꿀까요?"));
        using var session = StartSession();

        var result = await Loop(llm).RunAsync(session, "채워줘");

        Assert.Equal(AgentStopReason.AnsweredWithoutTools, result.Reason);
        Assert.Equal(4, result.Turns);   // 이어가라는 안내 2번 뒤 끝낸다
    }

    [Fact]
    public async Task 템플릿에_원래_있던_문제는_판정하지_않는다()
    {
        // 원본 1쪽 C1_title이 이미 고아 단어를 갖고 있다. 복제 후 본문만 바꾸면 제목은 판정 대상이 아니다.
        _ppt.Slides[0][0].Paragraphs[0] = "스마트 제조 공정 품질 예측 기술 개발";   // 20자 줄 → '개발' 꼬리
        _ppt.SnapshotOriginal(_source);
        var llm = new FakeLlmClient(
            FakeLlmClient.Calls(("duplicate_slide", """{"source": 1, "after": 1}""")),
            FakeLlmClient.Calls(("set_paragraphs", """{"slide": 2, "shape": "C1_rnd", "paragraphs": ["새 내용"]}""")),
            FakeLlmClient.Calls(("finish", """{"summary": "끝"}""")));
        using var session = StartSession();

        var result = await Loop(llm).RunAsync(session, "채워줘");

        Assert.Equal(3, result.Turns);   // 되돌려 보내지 않았다
        Assert.Empty(result.RemainingIssues);
    }

    [Fact]
    public async Task 문맥_한도를_넘으면_더_줄이고_다시_묻는다()
    {
        var llm = new FakeLlmClient(
            new LlmException("request exceeds the model's context window of 32768 tokens; this request needs at least 32769."),
            FakeLlmClient.Text("줄인 뒤 답"));
        using var session = StartSession();
        for (int i = 0; i < 6; i++)
        {
            session.Append(LlmMessage.Assistant("", [new LlmToolCall($"c{i}", "read_source", "{}")]));
            session.Append(LlmMessage.Tool($"c{i}", new string('가', 3000)));
        }

        var result = await Loop(llm).RunAsync(session, "계속");

        Assert.Equal("줄인 뒤 답", result.Summary);
        Assert.Equal(2, llm.Requests.Count);
        Assert.True(llm.Requests[1].Messages.Sum(m => m.Content.Length) < llm.Requests[0].Messages.Sum(m => m.Content.Length));
    }

    [Fact]
    public void 문맥_예산은_고정_비용과_글자당_토큰으로_추정한다()
    {
        using var session = StartSession();
        Assert.Null(session.EstimateCharBudget(26000));
        session.ObserveContext(2000, 6000);                 // 도구 정의 약 5,000토큰 + 대화 2,000자
        Assert.Equal(8666, session.EstimateCharBudget(26000));   // 아직은 보수적 비율(1/3)
        session.ObserveContext(12000, 16000);               // 1만 자 늘 때 1만 토큰 → 글자당 1토큰, 고정 4,000
        Assert.Equal(22000, session.EstimateCharBudget(26000));
    }

    [Fact]
    public async Task 묶음_도구로_이름과_문단을_한_번에_바꾼다()
    {
        var llm = new FakeLlmClient(
            FakeLlmClient.Calls(("rename_shapes", """{"slide": 1, "items": [{"shape": "C1_title", "new_name": "K_title"}, {"shape": "C1_rnd", "new_name": "K_rnd"}, {"shape": "없음", "new_name": "X"}]}""")),
            FakeLlmClient.Calls(("set_texts", """{"slide": 1, "items": [{"shape": "K_title", "paragraphs": ["새 제목"]}, {"shape": "K_rnd", "paragraphs": ["가", "나"]}]}""")),
            FakeLlmClient.Calls(("finish", """{"summary": "끝"}""")));
        using var session = StartSession();

        var result = await Loop(llm).RunAsync(session, "채워줘");

        Assert.Equal(AgentStopReason.Finished, result.Reason);
        Assert.Equal(["K_title", "K_rnd"], _ppt.Slides[0].Select(s => s.Name));
        Assert.Equal(["새 제목"], _ppt.Slides[0][0].Paragraphs);
        Assert.Equal(["가", "나"], _ppt.Slides[0][1].Paragraphs);
        var tools = session.Messages.Where(m => m.Role == "tool").Select(m => m.Content).ToList();
        Assert.Contains("2개 이름을 바꿨습니다", tools[0]);
        Assert.Contains("없음", tools[0]);                         // 실패는 적고 나머지는 계속
        Assert.Contains("K_title: 1줄", tools[1]);
        Assert.Contains("K_rnd: 2줄", tools[1]);
    }

    [Fact]
    public async Task 턴_상한에서_멈춘다()
    {
        var calls = Enumerable.Range(0, 3).Select(_ => FakeLlmClient.Calls(("list_shapes", """{"slide": 1}"""))).ToArray();
        using var session = StartSession();
        var result = await Loop(new FakeLlmClient(calls), new AgentLoopOptions { MaxTurns = 3 }).RunAsync(session, "봐줘");
        Assert.Equal(AgentStopReason.TurnLimit, result.Reason);
    }

    [Fact]
    public async Task 후속_요청은_같은_대화를_이어간다()
    {
        var llm = new FakeLlmClient(FakeLlmClient.Text("첫 답"), FakeLlmClient.Text("둘째 답"));
        using var session = StartSession();
        var loop = Loop(llm);
        await loop.RunAsync(session, "첫 요청");
        await loop.RunAsync(session, "1줄씩 줄여줘");

        var second = llm.Requests[1].Messages;
        Assert.Equal(["system", "user", "assistant", "user"], second.Select(m => m.Role));
        Assert.Equal("1줄씩 줄여줘", second[^1].Content);
    }

    [Fact]
    public void 예산을_넘으면_오래된_도구_결과부터_줄인다()
    {
        var messages = new List<LlmMessage> { LlmMessage.System("s"), LlmMessage.User("u") };
        for (int i = 0; i < 10; i++)
        {
            messages.Add(LlmMessage.Assistant("", [new LlmToolCall($"c{i}", "line_breaks", "{}")]));
            messages.Add(LlmMessage.Tool($"c{i}", new string('가', 1000)));
        }

        int elided = AgentLoop.Compact(messages, budget: 5000);

        Assert.True(elided > 0);
        Assert.True(messages.Sum(m => m.Content.Length) <= 5000);
        Assert.Equal(22, messages.Count);                                    // 메시지는 지우지 않는다
        Assert.StartsWith(AgentLoop.ElidedPrefix, messages[3].Content);
        Assert.Equal(1000, messages[^1].Content.Length);                     // 최근 결과는 남긴다
    }

    [Fact]
    public void 긴_도구_결과는_자른다()
    {
        var text = AgentLoop.Truncate(new string('x', 9000), 8000);
        Assert.StartsWith(new string('x', 8000), text);
        Assert.Contains("1000자 생략", text);
    }

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); } catch { }
    }

    /// <summary>Progress&lt;T&gt;는 동기화 컨텍스트로 넘겨 비동기로 부르므로 테스트에서는 바로 부른다.</summary>
    private sealed class SyncProgress<T>(Action<T> report) : IProgress<T>
    {
        public void Report(T value) => report(value);
    }
}

public class ToolArgumentTests
{
    [Theory]
    [InlineData("216*4", 864f)]
    [InlineData("32 + 236", 268f)]
    [InlineData("(896 - 20) / 4", 219f)]
    public void 계산식_인자를_계산한다(string expression, float expected)
    {
        using var doc = System.Text.Json.JsonDocument.Parse($$"""{"width": "{{expression}}"}""");
        Assert.Equal(expected, BckPPT.Agent.Tools.ToolArgs.OptionalFloat(doc.RootElement, "width"));
    }

    [Fact]
    public void 숫자와_연산자_외의_글자는_거부한다()
    {
        using var doc = System.Text.Json.JsonDocument.Parse("""{"width": "Width*2"}""");
        Assert.Throws<BckPPT.Agent.Tools.AgentToolException>(() => BckPPT.Agent.Tools.ToolArgs.OptionalFloat(doc.RootElement, "width"));
    }

    [Fact]
    public void 카드_영역_안의_도형을_고른다()
    {
        BckPPT.Models.ShapeEntry E(string a, float l, float t, float w, float h, int id) => new(a, a, "", 1, l, t, w, h, false, "", id);
        var card = E("@20", 262, 112, 216, 390, 20);
        var shapes = new[]
        {
            card,
            E("@24", 271, 117, 27, 27, 24),       // 카드 안 번호
            E("@27", 272, 144, 206, 35, 27),      // 카드 안 제목
            E("@33", 482, 112, 216, 390, 33),     // 옆 카드
            E("Panel", 32, 93, 896, 422, 1),      // 바깥 틀(더 크다)
        };
        var inside = BckPPT.Agent.Tools.PptAgentTools.InsideOf(card, shapes);
        Assert.Equal(["@24", "@27"], inside.Select(s => s.Address));
    }
}
