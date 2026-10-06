using System.Text.Json;
using BckPPT.Agent;
using BckPPT.Llm;

namespace BckPPT.Tests;

/// <summary>2026-10-06 gitoky 실측 응답 형식으로 요청 본문과 응답 해석을 확인한다.</summary>
public class GitokyLlmProtocolTests
{
    private static readonly JsonElement EchoSchema =
        JsonDocument.Parse("""{"type":"object","properties":{"text":{"type":"string"}},"required":["text"]}""").RootElement;

    [Fact]
    public void 도구_정의와_도구_결과를_OpenAI_형식으로_보낸다()
    {
        var request = new LlmRequest(
        [
            LlmMessage.User("echo 해줘"),
            LlmMessage.Assistant("", [new LlmToolCall("tooluse_1", "echo", """{"text":"안녕"}""")]),
            LlmMessage.Tool("tooluse_1", "ECHO: 안녕"),
        ], "agent", 512)
        {
            Tools = [new LlmToolDefinition("echo", "그대로 돌려준다", EchoSchema)],
            ToolChoice = "auto"
        };

        var payload = GitokyLlmClient.BuildPayload(request, "auto", stream: false);

        Assert.Equal("agent", (string?)payload["model"]);
        Assert.False((bool)payload["stream"]!);
        Assert.Null(payload["stream_options"]);
        Assert.Equal("echo", (string?)payload["tools"]![0]!["function"]!["name"]);
        Assert.Equal("object", (string?)payload["tools"]![0]!["function"]!["parameters"]!["type"]);
        var assistant = payload["messages"]![1]!;
        Assert.Equal("""{"text":"안녕"}""", (string?)assistant["tool_calls"]![0]!["function"]!["arguments"]);
        Assert.Equal("tooluse_1", (string?)payload["messages"]![2]!["tool_call_id"]);
    }

    [Fact]
    public void 병렬_도구_호출은_호출과_결과를_번갈아_풀어_보낸다()
    {
        var request = new LlmRequest(
        [
            LlmMessage.User("두 쪽 복제"),
            LlmMessage.Assistant("복제합니다", [new LlmToolCall("a", "duplicate_slide", "{}"), new LlmToolCall("b", "duplicate_slide", "{}")]),
            LlmMessage.Tool("a", "35쪽"),
            LlmMessage.Tool("b", "36쪽"),
        ]);

        var messages = GitokyLlmClient.BuildPayload(request, "agent", stream: false)["messages"]!.AsArray();

        Assert.Equal(["user", "assistant", "tool", "assistant", "tool"], messages.Select(m => (string)m!["role"]!));
        Assert.Equal("a", (string?)messages[1]!["tool_calls"]![0]!["id"]);
        Assert.Equal("a", (string?)messages[2]!["tool_call_id"]);
        Assert.Equal("b", (string?)messages[3]!["tool_calls"]![0]!["id"]);
        Assert.Equal("", (string?)messages[3]!["content"]);
        Assert.Equal("36쪽", (string?)messages[4]!["content"]);
    }

    [Fact]
    public void 도구가_없으면_tools를_넣지_않는다()
    {
        var payload = GitokyLlmClient.BuildPayload(new LlmRequest([LlmMessage.User("hi")]), "auto", stream: true);
        Assert.Null(payload["tools"]);
        Assert.Null(payload["tool_choice"]);
        Assert.Equal("auto", (string?)payload["model"]);
        Assert.NotNull(payload["stream_options"]);
    }

    [Fact]
    public void 이미지는_content_파트로_보낸다()
    {
        var message = LlmMessage.User("무엇이 보이나") with { ImageDataUrls = ["data:image/png;base64,AAAA"] };
        var payload = GitokyLlmClient.BuildPayload(new LlmRequest([message], "vision"), "auto", stream: false);
        var parts = payload["messages"]![0]!["content"]!.AsArray();
        Assert.Equal("text", (string?)parts[0]!["type"]);
        Assert.Equal("data:image/png;base64,AAAA", (string?)parts[1]!["image_url"]!["url"]);
    }

    [Fact]
    public void 도구_호출_응답을_해석한다()
    {
        const string body = """
            {"id":"chatcmpl-b2eb","object":"chat.completion","model":"agent","choices":[{"index":0,"message":{"role":"assistant","content":"","tool_calls":[{"id":"tooluse_JaMh","type":"function","index":0,"function":{"name":"echo","arguments":"{\"text\":\"안녕 BckPPT\"}"}}]},"finish_reason":"tool_calls"}],"usage":{"prompt_tokens":190,"completion_tokens":24,"total_tokens":214}}
            """;
        var completion = GitokyLlmClient.ParseCompletion(body);

        Assert.Equal("", completion.Content);
        Assert.Equal("tool_calls", completion.FinishReason);
        var call = Assert.Single(completion.ToolCalls);
        Assert.Equal(("tooluse_JaMh", "echo"), (call.Id, call.Name));
        Assert.Equal("""{"text":"안녕 BckPPT"}""", call.ArgumentsJson);
        Assert.Equal(new LlmUsage(190, 24), completion.Usage);
        Assert.Equal("chatcmpl-b2eb", completion.RequestId);
    }

    [Fact]
    public void 오류_본문은_예외로_바꾼다()
    {
        var ex = Assert.Throws<LlmException>(() => GitokyLlmClient.ParseCompletion(
            """{"error":{"message":"model is not available for this request","type":"forbidden","code":"model_not_allowed"}}"""));
        Assert.Equal("forbidden", ex.ErrorCode);
    }

    /// <summary>
    /// 실제 gitoky 호출. GITOKY_LLM_API_KEY가 있을 때만 의미가 있고, 없으면 바로 통과한다.
    /// Phase A 완료 기준: 모델이 echo를 부르고 결과를 받아 답한다.
    /// </summary>
    [Fact]
    public async Task 실서버_echo_도구_왕복()
    {
        if (Environment.GetEnvironmentVariable("GITOKY_LLM_API_KEY") is not { Length: > 0 }) return;
        var options = GitokyLlmOptions.FromEnvironment(AppContext.BaseDirectory);
        using var client = new GitokyLlmClient(options);
        var tools = new[] { new LlmToolDefinition("echo", "입력 문자열을 그대로 돌려준다", EchoSchema) };
        var messages = new List<LlmMessage>
        {
            LlmMessage.System("도구를 써서 답하라."),
            LlmMessage.User("echo 도구로 '안녕 BckPPT'를 보내고 결과를 그대로 알려줘.")
        };

        var first = await client.CompleteAsync(new LlmRequest(messages, options.AgentModel, 512) { Tools = tools, ToolChoice = "auto" });
        var call = Assert.Single(first.ToolCalls);
        Assert.Equal("echo", call.Name);
        var text = JsonDocument.Parse(call.ArgumentsJson).RootElement.GetProperty("text").GetString();

        messages.Add(LlmMessage.Assistant(first.Content, first.ToolCalls));
        messages.Add(LlmMessage.Tool(call.Id, $"ECHO: {text}"));
        var second = await client.CompleteAsync(new LlmRequest(messages, options.AgentModel, 512) { Tools = tools });

        Assert.Empty(second.ToolCalls);
        Assert.Contains("BckPPT", second.Content);
    }
}
