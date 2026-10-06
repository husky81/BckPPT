using System.Text.Json;

namespace BckPPT.Llm;

/// <summary>
/// 대화 메시지. 도구 루프에서는 assistant 메시지가 <see cref="ToolCalls"/>를,
/// role "tool" 메시지가 <see cref="ToolCallId"/>를 가진다.
/// <see cref="ImageDataUrls"/>는 "data:image/png;base64,..." 형식이며 비전 모델에만 보낸다.
/// </summary>
public sealed record LlmMessage(string Role, string Content)
{
    public IReadOnlyList<LlmToolCall>? ToolCalls { get; init; }
    public string? ToolCallId { get; init; }
    public IReadOnlyList<string>? ImageDataUrls { get; init; }

    public static LlmMessage System(string content) => new("system", content);
    public static LlmMessage User(string content) => new("user", content);
    public static LlmMessage Assistant(string content, IReadOnlyList<LlmToolCall>? toolCalls = null) =>
        new("assistant", content) { ToolCalls = toolCalls is { Count: > 0 } ? toolCalls : null };
    public static LlmMessage Tool(string toolCallId, string content) => new("tool", content) { ToolCallId = toolCallId };
}

public sealed record LlmRequest(
    IReadOnlyList<LlmMessage> Messages,
    string? Model = null,
    int MaxTokens = 4096)
{
    public IReadOnlyList<LlmToolDefinition>? Tools { get; init; }

    /// <summary>"auto", "none", "required". null이면 서버 기본값.</summary>
    public string? ToolChoice { get; init; }
}

/// <summary>OpenAI 호환 function 도구 정의. <see cref="Parameters"/>는 JSON Schema 객체.</summary>
public sealed record LlmToolDefinition(string Name, string Description, JsonElement Parameters);

/// <summary>모델이 요청한 도구 호출. <see cref="ArgumentsJson"/>은 모델이 낸 문자열 그대로다.</summary>
public sealed record LlmToolCall(string Id, string Name, string ArgumentsJson);

public sealed record LlmUsage(int PromptTokens, int CompletionTokens)
{
    public int TotalTokens => PromptTokens + CompletionTokens;
    public static LlmUsage operator +(LlmUsage a, LlmUsage b) =>
        new(a.PromptTokens + b.PromptTokens, a.CompletionTokens + b.CompletionTokens);
    public static readonly LlmUsage Zero = new(0, 0);
}

/// <summary>비스트리밍 응답 한 건.</summary>
public sealed record LlmCompletion(
    string Content,
    IReadOnlyList<LlmToolCall> ToolCalls,
    string? FinishReason,
    LlmUsage Usage,
    string? RequestId = null,
    string? Model = null);

public enum LlmStreamEventKind { Started, ContentDelta, Completed, Error }

public sealed record LlmStreamEvent(
    LlmStreamEventKind Kind,
    string? Content = null,
    string? RequestId = null,
    string? Route = null,
    string? Model = null,
    JsonElement? Usage = null);

public sealed class LlmException(string message, string? errorCode = null, int? statusCode = null, Exception? inner = null)
    : Exception(message, inner)
{
    public string? ErrorCode { get; } = errorCode;
    public int? StatusCode { get; } = statusCode;
}
