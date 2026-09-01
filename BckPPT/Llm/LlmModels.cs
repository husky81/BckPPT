using System.Text.Json;

namespace BckPPT.Llm;

public sealed record LlmMessage(string Role, string Content);

public sealed record LlmRequest(
    IReadOnlyList<LlmMessage> Messages,
    string? Model = null,
    int MaxTokens = 4096);

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
