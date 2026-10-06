using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.IO;

namespace BckPPT.Llm;

public sealed class GitokyLlmClient : ILlmClient, IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly GitokyLlmOptions _options;

    public GitokyLlmClient(GitokyLlmOptions options, HttpMessageHandler? handler = null)
    {
        _options = options;
        _httpClient = handler is null ? new HttpClient() : new HttpClient(handler);
        _httpClient.BaseAddress = options.BaseUri;
        _httpClient.Timeout = Timeout.InfiniteTimeSpan;
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", options.ApiKey);
        _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("BckPPT/2.0");
    }

    public async Task TestConnectionAsync(CancellationToken cancellationToken = default)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(15));
        using var response = await _httpClient.GetAsync("v1/models", timeout.Token);
        if (!response.IsSuccessStatusCode) throw await CreateExceptionAsync(response, timeout.Token);
    }

    public async IAsyncEnumerable<LlmStreamEvent> StreamAsync(
        LlmRequest request,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(_options.Timeout);
        using var response = await SendWithRetryAsync(request, stream: true, timeout.Token, cancellationToken);
        var requestId = Header(response, "X-LLM-Request-ID");
        yield return new LlmStreamEvent(LlmStreamEventKind.Started,
            RequestId: requestId, Route: Header(response, "X-LLM-Route"), Model: Header(response, "X-LLM-Model"));

        await using var stream = await response.Content.ReadAsStreamAsync(timeout.Token);
        using var reader = new StreamReader(stream);
        while (await reader.ReadLineAsync(timeout.Token) is { } line)
        {
            if (!line.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) continue;
            var data = line[5..].Trim();
            if (data.Length == 0) continue;
            if (data == "[DONE]")
            {
                yield return new LlmStreamEvent(LlmStreamEventKind.Completed, RequestId: requestId);
                yield break;
            }

            using var document = JsonDocument.Parse(data);
            var root = document.RootElement;
            if (root.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.Object)
                throw ParseStreamException(error, requestId);
            if (root.TryGetProperty("usage", out var usage) && usage.ValueKind == JsonValueKind.Object)
                yield return new LlmStreamEvent(LlmStreamEventKind.ContentDelta, Usage: usage.Clone(), RequestId: requestId);
            if (root.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0 &&
                choices[0].TryGetProperty("delta", out var delta) &&
                delta.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.String)
                yield return new LlmStreamEvent(LlmStreamEventKind.ContentDelta, content.GetString(), requestId);
        }
        throw new LlmException("LLM 스트림이 완료 신호 없이 종료되었습니다.", "incomplete_stream");
    }

    public async Task<LlmCompletion> CompleteAsync(LlmRequest request, CancellationToken cancellationToken = default)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(_options.Timeout);
        using var response = await SendWithRetryAsync(request, stream: false, timeout.Token, cancellationToken);
        string body;
        try { body = await response.Content.ReadAsStringAsync(timeout.Token); }
        catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested)
        {
            throw new LlmException("LLM 응답 시간이 초과되었습니다.", "timeout", inner: ex);
        }
        return ParseCompletion(body, Header(response, "X-LLM-Request-ID"), Header(response, "X-LLM-Model"));
    }

    /// <summary>OpenAI 호환 chat.completion 본문을 해석한다. 내용이 없고 도구 호출만 있는 응답도 정상이다.</summary>
    public static LlmCompletion ParseCompletion(string body, string? requestId = null, string? model = null)
    {
        using var document = JsonDocument.Parse(body);
        var root = document.RootElement;
        if (root.TryGetProperty("error", out var error) && error.ValueKind == JsonValueKind.Object)
            throw ParseStreamException(error, requestId);
        if (!root.TryGetProperty("choices", out var choices) || choices.GetArrayLength() == 0)
            throw new LlmException("LLM 응답에 choices가 없습니다.", "empty_response");

        var choice = choices[0];
        var message = choice.GetProperty("message");
        var content = message.TryGetProperty("content", out var c) && c.ValueKind == JsonValueKind.String
            ? c.GetString() ?? "" : "";
        var toolCalls = new List<LlmToolCall>();
        if (message.TryGetProperty("tool_calls", out var calls) && calls.ValueKind == JsonValueKind.Array)
        {
            foreach (var call in calls.EnumerateArray())
            {
                var function = call.GetProperty("function");
                var arguments = function.TryGetProperty("arguments", out var a)
                    ? a.ValueKind == JsonValueKind.String ? a.GetString() ?? "{}" : a.GetRawText()
                    : "{}";
                toolCalls.Add(new LlmToolCall(
                    call.TryGetProperty("id", out var id) ? id.GetString() ?? "" : "",
                    function.GetProperty("name").GetString() ?? "",
                    arguments));
            }
        }
        var finish = choice.TryGetProperty("finish_reason", out var f) && f.ValueKind == JsonValueKind.String
            ? f.GetString() : null;
        var usage = LlmUsage.Zero;
        if (root.TryGetProperty("usage", out var u) && u.ValueKind == JsonValueKind.Object)
            usage = new LlmUsage(
                u.TryGetProperty("prompt_tokens", out var p) ? p.GetInt32() : 0,
                u.TryGetProperty("completion_tokens", out var ct) ? ct.GetInt32() : 0);
        model ??= root.TryGetProperty("model", out var m) ? m.GetString() : null;
        requestId ??= root.TryGetProperty("id", out var rid) ? rid.GetString() : null;
        return new LlmCompletion(content, toolCalls, finish, usage, requestId, model);
    }

    /// <summary>요청 본문. 도구·도구 결과·이미지 파트는 있을 때만 넣는다.</summary>
    /// <summary>
    /// 한 assistant 메시지에 도구 호출이 여러 개면 "호출1, 결과1, 호출2, 결과2"로 풀어 보낸다.
    /// gitoky(Bedrock 변환)는 병렬 호출 뒤 도구 결과가 연달아 오는 요청을 upstream_rejected로 거부한다
    /// (2026-10-06 실측, 호출 하나씩 풀면 통과). 의미는 같고 형식만 바뀐다.
    /// </summary>
    public static IReadOnlyList<LlmMessage> InterleaveParallelToolCalls(IReadOnlyList<LlmMessage> messages)
    {
        var result = new List<LlmMessage>(messages.Count);
        for (int i = 0; i < messages.Count; i++)
        {
            var m = messages[i];
            if (m.Role != "assistant" || m.ToolCalls is not { Count: > 1 } calls)
            {
                result.Add(m);
                continue;
            }
            // 바로 뒤에 이어지는 도구 결과를 호출 Id로 짝짓는다.
            var answers = new Dictionary<string, LlmMessage>();
            int j = i + 1;
            for (; j < messages.Count && messages[j].Role == "tool"; j++)
                if (messages[j].ToolCallId is { } id) answers[id] = messages[j];
            if (calls.Any(c => !answers.ContainsKey(c.Id)))
            {
                result.Add(m);   // 결과가 아직 없는(마지막) 호출은 그대로 둔다
                continue;
            }
            for (int k = 0; k < calls.Count; k++)
            {
                result.Add(m with { Content = k == 0 ? m.Content : "", ToolCalls = [calls[k]] });
                result.Add(answers[calls[k].Id]);
            }
            i = j - 1;
        }
        return result;
    }

    public static JsonObject BuildPayload(LlmRequest request, string defaultModel, bool stream)
    {
        var messages = new JsonArray();
        foreach (var m in InterleaveParallelToolCalls(request.Messages))
        {
            var node = new JsonObject { ["role"] = m.Role };
            if (m.ImageDataUrls is { Count: > 0 } images)
            {
                var parts = new JsonArray { new JsonObject { ["type"] = "text", ["text"] = m.Content } };
                foreach (var url in images)
                    parts.Add(new JsonObject { ["type"] = "image_url", ["image_url"] = new JsonObject { ["url"] = url } });
                node["content"] = parts;
            }
            else
            {
                node["content"] = m.Content;
            }
            if (m.ToolCalls is { Count: > 0 } calls)
            {
                node["tool_calls"] = new JsonArray(calls.Select(call => (JsonNode)new JsonObject
                {
                    ["id"] = call.Id,
                    ["type"] = "function",
                    ["function"] = new JsonObject { ["name"] = call.Name, ["arguments"] = call.ArgumentsJson }
                }).ToArray());
            }
            if (m.ToolCallId is not null) node["tool_call_id"] = m.ToolCallId;
            messages.Add(node);
        }

        var payload = new JsonObject
        {
            ["model"] = request.Model ?? defaultModel,
            ["messages"] = messages,
            ["max_tokens"] = request.MaxTokens,
            ["stream"] = stream
        };
        if (stream) payload["stream_options"] = new JsonObject { ["include_usage"] = true };
        if (request.Tools is { Count: > 0 } tools)
        {
            payload["tools"] = new JsonArray(tools.Select(tool => (JsonNode)new JsonObject
            {
                ["type"] = "function",
                ["function"] = new JsonObject
                {
                    ["name"] = tool.Name,
                    ["description"] = tool.Description,
                    ["parameters"] = JsonNode.Parse(tool.Parameters.GetRawText())
                }
            }).ToArray());
            if (request.ToolChoice is not null) payload["tool_choice"] = request.ToolChoice;
        }
        return payload;
    }

    private async Task<HttpResponseMessage> SendWithRetryAsync(
        LlmRequest request, bool stream, CancellationToken timeoutToken, CancellationToken cancellationToken)
    {
        var json = BuildPayload(request, _options.Model, stream).ToJsonString();
        for (var attempt = 0; ; attempt++)
        {
            using var message = new HttpRequestMessage(HttpMethod.Post, "v1/chat/completions")
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
            try
            {
                var response = await _httpClient.SendAsync(message, HttpCompletionOption.ResponseHeadersRead, timeoutToken);
                if (response.IsSuccessStatusCode) return response;
                var exception = await CreateExceptionAsync(response, timeoutToken);
                var transient = IsTransient(response.StatusCode);
                response.Dispose();
                if (!transient || attempt >= _options.MaxRetries) throw exception;
            }
            catch (HttpRequestException) when (attempt < _options.MaxRetries) { }
            catch (HttpRequestException ex)
            {
                // 게이트웨이가 응답 없이 연결을 끊는 경우(2026-10-06 실측: 지연이 15초를 넘으면 끊김).
                throw new LlmException("LLM 서버 연결이 끊겼습니다. 잠시 후 다시 시도해 주세요.", "connection_lost", inner: ex);
            }
            catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                throw new LlmException("LLM 응답 시간이 초과되었습니다.", "timeout", inner: ex);
            }
            await Task.Delay(RetryDelay(attempt), cancellationToken);
        }
    }

    private static string? Header(HttpResponseMessage response, string name) =>
        response.Headers.TryGetValues(name, out var values) ? values.FirstOrDefault() : null;
    private static bool IsTransient(HttpStatusCode status) => status is HttpStatusCode.RequestTimeout or HttpStatusCode.TooManyRequests or >= HttpStatusCode.InternalServerError;
    private static TimeSpan RetryDelay(int attempt) => TimeSpan.FromMilliseconds(500 * Math.Pow(2, attempt));

    private static async Task<LlmException> CreateExceptionAsync(HttpResponseMessage response, CancellationToken token)
    {
        var body = await response.Content.ReadAsStringAsync(token);
        string? code = null;
        string message = $"LLM 서버 오류 ({(int)response.StatusCode})";
        try
        {
            using var json = JsonDocument.Parse(body);
            var error = json.RootElement.TryGetProperty("error", out var nested) ? nested : json.RootElement;
            if (error.TryGetProperty("message", out var m)) message = m.GetString() ?? message;
            if (error.TryGetProperty("type", out var t)) code = t.GetString();
            if (error.TryGetProperty("code", out var c)) code ??= c.GetString();
        }
        catch (JsonException) { }
        return new LlmException(message, code, (int)response.StatusCode);
    }

    private static LlmException ParseStreamException(JsonElement error, string? requestId)
    {
        var message = error.TryGetProperty("message", out var m) ? m.GetString() : "LLM 스트림 오류";
        var code = error.TryGetProperty("type", out var t) ? t.GetString() : "stream_error";
        return new LlmException($"{message} (요청 ID: {requestId ?? "없음"})", code);
    }

    public void Dispose() => _httpClient.Dispose();
}
