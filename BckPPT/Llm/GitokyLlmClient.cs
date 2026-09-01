using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Runtime.CompilerServices;
using System.Text.Json;
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
        using var response = await SendWithRetryAsync(request, timeout.Token, cancellationToken);
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

    private async Task<HttpResponseMessage> SendWithRetryAsync(
        LlmRequest request, CancellationToken timeoutToken, CancellationToken cancellationToken)
    {
        for (var attempt = 0; ; attempt++)
        {
            var payload = new
            {
                model = request.Model ?? _options.Model,
                messages = request.Messages.Select(m => new { role = m.Role, content = m.Content }),
                max_tokens = request.MaxTokens,
                stream = true,
                stream_options = new { include_usage = true }
            };
            using var message = new HttpRequestMessage(HttpMethod.Post, "v1/chat/completions")
            {
                Content = JsonContent.Create(payload)
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
