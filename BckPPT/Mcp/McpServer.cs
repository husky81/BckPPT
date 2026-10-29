using System.IO;
using System.Net;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace BckPPT.Mcp;

/// <summary>
/// localhost MCP 서버(Streamable HTTP, JSON 응답). Claude Code 등에서
/// <c>claude mcp add --transport http bckppt http://localhost:38701/mcp</c> 로 연결한다.
/// 세션을 두지 않는다 — 요청마다 독립이고 상태는 PowerPoint 문서에 있다.
/// 브라우저 페이지가 localhost로 부르는 것(DNS rebinding)을 막으려고 localhost 밖의 Origin은 거절한다.
/// </summary>
public sealed class McpServer(McpToolCatalog catalog) : IDisposable
{
    public const int Port = 38701;
    public static string Endpoint => $"http://localhost:{Port}/mcp";
    public static string ClaudeAddCommand => $"claude mcp add --transport http bckppt {Endpoint}";

    private static readonly string[] SupportedVersions = ["2025-06-18", "2025-03-26", "2024-11-05"];
    private static readonly JsonSerializerOptions Json = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    private readonly HttpListener _listener = new();
    private readonly CancellationTokenSource _cts = new();

    /// <summary>시작하지 못한 이유(포트 사용 중 등). 시작했으면 null.</summary>
    public string? StartError { get; private set; }

    public void Start()
    {
        try
        {
            _listener.Prefixes.Add($"http://localhost:{Port}/");
            _listener.Start();
            _ = AcceptLoopAsync();
        }
        catch (Exception ex) { StartError = ex.Message; }
    }

    private async Task AcceptLoopAsync()
    {
        while (_listener.IsListening)
        {
            HttpListenerContext ctx;
            try { ctx = await _listener.GetContextAsync(); }
            catch { break; }
            _ = Task.Run(() => HandleAsync(ctx));
        }
    }

    private async Task HandleAsync(HttpListenerContext ctx)
    {
        var response = ctx.Response;
        try
        {
            if (!IsAllowedOrigin(ctx.Request.Headers["Origin"])) { response.StatusCode = 403; return; }
            if (ctx.Request.Url?.AbsolutePath.TrimEnd('/') != "/mcp") { response.StatusCode = 404; return; }
            // 서버가 먼저 보내는 알림이 없으므로 SSE 스트림(GET)은 열지 않는다.
            if (ctx.Request.HttpMethod != "POST") { response.StatusCode = 405; response.AddHeader("Allow", "POST"); return; }

            string body;
            using (var reader = new StreamReader(ctx.Request.InputStream, Encoding.UTF8)) body = await reader.ReadToEndAsync();
            var reply = await HandleBodyAsync(body, _cts.Token);
            if (reply is null) { response.StatusCode = 202; return; }   // 알림만 왔다

            var bytes = Encoding.UTF8.GetBytes(reply.ToJsonString(Json));
            response.StatusCode = 200;
            response.ContentType = "application/json; charset=utf-8";
            response.ContentLength64 = bytes.Length;
            await response.OutputStream.WriteAsync(bytes, _cts.Token);
        }
        catch { try { response.StatusCode = 500; } catch { } }
        finally { try { response.Close(); } catch { } }
    }

    /// <summary>JSON-RPC 본문 하나(단건 또는 배치)를 처리한다. 돌려줄 응답이 없으면(알림뿐) null.</summary>
    public async Task<JsonNode?> HandleBodyAsync(string body, CancellationToken cancellationToken)
    {
        JsonNode? parsed;
        try { parsed = JsonNode.Parse(body); }
        catch (JsonException) { return ErrorResponse(null, -32700, "JSON을 읽을 수 없습니다."); }

        if (parsed is JsonArray batch)
        {
            var replies = new JsonArray();
            foreach (var item in batch)
                if (await HandleMessageAsync(item, cancellationToken) is { } r) replies.Add(r);
            return replies.Count == 0 ? null : replies;
        }
        return await HandleMessageAsync(parsed, cancellationToken);
    }

    private async Task<JsonNode?> HandleMessageAsync(JsonNode? message, CancellationToken cancellationToken)
    {
        if (message is not JsonObject request || request["method"]?.GetValue<string>() is not { } method)
            return message is JsonObject { } maybeResponse && maybeResponse.ContainsKey("result") ? null
                : ErrorResponse(null, -32600, "JSON-RPC 요청이 아닙니다.");
        var id = request["id"]?.DeepClone();
        if (id is null) return null;   // 알림(notifications/initialized 등)은 답하지 않는다
        var parameters = request["params"] as JsonObject;

        try
        {
            JsonNode result = method switch
            {
                "initialize" => Initialize(parameters),
                "ping" => new JsonObject(),
                "tools/list" => ListTools(),
                "tools/call" => await CallToolAsync(parameters, cancellationToken),
                _ => throw new McpMethodNotFoundException(method),
            };
            return new JsonObject { ["jsonrpc"] = "2.0", ["id"] = id, ["result"] = result };
        }
        catch (McpMethodNotFoundException) { return ErrorResponse(id, -32601, $"지원하지 않는 메서드: {method}"); }
        catch (Exception ex) { return ErrorResponse(id, -32603, ex.Message); }
    }

    private static JsonObject Initialize(JsonObject? parameters)
    {
        var requested = parameters?["protocolVersion"]?.GetValue<string>();
        var version = requested is not null && SupportedVersions.Contains(requested) ? requested : SupportedVersions[0];
        return new JsonObject
        {
            ["protocolVersion"] = version,
            ["capabilities"] = new JsonObject { ["tools"] = new JsonObject() },
            ["serverInfo"] = new JsonObject
            {
                ["name"] = "bckppt",
                ["version"] = typeof(McpServer).Assembly.GetName().Version?.ToString(3) ?? "0.0.0",
            },
            ["instructions"] =
                "PowerPoint 데스크톱 앱의 문서를 읽고 고친다. 쪽 번호는 1부터, 크기는 pt. " +
                "file을 비우면 사용자가 PowerPoint에서 보고 있는 문서다. 열린 창의 문서는 그 자리에서 고치며 저장은 사용자가 한다. " +
                "글을 바꾼 뒤에는 결과의 줄 측정(line_breaks)이나 render_slide로 넘침을 확인한다.",
        };
    }

    private JsonObject ListTools()
    {
        var tools = new JsonArray();
        foreach (var tool in catalog.Tools.OrderBy(t => t.Name, StringComparer.Ordinal))
            tools.Add(new JsonObject
            {
                ["name"] = tool.Name,
                ["description"] = tool.Description,
                ["inputSchema"] = tool.InputSchema.DeepClone(),
            });
        return new JsonObject { ["tools"] = tools };
    }

    private async Task<JsonObject> CallToolAsync(JsonObject? parameters, CancellationToken cancellationToken)
    {
        var name = parameters?["name"]?.GetValue<string>() ?? throw new ArgumentException("도구 이름(name)이 없습니다.");
        using var args = JsonDocument.Parse(parameters["arguments"]?.ToJsonString() ?? "{}");
        var result = await catalog.CallAsync(name, args.RootElement, cancellationToken);
        var content = new JsonArray();
        foreach (var c in result.Content)
            content.Add(c.Type == "image"
                ? new JsonObject { ["type"] = "image", ["data"] = c.Data, ["mimeType"] = c.MimeType }
                : new JsonObject { ["type"] = "text", ["text"] = c.Text });
        return new JsonObject { ["content"] = content, ["isError"] = result.IsError };
    }

    private static JsonObject ErrorResponse(JsonNode? id, int code, string message) => new()
    {
        ["jsonrpc"] = "2.0",
        ["id"] = id,
        ["error"] = new JsonObject { ["code"] = code, ["message"] = message },
    };

    /// <summary>Origin이 없으면(CLI 클라이언트) 통과, 있으면 localhost만.</summary>
    public static bool IsAllowedOrigin(string? origin)
    {
        if (string.IsNullOrEmpty(origin)) return true;
        return Uri.TryCreate(origin, UriKind.Absolute, out var uri) && uri.IsLoopback;
    }

    public void Dispose()
    {
        _cts.Cancel();
        try { _listener.Stop(); _listener.Close(); } catch { }
        _cts.Dispose();
    }

    private sealed class McpMethodNotFoundException(string method) : Exception(method);
}
