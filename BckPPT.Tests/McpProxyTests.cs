using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json.Nodes;

namespace BckPPT.Tests;

/// <summary>
/// stdio 프록시(BckPptMcp.exe)를 진짜로 띄운다: 앱의 HTTP MCP로 넘기기, 앱이 없을 때의 답, 셀박스 트레이의 등록(--mcp-*).
/// 앱 주소·앱 위치·사용자 폴더·고정 경로는 환경 변수로 바꿔 이 PC의 BckPPT와 설정을 건드리지 않는다.
/// </summary>
public sealed class McpProxyTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "bckppt-proxy-" + Guid.NewGuid().ToString("N"));
    private static readonly string ProxyExe = Path.Combine(AppContext.BaseDirectory, "BckPptMcp.exe");

    public McpProxyTests()
    {
        Directory.CreateDirectory(Path.Combine(_root, "home", ".codex"));
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    private string Home => Path.Combine(_root, "home");
    private string ProxyDir => Path.Combine(_root, "proxy");
    private string FakeHost => Path.Combine(_root, "plugin", "BckPPT.exe");

    private ProcessStartInfo Start(string url, params string[] args)
    {
        var start = new ProcessStartInfo(ProxyExe)
        {
            UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
        };
        foreach (var a in args) start.ArgumentList.Add(a);
        start.Environment["BCKPPT_MCP_URL"] = url;
        start.Environment["BCKPPT_HOST_PATH"] = FakeHost;   // 없는 파일이면 앱을 못 찾은 것으로 본다
        start.Environment["BCKPPT_MCP_PROXY_DIR"] = ProxyDir;
        start.Environment["BCKSOFT_MCP_CONFIG_HOME"] = Home;
        return start;
    }

    /// <summary>줄들을 보내고 답 <paramref name="expected"/>줄을 읽은 뒤 stdin을 닫는다(클라이언트처럼 답을 받을 때까지 열어 둔다).</summary>
    private async Task<string[]> RunAsync(string url, int expected, params string[] lines)
    {
        using var process = Process.Start(Start(url))!;
        _ = process.StandardError.ReadToEndAsync();
        foreach (var line in lines) await process.StandardInput.WriteLineAsync(line);
        await process.StandardInput.FlushAsync();
        var replies = new List<string>();
        using var timeout = new CancellationTokenSource(30_000);
        while (replies.Count < expected && await process.StandardOutput.ReadLineAsync(timeout.Token) is { } reply)
            replies.Add(reply);
        process.StandardInput.Close();
        replies.AddRange((await process.StandardOutput.ReadToEndAsync()).Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
        Assert.True(process.WaitForExit(10_000), "stdin을 닫아도 프록시가 끝나지 않았다");
        return [.. replies];
    }

    private JsonObject Register(params string[] args)
    {
        using var process = Process.Start(Start("http://localhost:1/mcp", args))!;
        process.StandardInput.Close();
        var output = process.StandardOutput.ReadToEnd();
        Assert.True(process.WaitForExit(20_000));
        return (JsonObject)JsonNode.Parse(output)!;
    }

    private static int FreePort()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        int port = ((IPEndPoint)listener.LocalEndpoint).Port;
        listener.Stop();
        return port;
    }

    [Fact]
    public async Task Forwards_EachLineAsOnePost_AndDropsNotificationReplies()
    {
        int port = FreePort();
        using var server = new HttpListener();
        server.Prefixes.Add($"http://localhost:{port}/");
        server.Start();
        var bodies = new List<string>();
        _ = Task.Run(async () =>
        {
            while (server.IsListening)
            {
                HttpListenerContext ctx;
                try { ctx = await server.GetContextAsync(); } catch { break; }
                string body = await new StreamReader(ctx.Request.InputStream, Encoding.UTF8).ReadToEndAsync();
                lock (bodies) bodies.Add(body);
                var request = JsonNode.Parse(body)!;
                if (request["id"] is null) ctx.Response.StatusCode = 202;   // 알림
                else
                {
                    var bytes = Encoding.UTF8.GetBytes(new JsonObject
                    {
                        ["jsonrpc"] = "2.0", ["id"] = request["id"]!.DeepClone(),
                        ["result"] = new JsonObject { ["method"] = (string?)request["method"], ["글"] = "한국어" },
                    }.ToJsonString());
                    ctx.Response.ContentType = "application/json; charset=utf-8";
                    await ctx.Response.OutputStream.WriteAsync(bytes);
                }
                ctx.Response.Close();
            }
        });

        var replies = await RunAsync($"http://localhost:{port}/mcp", 2,
            """{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}""",
            """{"jsonrpc":"2.0","method":"notifications/initialized"}""",
            """{"jsonrpc":"2.0","id":"t","method":"tools/list"}""");

        Assert.Equal(3, bodies.Count);
        Assert.Equal(2, replies.Length);
        var byId = replies.Select(r => JsonNode.Parse(r)!).ToDictionary(r => r["id"]!.ToJsonString());
        Assert.Equal("initialize", (string?)byId["1"]["result"]!["method"]);
        Assert.Equal("tools/list", (string?)byId["\"t\""]["result"]!["method"]);
        Assert.Equal("한국어", (string?)byId["1"]["result"]!["글"]);
    }

    [Fact]
    public async Task HostMissing_AnswersRequestsWithAnError()
    {
        var replies = await RunAsync($"http://localhost:{FreePort()}/mcp", 1,
            """{"jsonrpc":"2.0","id":7,"method":"tools/list"}""",
            """{"jsonrpc":"2.0","method":"notifications/initialized"}""");

        var reply = JsonNode.Parse(Assert.Single(replies))!;
        Assert.Equal(7, (int)reply["id"]!);
        Assert.Contains("BckPPT.exe를 찾을 수 없습니다", (string?)reply["error"]!["message"]);
    }

    [Fact]
    public void Register_CopiesProxy_AndWritesStdioEntries()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FakeHost)!);
        File.WriteAllText(FakeHost, "fake");
        File.WriteAllText(Path.Combine(Home, ".claude.json"), """{ "theme": "dark", "mcpServers": { "bckppt": { "type": "http", "url": "http://localhost:38701/mcp" } } }""");
        string proxy = Path.Combine(ProxyDir, "BckPptMcp.exe");

        // 손으로 등록해 둔 http 주소는 다른 값으로 등록된 것으로 보인다 — 트레이가 덮어쓸지 묻는다.
        var before = Register("--mcp-status");
        Assert.Equal("stdio", (string?)before["transport"]);
        Assert.Equal(proxy, (string?)before["target"]);
        var claude = before["clients"]!.AsArray().Single(c => (string?)c!["id"] == "claude-code")!;
        Assert.True((bool)claude["registered"]!);
        Assert.False((bool)claude["current"]!);

        var after = Register("--mcp-register", "claude-code");
        Assert.True((bool)after["ok"]!);
        Assert.True(File.Exists(proxy));
        Assert.Equal(FakeHost, File.ReadAllText(Path.Combine(ProxyDir, "host.txt")));
        var config = JsonNode.Parse(File.ReadAllText(Path.Combine(Home, ".claude.json")))!;
        Assert.Equal("stdio", (string?)config["mcpServers"]!["bckppt"]!["type"]);
        Assert.Equal(proxy, (string?)config["mcpServers"]!["bckppt"]!["command"]);
        Assert.Equal("dark", (string?)config["theme"]);

        Register("--mcp-register", "codex");
        Assert.Contains($"command = '{proxy}'", File.ReadAllText(Path.Combine(Home, ".codex", "config.toml")));

        Register("--mcp-unregister", "claude-code");
        Assert.True(File.Exists(proxy));   // Codex가 아직 쓴다
        Register("--mcp-unregister", "codex");
        Assert.False(File.Exists(proxy));
        Assert.DoesNotContain("bckppt", File.ReadAllText(Path.Combine(Home, ".codex", "config.toml")));
    }
}
