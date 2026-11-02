using System.Text.Json.Nodes;
using BckPPT.Mcp;

namespace BckPPT.Tests;

/// <summary>
/// 셀박스 트레이가 부르는 등록(http 방식). 가짜 사용자 폴더(BCKSOFT_MCP_CONFIG_HOME)에서만 읽고 쓰며 CLI는 쓰지 않는다.
/// 사용자 폴더는 프로세스 전체 환경 변수라 이 클래스 안에서만 바꾼다(xunit은 한 클래스의 시험을 차례로 돌린다).
/// </summary>
public sealed class McpClientRegistrationTests : IDisposable
{
    private readonly string _home = Path.Combine(Path.GetTempPath(), "bckppt-registration-" + Guid.NewGuid().ToString("N"));
    private readonly McpClientRegistration _registration = McpServer.ClientRegistration;

    public McpClientRegistrationTests()
    {
        Directory.CreateDirectory(Path.Combine(_home, ".codex"));
        Environment.SetEnvironmentVariable("BCKSOFT_MCP_CONFIG_HOME", _home);
    }

    public void Dispose()
    {
        Environment.SetEnvironmentVariable("BCKSOFT_MCP_CONFIG_HOME", null);
        try { Directory.Delete(_home, recursive: true); } catch { }
    }

    private string ClaudeJson => Path.Combine(_home, ".claude.json");
    private string CodexToml => Path.Combine(_home, ".codex", "config.toml");

    private JsonNode Client(string id) =>
        _registration.Status()["clients"]!.AsArray().Single(c => (string?)c!["id"] == id)!;

    [Fact]
    public void Register_WritesHttpEntries_AndKeepsOtherSettings()
    {
        File.WriteAllText(ClaudeJson, """{ "theme": "dark", "mcpServers": { "BckHWP": { "type": "stdio", "command": "x.exe" } } }""");
        File.WriteAllText(CodexToml, "model = \"gpt\"\n\n[mcp_servers.bckhwp]\ncommand = 'x.exe'\n");

        _registration.Register("claude-code");
        _registration.Register("codex");

        var claude = JsonNode.Parse(File.ReadAllText(ClaudeJson))!;
        Assert.Equal("http", (string?)claude["mcpServers"]!["bckppt"]!["type"]);
        Assert.Equal(McpServer.Endpoint, (string?)claude["mcpServers"]!["bckppt"]!["url"]);
        Assert.Equal("dark", (string?)claude["theme"]);
        Assert.NotNull(claude["mcpServers"]!["BckHWP"]);

        var toml = File.ReadAllText(CodexToml);
        Assert.Contains($"[mcp_servers.bckppt]", toml);
        Assert.Contains($"url = '{McpServer.Endpoint}'", toml);
        Assert.Contains("[mcp_servers.bckhwp]", toml);
        Assert.Contains("model = \"gpt\"", toml);

        Assert.True((bool)Client("claude-code")["current"]!);
        Assert.True((bool)Client("codex")["current"]!);
        Assert.Equal("http", (string?)_registration.Status()["transport"]);
    }

    [Fact]
    public void Unregister_RemovesOnlyBckPpt()
    {
        File.WriteAllText(CodexToml, "[mcp_servers.bckhwp]\ncommand = 'x.exe'\n");
        _registration.Register("codex");

        _registration.Unregister("codex");

        var toml = File.ReadAllText(CodexToml);
        Assert.DoesNotContain("bckppt", toml);
        Assert.Contains("[mcp_servers.bckhwp]", toml);
        Assert.False((bool)Client("codex")["registered"]!);
    }

    [Fact]
    public void OtherAddress_IsRegisteredButNotCurrent()
    {
        File.WriteAllText(ClaudeJson, """{ "mcpServers": { "bckppt": { "type": "http", "url": "http://localhost:9999/mcp" } } }""");

        var claude = Client("claude-code");

        Assert.True((bool)claude["registered"]!);
        Assert.False((bool)claude["current"]!);
        Assert.Equal("http://localhost:9999/mcp", (string?)claude["command"]);
    }

    [Fact]
    public void Refresh_HasNothingToDoForHttp()
    {
        _registration.Register("claude-code");
        var before = File.ReadAllText(ClaudeJson);

        _registration.Refresh();

        Assert.Equal(before, File.ReadAllText(ClaudeJson));
    }

    [Fact]
    public void NotRegistrationArguments_AreLeftToTheApp()
    {
        Assert.Null(_registration.TryRun([]));
        Assert.Null(_registration.TryRun(["--request-json", "{}"]));
    }
}
