using System.Text.Json.Nodes;
using BckPPT.Mcp;
using BckPPT.Models;

namespace BckPPT.Tests;

public sealed class McpServerTests
{
    private const string DeckPath = @"C:\Decks\보고.pptx";

    private static (McpServer Server, FakePptComService Ppt) Create()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([new FakeShape("제목", 0, 0, 400, 40, 30, "1쪽 제목")]);
        ppt.Slides.Add([new FakeShape("제목", 0, 0, 400, 40, 30, "2쪽 제목")]);
        ppt.Active = new ActivePresentationInfo("보고.pptx", DeckPath, 2, 1);
        ppt.OpenWindows.Add(new OpenPresentationInfo("보고.pptx", DeckPath, 2, true, true));
        return (new McpServer(new McpToolCatalog(ppt)), ppt);
    }

    private static async Task<JsonNode> Rpc(McpServer server, string method, string paramsJson = "{}")
    {
        var reply = await server.HandleBodyAsync(
            $$"""{"jsonrpc":"2.0","id":1,"method":"{{method}}","params":{{paramsJson}}}""", CancellationToken.None);
        return reply ?? throw new InvalidOperationException("응답이 없습니다.");
    }

    private static string Text(JsonNode reply) => reply["result"]!["content"]![0]!["text"]!.GetValue<string>();

    [Fact]
    public async Task Initialize_AnswersRequestedVersionWithToolsCapability()
    {
        var (server, _) = Create();
        var reply = await Rpc(server, "initialize", """{"protocolVersion":"2025-03-26","capabilities":{},"clientInfo":{"name":"t","version":"1"}}""");
        Assert.Equal("2025-03-26", reply["result"]!["protocolVersion"]!.GetValue<string>());
        Assert.NotNull(reply["result"]!["capabilities"]!["tools"]);
        Assert.Equal("bckppt", reply["result"]!["serverInfo"]!["name"]!.GetValue<string>());
    }

    [Fact]
    public async Task ToolsList_ExposesPptToolsWithFileArgument_AndHidesLlmOnlyTools()
    {
        var (server, _) = Create();
        var tools = (await Rpc(server, "tools/list"))["result"]!["tools"]!.AsArray();
        var names = tools.Select(t => t!["name"]!.GetValue<string>()).ToList();

        Assert.Contains("list_slides", names);
        Assert.Contains("set_paragraphs", names);
        Assert.Contains("list_fonts", names);
        Assert.Contains("render_slide", names);
        Assert.DoesNotContain("finish", names);
        Assert.DoesNotContain("fill_cards", names);
        var listShapes = tools.First(t => t!["name"]!.GetValue<string>() == "list_shapes")!;
        Assert.NotNull(listShapes["inputSchema"]!["properties"]!["file"]);
    }

    [Fact]
    public async Task ToolsCall_WithoutFile_UsesThePresentationTheUserIsLookingAt()
    {
        var (server, _) = Create();
        var reply = await Rpc(server, "tools/call", """{"name":"list_slides","arguments":{}}""");
        Assert.False(reply["result"]!["isError"]!.GetValue<bool>());
        Assert.Contains("2쪽 제목", Text(reply));
    }

    [Fact]
    public async Task ToolsCall_FileNameFragment_PicksTheOpenPresentation()
    {
        var (server, ppt) = Create();
        ppt.Active = null;
        var reply = await Rpc(server, "tools/call", """{"name":"list_slides","arguments":{"file":"보고"}}""");
        Assert.False(reply["result"]!["isError"]!.GetValue<bool>());
    }

    [Fact]
    public async Task ToolsCall_WriteOnOpenWindow_WatchesInsteadOfSaving()
    {
        var (server, ppt) = Create();
        ppt.OpenInWindow = true;
        var reply = await Rpc(server, "tools/call", """{"name":"duplicate_slide","arguments":{"source":1,"after":2}}""");
        Assert.False(reply["result"]!["isError"]!.GetValue<bool>());
        Assert.Contains("watch 보고.pptx", ppt.Calls);
        Assert.Equal(3, ppt.Slides.Count);
    }

    [Fact]
    public async Task ToolsCall_NoPresentation_ReturnsToolErrorNotRpcError()
    {
        var (server, ppt) = Create();
        ppt.Active = null;
        var reply = await Rpc(server, "tools/call", """{"name":"list_slides","arguments":{}}""");
        Assert.True(reply["result"]!["isError"]!.GetValue<bool>());
        Assert.Contains("보고 있는 문서가 없습니다", Text(reply));
    }

    [Fact]
    public async Task Notification_GetsNoReply_AndUnknownMethodIsRpcError()
    {
        var (server, _) = Create();
        Assert.Null(await server.HandleBodyAsync("""{"jsonrpc":"2.0","method":"notifications/initialized"}""", CancellationToken.None));
        var reply = await Rpc(server, "resources/list");
        Assert.Equal(-32601, reply["error"]!["code"]!.GetValue<int>());
    }

    [Theory]
    [InlineData(null, true)]
    [InlineData("http://localhost:3000", true)]
    [InlineData("http://127.0.0.1", true)]
    [InlineData("https://evil.example", false)]
    public void Origin_OnlyLocalPagesMayCall(string? origin, bool allowed) =>
        Assert.Equal(allowed, McpServer.IsAllowedOrigin(origin));
}
