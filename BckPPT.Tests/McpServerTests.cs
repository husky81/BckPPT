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

    private static string AllText(JsonNode reply) =>
        string.Join("\n", reply["result"]!["content"]!.AsArray().Select(c => c!["text"]?.GetValue<string>()));

    [Fact]
    public async Task ImportSlides_InsertsTheOtherFilesPagesAfterTheGivenSlide()
    {
        var (server, ppt) = Create();
        ppt.OpenInWindow = true;
        var source = Path.Combine(Path.GetTempPath(), $"bckppt-import-{Guid.NewGuid():N}.pptx");
        File.WriteAllBytes(source, []);
        try
        {
            ppt.OtherFiles[source] = [.. new[] { "다른 1쪽", "다른 2쪽", "다른 3쪽" }.Select(t => new List<FakeShape> { new("제목", 0, 0, 400, 40, 30, t) })];
            var reply = await Rpc(server, "tools/call",
                new JsonObject
                {
                    ["name"] = "import_slides",
                    ["arguments"] = new JsonObject { ["source"] = source, ["from"] = 2, ["to"] = 3, ["after"] = 1 },
                }.ToJsonString());

            Assert.False(reply["result"]!["isError"]!.GetValue<bool>(), Text(reply));
            Assert.Equal(["1쪽 제목", "다른 2쪽", "다른 3쪽", "2쪽 제목"], ppt.ListSlides(DeckPath).Select(s => s.Title));
            Assert.Contains("2~3쪽으로 붙였습니다", Text(reply));
            Assert.Contains("watch 보고.pptx", ppt.Calls);
        }
        finally { File.Delete(source); }
    }

    [Fact]
    public async Task DeleteSlides_FileWithoutWindow_NeedsConfirm()
    {
        var (server, ppt) = Create();
        ppt.OpenInWindow = false;

        var refused = await Rpc(server, "tools/call", """{"name":"delete_slides","arguments":{"slides":[2]}}""");
        Assert.True(refused["result"]!["isError"]!.GetValue<bool>());
        Assert.Equal(2, ppt.Slides.Count);

        var reply = await Rpc(server, "tools/call", """{"name":"delete_slides","arguments":{"slides":[2],"confirm":true}}""");
        Assert.False(reply["result"]!["isError"]!.GetValue<bool>());
        Assert.Single(ppt.Slides);
        Assert.Contains("2쪽 「2쪽 제목」", Text(reply));
    }

    [Fact]
    public async Task MoveSlide_PutsTheSlideAtTheNewNumber()
    {
        var (server, ppt) = Create();
        ppt.OpenInWindow = true;
        var reply = await Rpc(server, "tools/call", """{"name":"move_slide","arguments":{"slide":2,"to":1}}""");
        Assert.False(reply["result"]!["isError"]!.GetValue<bool>());
        Assert.Equal(["2쪽 제목", "1쪽 제목"], ppt.ListSlides(DeckPath).Select(s => s.Title));
    }

    [Fact]
    public async Task WriteOnReadOnlyWindow_WarnsThatItCannotBeSavedUnderTheSameName()
    {
        var (server, ppt) = Create();
        ppt.OpenInWindow = true;
        ppt.OpenWindows[0] = ppt.OpenWindows[0] with { IsReadOnly = true };

        var listed = await Rpc(server, "tools/call", """{"name":"list_presentations","arguments":{}}""");
        Assert.Contains("읽기 전용", Text(listed));

        var reply = await Rpc(server, "tools/call", """{"name":"duplicate_slide","arguments":{"source":1,"after":2}}""");
        Assert.False(reply["result"]!["isError"]!.GetValue<bool>());
        Assert.Contains("읽기 전용으로 열려 있습니다", AllText(reply));
    }

    [Fact]
    public async Task ReadTable_GivesEveryCellItsAddress()
    {
        var (server, ppt) = Create();
        ppt.Slides[0].Add(new FakeShape("표 1", 0, 50, 300, 100, 30, "항목|값", "진행률|"));
        var reply = await Rpc(server, "tools/call", """{"name":"read_table","arguments":{"slide":1,"shape":"표 1"}}""");
        var text = Text(reply);
        Assert.Contains("표 2행×2열", text);
        Assert.Contains("표 1 [R1C2] | 값", text);
        Assert.Contains("표 1 [R2C2] | (빈 칸)", text);
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
