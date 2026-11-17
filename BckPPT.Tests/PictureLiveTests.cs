using System.Text;
using System.Text.Json.Nodes;
using BckPPT.Mcp;
using BckPPT.Services;

namespace BckPPT.Tests;

/// <summary>
/// 실제 PowerPoint로 그림 도구를 돌려 본다. BCKPPT_LIVE_PICTURES에 폴더를 주면 돈다:
/// live-pictures.pptx(1쪽에 「AS-IS 칸」·「[그림 삽입]」·「TO-BE 칸」·「[그림 삽입] 2」, 2쪽), big43.png(4032×3024)·small43.png·tall.png.
/// 사본(live-run.pptx)에 쓰고 결과 글·렌더 그림을 그 폴더에 남긴다. 사용자 문서는 건드리지 않는다.
/// BCKPPT_LIVE_PROTECTED=1이면 받은 파일 표시를 붙인 사본을 창에 열어 「제한된 보기」 감지·unblock도 본다(창이 잠깐 뜬다).
/// </summary>
[Collection("PptDiagnosticLog")]
public sealed class PictureLiveTests
{
    [Fact]
    public async Task PictureTools_OnRealPowerPoint()
    {
        var dir = Environment.GetEnvironmentVariable("BCKPPT_LIVE_PICTURES");
        if (string.IsNullOrEmpty(dir)) return;
        PptDiagnosticLog.Directory = Path.Combine(dir, "logs");
        var deck = Path.Combine(dir, "live-run.pptx");
        File.Copy(Path.Combine(dir, "live-pictures.pptx"), deck, overwrite: true);
        var ppt = new PptComService();
        var server = new McpServer(new McpToolCatalog(ppt));
        var log = new StringBuilder();

        async Task<JsonNode> Call(string tool, JsonObject args)
        {
            args["file"] = deck;
            var body = new JsonObject
            {
                ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "tools/call",
                ["params"] = new JsonObject { ["name"] = tool, ["arguments"] = args },
            };
            var reply = (await server.HandleBodyAsync(body.ToJsonString(), CancellationToken.None))!;
            int image = 0;
            log.AppendLine($"## {tool} {args.ToJsonString()}");
            foreach (var c in reply["result"]!["content"]!.AsArray())
            {
                if (c!["type"]!.GetValue<string>() == "image")
                {
                    var png = Path.Combine(dir, $"out-{tool}-{++image}.png");
                    await File.WriteAllBytesAsync(png, Convert.FromBase64String(c["data"]!.GetValue<string>()));
                    log.AppendLine($"(그림 {png})");
                }
                else log.AppendLine(c["text"]!.GetValue<string>());
            }
            log.AppendLine($"isError={reply["result"]!["isError"]}");
            return reply;
        }
        bool Ok(JsonNode r) => !r["result"]!["isError"]!.GetValue<bool>();

        try
        {
            Assert.True(Ok(await Call("add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = Path.Combine(dir, "big43.png"), ["left"] = 100, ["top"] = 330, ["width"] = 281.52, ["height"] = 93.6,
                ["name"] = "큰 사진 cover",
            })));
            Assert.True(Ok(await Call("add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = Path.Combine(dir, "small43.png"), ["left"] = 420, ["top"] = 330, ["width"] = 281.52, ["height"] = 93.6,
                ["fit"] = "contain", ["name"] = "작은 사진 contain",
            })));
            Assert.True(Ok(await Call("add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = Path.Combine(dir, "small43.png"), ["left"] = 720, ["top"] = 330, ["width"] = 200, ["height"] = 100,
                ["crop"] = new JsonObject { ["left"] = 0.5 }, ["name"] = "오른쪽 절반 cover",
            })));
            Assert.True(Ok(await Call("add_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(Path.Combine(dir, "big43.png"), Path.Combine(dir, "small43.png")),
                ["box"] = "AS-IS 칸", ["replace_placeholder"] = "[그림 삽입]", ["names"] = new JsonArray("AS-IS 1", "AS-IS 2"),
            })));
            Assert.True(Ok(await Call("add_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(Path.Combine(dir, "tall.png")),
                ["box"] = "[그림 삽입] 2", ["replace_placeholder"] = true,
            })));
            Assert.True(Ok(await Call("replace_picture", new JsonObject
            {
                ["slide"] = 1, ["shape"] = "AS-IS 2", ["path"] = Path.Combine(dir, "tall.png"),
            })));
            var listed = await Call("list_shapes", new JsonObject
            {
                ["slide"] = 1, ["type"] = "picture", ["in_box"] = "AS-IS 칸",
            });
            Assert.True(Ok(listed));
            await Call("list_shapes", new JsonObject { ["slide"] = 1 });
            Assert.True(Ok(await Call("render_slide", new JsonObject { ["slide"] = 1 })));
            Assert.True(Ok(await Call("render_slides", new JsonObject { ["slides"] = new JsonArray(1, 2), ["sheet"] = true })));
            Assert.True(Ok(await Call("render_slides", new JsonObject { ["from"] = 1, ["to"] = 2 })));

            if (Environment.GetEnvironmentVariable("BCKPPT_LIVE_PROTECTED") == "1")
            {
                var pv = Path.Combine(dir, "live-protected.pptx");
                File.Copy(Path.Combine(dir, "live-pictures.pptx"), pv, overwrite: true);
                File.WriteAllText(pv + ":Zone.Identifier", "[ZoneTransfer]\r\nZoneId=3\r\n");
                Assert.True(PptComService.HasDownloadMark(pv));
                async Task<JsonNode> Open(JsonObject args)
                {
                    var body = new JsonObject
                    {
                        ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "tools/call",
                        ["params"] = new JsonObject { ["name"] = "open_presentation", ["arguments"] = args },
                    };
                    var reply = (await server.HandleBodyAsync(body.ToJsonString(), CancellationToken.None))!;
                    log.AppendLine($"## open_presentation {args.ToJsonString()}");
                    log.AppendLine(reply["result"]!["content"]![0]!["text"]!.GetValue<string>());
                    return reply;
                }
                await Open(new JsonObject { ["path"] = pv, ["wait_seconds"] = 30 });
                var protectedViews = ppt.ListProtectedViews();
                log.AppendLine("제한된 보기 창: " + string.Join(", ", protectedViews.Select(v => $"{v.Name}|{v.Path}")));
                var listSlides = new JsonObject
                {
                    ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "tools/call",
                    ["params"] = new JsonObject { ["name"] = "list_slides", ["arguments"] = new JsonObject { ["file"] = "live-protected" } },
                };
                McpToolCatalog.RecentOpenWindow = TimeSpan.Zero;   // 다시 찾기를 끄고 바로 오류 글을 본다
                var miss = (await server.HandleBodyAsync(listSlides.ToJsonString(), CancellationToken.None))!;
                log.AppendLine("## list_slides(제한된 보기 문서)\n" + miss["result"]!["content"]![0]!["text"]!.GetValue<string>());
                await Open(new JsonObject { ["path"] = pv, ["unblock"] = true, ["wait_seconds"] = 30 });
                log.AppendLine($"표시 남음: {PptComService.HasDownloadMark(pv)}");
                log.AppendLine("열린 창: " + string.Join(", ", ppt.ListOpenWindows().Select(o => $"{o.Name}{(o.IsReadOnly ? "(읽기 전용)" : "")}")));
            }
        }
        finally
        {
            await File.WriteAllTextAsync(Path.Combine(dir, "live-result.txt"), log.ToString());
        }
    }
}
