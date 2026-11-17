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
        var deck = Path.Combine(dir, $"live-run-{DateTime.Now:HHmmss}.pptx");   // 지난 사본을 PowerPoint가 쥐고 있을 수 있다
        File.Copy(Path.Combine(dir, "live-pictures.pptx"), deck, overwrite: true);
        var ppt = new PptComService();
        var server = new McpServer(new McpToolCatalog(ppt));
        var log = new StringBuilder();
        int callNo = 0;

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
            callNo++;
            log.AppendLine($"## {tool} {args.ToJsonString()}");
            foreach (var c in reply["result"]!["content"]!.AsArray())
            {
                if (c!["type"]!.GetValue<string>() == "image")
                {
                    var png = Path.Combine(dir, $"out-{callNo:00}-{tool}-{++image}.png");
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

            // ── 2026-11-17b: 자르기 바꾸기·미리보기·격자·inset ──
            Assert.True(Ok(await Call("picture_preview", new JsonObject { ["slide"] = 1, ["shape"] = "큰 사진 cover" })));   // ① 태그의 원본 파일
            Assert.True(Ok(await Call("picture_preview", new JsonObject
            {
                ["slide"] = 1, ["shape"] = "큰 사진 cover", ["focus"] = new JsonObject { ["x"] = 0.3, ["y"] = 0.7 }, ["zoom"] = 2,
            })));
            Assert.True(Ok(await Call("set_picture_crop", new JsonObject
            {
                ["slide"] = 1, ["shape"] = "큰 사진 cover", ["focus"] = new JsonObject { ["x"] = 0.3, ["y"] = 0.7 }, ["zoom"] = 2,
            })));
            Assert.True(Ok(await Call("picture_preview", new JsonObject { ["slide"] = 1, ["shape"] = "큰 사진 cover" })));
            Assert.True(Ok(await Call("set_picture_crop", new JsonObject
            {
                ["slide"] = 1, ["shape"] = "오른쪽 절반 cover", ["crop"] = new JsonObject { ["left"] = 0.25, ["top"] = 0.1 }, ["keep_frame"] = false,
            })));
            Assert.True(Ok(await Call("set_picture_crop", new JsonObject { ["slide"] = 1, ["shape"] = "작은 사진 contain", ["reset"] = true })));
            Assert.True(Ok(await Call("image_preview", new JsonObject
            {
                ["path"] = Path.Combine(dir, "big43.png"), ["width"] = 281.52, ["height"] = 93.6, ["focus"] = new JsonObject { ["x"] = 0.5, ["y"] = 0.25 },
            })));
            Assert.True(Ok(await Call("add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = Path.Combine(dir, "small43.png"), ["left"] = 300, ["top"] = 350, ["width"] = 70, ["height"] = 52.5,
                ["z"] = "front", ["border"] = new JsonObject { ["color"] = "#FFFFFF", ["width"] = 2 }, ["shadow"] = true, ["name"] = "inset",
            })));
            Assert.True(Ok(await Call("add_pictures_in_box", new JsonObject
            {
                ["slide"] = 2, ["paths"] = new JsonArray(Enumerable.Range(0, 6).Select(i => (JsonNode)Path.Combine(dir, i % 2 == 0 ? "big43.png" : "tall.png")).ToArray()),
                ["box"] = new JsonObject { ["left"] = 480, ["top"] = 120, ["width"] = 420, ["height"] = 300 }, ["rows"] = 2, ["columns"] = 3,
                ["focuses"] = new JsonArray(null, new JsonObject { ["x"] = 0.5, ["y"] = 0.1 }), ["zoom"] = 1.2,
            })));
            await Call("list_shapes", new JsonObject { ["slide"] = 1, ["type"] = "picture" });
            Assert.True(Ok(await Call("render_slides", new JsonObject { ["slides"] = new JsonArray(1, 2) })));

            // ② 원본 파일이 없어진 그림 → 저장된 문서의 그림
            var gone = Path.Combine(dir, "gone.png");
            File.Copy(Path.Combine(dir, "small43.png"), gone, overwrite: true);
            Assert.True(Ok(await Call("add_picture", new JsonObject
            {
                ["slide"] = 2, ["path"] = gone, ["left"] = 40, ["top"] = 300, ["width"] = 120, ["height"] = 60, ["name"] = "사라진 원본",
            })));
            File.Delete(gone);
            Assert.True(Ok(await Call("picture_preview", new JsonObject { ["slide"] = 2, ["shape"] = "사라진 원본" })));
            // ③ 저장하지 않은 그림 + 원본 없음 → PowerPoint에서 뽑는다
            using (ppt.KeepOpen(deck, readOnly: false))
            {
                File.Copy(Path.Combine(dir, "tall.png"), gone, overwrite: true);
                Assert.True(Ok(await Call("add_picture", new JsonObject
                {
                    ["slide"] = 2, ["path"] = gone, ["left"] = 200, ["top"] = 300, ["width"] = 120, ["height"] = 60, ["name"] = "저장 안 한 그림",
                })));
                File.Delete(gone);
                Assert.True(Ok(await Call("picture_preview", new JsonObject { ["slide"] = 2, ["shape"] = "저장 안 한 그림" })));
                await Call("list_shapes", new JsonObject { ["slide"] = 2, ["type"] = "picture" });
            }

            // ── 2026-11-17c: 칸 통째로 바꾸기·사용자 수정 알아보기·계획·inset 기준·출처·크기·영역 렌더 ──
            Assert.True(Ok(await Call("set_geometry", new JsonObject { ["slide"] = 1, ["shape"] = "AS-IS 1", ["width"] = 120 })));   // 「사용자」가 줄임
            await Call("list_shapes", new JsonObject { ["slide"] = 1, ["type"] = "picture", ["in_box"] = "AS-IS 칸" });
            Assert.False(Ok(await Call("replace_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(Path.Combine(dir, "small43.png")), ["box"] = "AS-IS 칸",
            })));
            Assert.True(Ok(await Call("replace_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(Path.Combine(dir, "small43.png")), ["box"] = "AS-IS 칸", ["skip_user_modified"] = true,
            })));
            Assert.True(Ok(await Call("replace_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(Path.Combine(dir, "big43.png"), Path.Combine(dir, "tall.png")), ["box"] = "AS-IS 칸", ["force"] = true,
                ["keeps"] = new JsonArray(new JsonObject { ["x0"] = 0.4, ["y0"] = 0.3, ["x1"] = 0.9, ["y1"] = 0.7 }), ["source"] = "시험 촬영", ["license"] = "자체",
            })));
            Assert.True(Ok(await Call("add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = Path.Combine(dir, "small43.png"), ["name"] = "기준 inset",
                ["anchor"] = new JsonObject { ["shape"] = "큰 사진 cover", ["corner"] = "bottom-right", ["size"] = 0.3, ["margin"] = 3 },
                ["border"] = new JsonObject { ["color"] = "#FFFFFF", ["width"] = 1.5 }, ["shadow"] = true, ["source"] = "https://example.org/inset",
            })));
            var planFile = Path.Combine(dir, "plan-live.txt");
            await File.WriteAllLinesAsync(planFile, [
                "# 시험 계획",
                "2|오른쪽|big43.png@focus=0.5,0.5,1.5;small43.png@keep=0,0,0.5,1",
                "2|아래|tall.png",
                "9|오른쪽|big43.png",
            ]);
            var cells = new JsonObject
            {
                ["오른쪽"] = new JsonObject { ["left"] = 480, ["top"] = 120, ["width"] = 420, ["height"] = 300 },
                ["아래"] = new JsonObject { ["left"] = 40, ["top"] = 430, ["width"] = 300, ["height"] = 90 },
            };
            Assert.True(Ok(await Call("apply_picture_plan", new JsonObject { ["plan_path"] = planFile, ["cells"] = cells.DeepClone(), ["dry_run"] = true })));
            Assert.True(Ok(await Call("apply_picture_plan", new JsonObject { ["plan_path"] = planFile, ["cells"] = cells.DeepClone() })));
            Assert.True(Ok(await Call("render_region", new JsonObject { ["slide"] = 1, ["shape"] = "AS-IS 칸" })));
            Assert.True(Ok(await Call("image_sheet", new JsonObject
            {
                ["paths"] = new JsonArray(Path.Combine(dir, "big43.png"), Path.Combine(dir, "small43.png"), Path.Combine(dir, "tall.png")),
            })));
            Assert.True(Ok(await Call("export_picture_credits", new JsonObject { ["write_notes"] = true })));
            await Call("list_shapes", new JsonObject { ["slide"] = 1, ["type"] = "picture" });
            Assert.True(Ok(await Call("document_size", new JsonObject { ["top"] = 8 })));
            Assert.True(Ok(await Call("compress_pictures", new JsonObject { ["dpi"] = 150 })));
            Assert.True(Ok(await Call("compress_pictures", new JsonObject { ["dpi"] = 150, ["confirm"] = true })));
            Assert.True(Ok(await Call("document_size", new JsonObject { ["top"] = 8 })));
            Assert.True(Ok(await Call("render_slides", new JsonObject { ["slides"] = new JsonArray(1, 2) })));
            // 창 없이 남은 문서(지난 시험 사본 등) 목록 — 읽기만 한다.
            log.AppendLine("창 없는 문서: " + string.Join(", ", ppt.ListWindowlessPresentations().Select(o => o.Name)));

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
