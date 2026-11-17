using System.Drawing;
using System.Drawing.Imaging;
using System.IO.Compression;
using System.Text.Json.Nodes;
using BckPPT.Agent.Tools;
using BckPPT.Mcp;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Tests;

public sealed class PicturePlanModelTests
{
    [Fact]
    public void PlacedState_RoundTrips_AndFindsWhatTheUserChanged()
    {
        var state = new PlacedState(new BoxF(49.68f, 220.32f, 281.52f, 93.6f), new CropFractions(0, 0.2778f, 0, 0.2778f));
        var back = PlacedState.Parse(state.Format())!;
        Assert.Null(back.Differences(new BoxF(49.68f, 220.32f, 281.52f, 93.6f), new CropFractions(0, 0.2778f, 0, 0.2778f)));
        Assert.Equal("크기", back.Differences(new BoxF(49.68f, 220.32f, 200, 93.6f), new CropFractions(0, 0.2778f, 0, 0.2778f)));
        Assert.Equal("자리·자르기", back.Differences(new BoxF(60, 220.32f, 281.52f, 93.6f), new CropFractions(0.1f, 0.2778f, 0, 0.2778f)));
        Assert.Null(PlacedState.Parse("엉망"));
    }

    [Fact]
    public void EffectiveDpi_CountsTheCrop()
    {
        // 300px 그림을 1인치(72pt) 칸에 자르지 않고: 300dpi. 가로 절반만 남기면 150dpi.
        Assert.Equal(300, PictureFitter.EffectiveDpi(300, 72, CropFractions.None), 1);
        Assert.Equal(150, PictureFitter.EffectiveDpi(300, 72, new CropFractions(Left: 0.5f)), 1);
    }

    internal static void Same(CropFractions expected, CropFractions? actual)
    {
        Assert.NotNull(actual);
        Assert.Equal(expected.Left, actual!.Left, 4);
        Assert.Equal(expected.Top, actual.Top, 4);
        Assert.Equal(expected.Right, actual.Right, 4);
        Assert.Equal(expected.Bottom, actual.Bottom, 4);
    }

    [Fact]
    public void FromKeep_IsOneMinusTheFarEdge() =>
        Same(new CropFractions(0.2f, 0.1f, 0.1f, 0.3f), PictureFitter.FromKeep(0.2f, 0.1f, 0.9f, 0.7f));

    [Fact]
    public void Anchor_BottomRight_InsideTheReferenceWithMargin()
    {
        var box = PictureFitter.Anchor(new BoxF(100, 200, 300, 100), "bottom-right", 0.4f, 3, 4f / 3);
        Assert.Equal(120, box.Width, 2);
        Assert.Equal(90, box.Height, 2);
        Assert.Equal(100 + 300 - 3 - 120, box.Left, 2);
        Assert.Equal(200 + 100 - 3 - 90, box.Top, 2);
        // 높이가 넘치면 높이에 맞춰 줄인다.
        var tall = PictureFitter.Anchor(new BoxF(0, 0, 300, 100), "top-left", 0.9f, 3, 1);
        Assert.Equal(94, tall.Height, 2);
        Assert.Throws<ArgumentException>(() => PictureFitter.Anchor(new BoxF(0, 0, 300, 100), "왼쪽", 0.4f, 3, 1));
    }

    [Fact]
    public void ParsePlan_ReadsSpecsAndResolvesRelativePathsFromTheFolderAbove()
    {
        var root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"bckppt-plan-{Guid.NewGuid():N}"));
        try
        {
            var scripts = root.CreateSubdirectory("_scripts");
            Directory.CreateDirectory(Path.Combine(root.FullName, "11_사진"));
            File.WriteAllBytes(Path.Combine(root.FullName, "11_사진", "a.png"), []);
            var plan = McpToolCatalog.ParsePlan([
                "# 주석",
                "11|AS-IS|11_사진/a.png@keep=0.1,0.2,0.9,0.8;11_사진/b.png@focus=0.5,0.3,1.5",
                "",
                "12|TO-BE|C:\\x\\c.jpg@crop: 0.1, 0, 0.1, 0",
            ], scripts.FullName);
            Assert.Equal(2, plan.Count);
            Assert.Equal(Path.Combine(root.FullName, "11_사진", "a.png"), plan[0].Images[0].Path);   // 위 폴더에 있으면 그쪽
            Same(new CropFractions(0.1f, 0.2f, 0.1f, 0.2f), plan[0].Images[0].Crop);
            Assert.Equal(new FocusPoint(0.5f, 0.3f), plan[0].Images[1].Focus);
            Assert.Equal(1.5f, plan[0].Images[1].Zoom);
            Same(new CropFractions(0.1f, 0, 0.1f, 0), plan[1].Images[0].Crop);
            Assert.Throws<AgentToolException>(() => McpToolCatalog.ParsePlan(["11|AS-IS"], null));
            Assert.Throws<AgentToolException>(() => McpToolCatalog.ParsePlan(["11|AS-IS|a.png"], null));
            Assert.Throws<AgentToolException>(() => McpToolCatalog.ParsePlan(["11|AS-IS|C:\\a.png@keep=0.5,0,0.4,1"], null));
        }
        finally { root.Delete(true); }
    }

    [Fact]
    public void ReadDocumentSize_ListsPicturesBySlideWithCropAndSharing()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bckppt-size-{Guid.NewGuid():N}.pptx");
        try
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                void Text(string name, string body) { using var w = new StreamWriter(zip.CreateEntry(name).Open()); w.Write(body); }
                const string rel = "http://schemas.openxmlformats.org/package/2006/relationships";
                const string ns = "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
                Text("ppt/presentation.xml", $"<p:presentation {ns}><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/></p:sldIdLst></p:presentation>");
                Text("ppt/_rels/presentation.xml.rels", $"<Relationships xmlns=\"{rel}\"><Relationship Id=\"rId1\" Type=\"slide\" Target=\"slides/slide1.xml\"/><Relationship Id=\"rId2\" Type=\"slide\" Target=\"slides/slide2.xml\"/></Relationships>");
                Text("ppt/slides/slide1.xml", $"<p:sld {ns}><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id=\"5\" name=\"큰 사진\"/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId9\"/><a:srcRect t=\"27778\" b=\"27778\"/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>");
                Text("ppt/slides/slide2.xml", $"<p:sld {ns}><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id=\"7\" name=\"같은 사진\"/></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId9\"/><a:srcRect/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>");
                foreach (var n in new[] { "slide1", "slide2" })
                    Text($"ppt/slides/_rels/{n}.xml.rels", $"<Relationships xmlns=\"{rel}\"><Relationship Id=\"rId9\" Type=\"image\" Target=\"../media/image1.png\"/></Relationships>");
                using var bmp = new Bitmap(64, 48);
                using var ms = new MemoryStream();
                bmp.Save(ms, ImageFormat.Png);
                using (var w = zip.CreateEntry("ppt/media/image1.png").Open()) w.Write(ms.ToArray());
            }
            var size = PptxPictureReader.ReadDocumentSize(path)!;
            Assert.Equal(1, size.MediaCount);
            Assert.Equal(2, size.Pictures.Count);
            Assert.True(size.Pictures[0].HasCrop);
            Assert.False(size.Pictures[1].HasCrop);
            Assert.Equal(2, size.Pictures[0].SharedBy);
            Assert.Equal((2, 7, 64), (size.Pictures[1].Slide, size.Pictures[1].ShapeId, size.Pictures[1].PixelWidth));
        }
        finally { File.Delete(path); }
    }
}

[Collection("PptDiagnosticLog")]
public sealed class PicturePlanMcpToolTests
{
    private const string DeckPath = @"C:\Decks\카드.pptx";
    static PicturePlanMcpToolTests() => PptDiagnosticLog.Directory = Path.Combine(Path.GetTempPath(), "bckppt-tests-logs");

    private static (McpServer Server, FakePptComService Ppt, string Image) Create()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([
            new FakeShape("제목", 0, 0, 400, 40, 30, "카드"),
            new FakeShape("AS-IS 칸", 100, 200, 300, 100, 30),
            new FakeShape("[그림 삽입]", 150, 230, 200, 30, 30, "[그림 삽입]"),
        ]);
        ppt.Slides.Add([new FakeShape("제목", 0, 0, 400, 40, 30, "2쪽")]);
        ppt.Active = new ActivePresentationInfo("카드.pptx", DeckPath, 2, 1);
        ppt.OpenWindows.Add(new OpenPresentationInfo("카드.pptx", DeckPath, 2, true, true));
        ppt.OpenInWindow = true;
        var image = Path.Combine(Path.GetTempPath(), $"bckppt-pic-{Guid.NewGuid():N}.png");
        using (var bmp = new Bitmap(400, 300)) bmp.Save(image, ImageFormat.Png);
        return (new McpServer(new McpToolCatalog(ppt)), ppt, image);
    }

    private static async Task<JsonNode> Call(McpServer server, string tool, JsonObject args)
    {
        var body = new JsonObject
        {
            ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "tools/call",
            ["params"] = new JsonObject { ["name"] = tool, ["arguments"] = args },
        };
        return (await server.HandleBodyAsync(body.ToJsonString(), CancellationToken.None))!;
    }

    private static string Text(JsonNode reply) =>
        string.Join("\n", reply["result"]!["content"]!.AsArray().Select(c => c!["text"]?.GetValue<string>()));
    private static bool IsError(JsonNode reply) => reply["result"]!["isError"]!.GetValue<bool>();
    private static byte[]? Image(JsonNode reply) => reply["result"]!["content"]!.AsArray()
        .Where(c => c!["type"]!.GetValue<string>() == "image").Select(c => Convert.FromBase64String(c!["data"]!.GetValue<string>())).FirstOrDefault();
    private static JsonArray Paths(string image, int n) => new(Enumerable.Range(0, n).Select(_ => (JsonNode)image).ToArray());

    [Fact]
    public async Task WindowlessDocument_IsNamedInTheError_AndShowWindowOpensIt()
    {
        var (server, ppt, image) = Create();
        File.Delete(image);
        ppt.Active = null;
        ppt.OpenWindows.Clear();
        ppt.Windowless.Add(new OpenPresentationInfo("카드.pptx", DeckPath, 2, false, true));

        var miss = await Call(server, "list_slides", new JsonObject { ["file"] = "카드" });
        Assert.True(IsError(miss));
        Assert.Contains("창 없이", Text(miss));
        Assert.Contains("show_window", Text(miss));
        Assert.Contains("창 없음", Text(await Call(server, "list_presentations", new JsonObject())));

        var shown = await Call(server, "list_slides", new JsonObject { ["file"] = "카드", ["show_window"] = true });
        Assert.False(IsError(shown), Text(shown));
        Assert.Contains("open 카드.pptx unblock=False", ppt.Calls);
    }

    [Fact]
    public async Task ReplacePicturesInBox_RemovesOldPicturesAndPlaceholder_KeepsTheCell()
    {
        var (server, ppt, image) = Create();
        try
        {
            Assert.False(IsError(await Call(server, "add_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = Paths(image, 2), ["box"] = "AS-IS 칸", ["names"] = new JsonArray("옛 1", "옛 2"),
            })));
            var reply = await Call(server, "replace_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = Paths(image, 1), ["box"] = "AS-IS 칸", ["names"] = new JsonArray("새"),
                ["keeps"] = new JsonArray(new JsonObject { ["x0"] = 0.1, ["y0"] = 0.1, ["x1"] = 0.9, ["y1"] = 0.6 }),
            });
            Assert.False(IsError(reply), Text(reply));
            var names = ppt.Slides[0].Select(s => s.Name).ToList();
            Assert.DoesNotContain("옛 1", names);
            Assert.DoesNotContain("옛 2", names);
            Assert.DoesNotContain("[그림 삽입]", names);
            Assert.Contains("AS-IS 칸", names);
            Assert.Contains("새", names);
            Assert.Contains("지움 3개", Text(reply));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task UserChangedPicture_StopsReplace_SkipOrForce()
    {
        var (server, ppt, image) = Create();
        try
        {
            await Call(server, "add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = image, ["left"] = 100, ["top"] = 200, ["width"] = 300, ["height"] = 100, ["name"] = "사진",
            });
            ppt.Slides[0].Single(s => s.Name == "사진").Width = 250;   // 사용자가 줄였다

            var listed = await Call(server, "list_shapes", new JsonObject { ["slide"] = 1, ["type"] = "picture" });
            Assert.Contains("사용자가 바꿈(크기)", Text(listed));

            var refused = await Call(server, "replace_picture", new JsonObject { ["slide"] = 1, ["shape"] = "사진", ["path"] = image });
            Assert.True(IsError(refused));
            Assert.Contains("force=true", Text(refused));

            var boxRefused = await Call(server, "replace_pictures_in_box", new JsonObject { ["slide"] = 1, ["paths"] = Paths(image, 1), ["box"] = "AS-IS 칸" });
            Assert.True(IsError(boxRefused));
            var skipped = await Call(server, "replace_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = Paths(image, 1), ["box"] = "AS-IS 칸", ["skip_user_modified"] = true,
            });
            Assert.False(IsError(skipped));
            Assert.Contains("건너뛰었습니다", Text(skipped));
            Assert.Equal(250, ppt.Slides[0].Single(s => s.Name == "사진").Width);

            var forced = await Call(server, "replace_picture", new JsonObject { ["slide"] = 1, ["shape"] = "사진", ["path"] = image, ["force"] = true });
            Assert.False(IsError(forced), Text(forced));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task ApplyPicturePlan_UsesCellDefinitions_ReportsTable_AndSheet()
    {
        var (server, ppt, image) = Create();
        ppt.RenderPng = true;
        try
        {
            var args = new JsonObject
            {
                ["plan"] = new JsonArray($"1|AS-IS|{image};{image}@keep=0,0,0.5,1", $"2|TO-BE|{image}", $"3|AS-IS|{image}", "1|없는칸|" + image),
                ["cells"] = new JsonObject
                {
                    ["AS-IS"] = "AS-IS 칸",
                    ["TO-BE"] = new JsonObject { ["left"] = 400, ["top"] = 200, ["width"] = 280, ["height"] = 90 },
                },
            };
            var dry = await Call(server, "apply_picture_plan", (JsonObject)args.DeepClone().AsObject().Also(a => a["dry_run"] = true));
            Assert.False(IsError(dry), Text(dry));
            Assert.Contains("| 할 일 |", Text(dry));
            Assert.DoesNotContain(ppt.Slides[0], s => s.PictureCrop is not null);

            var reply = await Call(server, "apply_picture_plan", args);
            Assert.False(IsError(reply), Text(reply));
            var text = Text(reply);
            Assert.Contains("| 1 | AS-IS | 넣음 |", text);
            Assert.Contains("| 2 | TO-BE | 넣음 |", text);
            Assert.Contains("| 3 | AS-IS | 오류 |", text);    // 3쪽 없음
            Assert.Contains("| 1 | 없는칸 | 오류 |", text);
            Assert.Equal(2, ppt.Slides[0].Count(s => s.PictureCrop is not null));
            Assert.Single(ppt.Slides[1], s => s.PictureCrop is not null);
            Assert.NotNull(Image(reply));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task AddPicture_Anchor_PutsItInTheCorner_AndWarnsLowDpi()
    {
        var (server, ppt, image) = Create();
        try
        {
            var reply = await Call(server, "add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = image, ["name"] = "inset",
                ["anchor"] = new JsonObject { ["shape"] = "AS-IS 칸", ["corner"] = "bottom-right", ["size"] = 0.25, ["margin"] = 3 },
                ["source"] = "https://example.org/a", ["license"] = "CC BY 4.0",
            });
            Assert.False(IsError(reply), Text(reply));
            var inset = ppt.Slides[0].Single(s => s.Name == "inset");
            Assert.Equal(75, inset.Width, 1);
            Assert.Equal(56.25f, inset.Height, 1);
            Assert.Equal(100 + 300 - 3 - 75, inset.Left, 1);
            Assert.Equal(PictureZ.Front, ((PictureRequest)inset.Spec!).Z);

            // 400px 원본을 300pt 칸에: 96dpi → 경고
            var big = await Call(server, "add_picture", new JsonObject
            {
                ["slide"] = 2, ["path"] = image, ["left"] = 0, ["top"] = 0, ["width"] = 300, ["height"] = 225,
            });
            Assert.Contains("해상도 낮음", Text(big));

            var credits = await Call(server, "export_picture_credits", new JsonObject { ["write_notes"] = true });
            Assert.False(IsError(credits), Text(credits));
            Assert.Contains("https://example.org/a (CC BY 4.0)", Text(credits));
            Assert.Contains("출처 미기록", Text(credits));
            Assert.Equal(["그림 출처: https://example.org/a (CC BY 4.0)"], ppt.Notes[1]);
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task ImageSheet_And_RenderRegion_ReturnPictures()
    {
        var (server, _, image) = Create();
        try
        {
            var sheet = await Call(server, "image_sheet", new JsonObject { ["paths"] = Paths(image, 3), ["columns"] = 2 });
            Assert.False(IsError(sheet), Text(sheet));
            Assert.Contains("400×300px", Text(sheet));
            Assert.NotNull(Image(sheet));

            var region = await Call(server, "render_region", new JsonObject { ["slide"] = 1, ["shape"] = "AS-IS 칸", ["scale"] = 2, ["margin"] = 0 });
            Assert.False(IsError(region), Text(region));
            using var png = System.Drawing.Image.FromStream(new MemoryStream(Image(region)!));
            Assert.Equal((600, 200), (png.Width, png.Height));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task CompressPictures_DryRunFirst_ThenConfirmReplaces()
    {
        var (server, ppt, image) = Create();
        try
        {
            await Call(server, "add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = image, ["left"] = 100, ["top"] = 200, ["width"] = 60, ["height"] = 30, ["name"] = "작게",
            });
            var dry = await Call(server, "compress_pictures", new JsonObject { ["dpi"] = 150 });
            Assert.False(IsError(dry), Text(dry));
            Assert.Contains("confirm=true", Text(dry));
            Assert.DoesNotContain(ppt.Calls, c => c.StartsWith("picture ") && c.Contains("1-"));
            int before = ppt.Calls.Count(c => c.StartsWith("picture "));

            var done = await Call(server, "compress_pictures", new JsonObject { ["dpi"] = 150, ["confirm"] = true });
            Assert.False(IsError(done), Text(done));
            Assert.Contains("적용: 1개", Text(done));
            Assert.Equal(before + 1, ppt.Calls.Count(c => c.StartsWith("picture ")));
            var pic = ppt.Slides[0].Single(s => s.Name == "작게");
            Assert.True(pic.PictureCrop!.IsNone);   // 잘린 부분을 버렸다
        }
        finally { File.Delete(image); }
    }
}

internal static class JsonTestExtensions
{
    public static JsonObject Also(this JsonObject o, Action<JsonObject> change) { change(o); return o; }
}
