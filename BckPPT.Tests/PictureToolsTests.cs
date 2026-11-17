using System.IO.Compression;
using System.Text.Json.Nodes;
using BckPPT.Mcp;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Tests;

public sealed class PictureFitterTests
{
    private static readonly BoxF Cell = new(100, 200, 281.52f, 93.6f);   // 카드 그림 칸(3:1)

    [Fact]
    public void Cover_4x3PhotoInto3x1Cell_CropsTopAndBottomEvenly()
    {
        var (shape, crop) = PictureFitter.Fit(4032, 3024, null, Cell, PictureFit.Cover);
        Assert.Equal(Cell, shape);
        Assert.Equal(0, crop.Left, 3);
        Assert.Equal(0, crop.Right, 3);
        // 남는 높이 = 281.52/(4/3)=211.14 중 93.6 → 위아래 각 (1-93.6/211.14)/2
        Assert.Equal((1 - 93.6 / 211.14) / 2, crop.Top, 3);
        Assert.Equal(crop.Top, crop.Bottom, 4);
    }

    [Fact]
    public void Cover_TallPhotoIntoWideCell_CropsTopAndBottom_WideIntoTall_CropsSides()
    {
        var (_, wide) = PictureFitter.Fit(300, 600, null, new BoxF(0, 0, 200, 100), PictureFit.Cover);
        Assert.True(wide.Top > 0.3f && wide.Left == 0);
        var (_, tall) = PictureFitter.Fit(600, 300, null, new BoxF(0, 0, 100, 200), PictureFit.Cover);
        Assert.True(tall.Left > 0.3f && tall.Top == 0);
    }

    [Fact]
    public void Contain_KeepsRatioCenteredInCell_WithoutCropping()
    {
        var (shape, crop) = PictureFitter.Fit(400, 300, null, Cell, PictureFit.Contain);
        Assert.True(crop.IsNone);
        Assert.Equal(93.6f, shape.Height, 2);
        Assert.Equal(93.6f * 4 / 3, shape.Width, 2);
        Assert.Equal(Cell.Left + (Cell.Width - shape.Width) / 2, shape.Left, 2);
    }

    [Fact]
    public void Stretch_FillsCellWithUserCropOnly()
    {
        var user = new CropFractions(0.1f, 0, 0.1f, 0);
        var (shape, crop) = PictureFitter.Fit(400, 300, user, Cell, PictureFit.Stretch);
        Assert.Equal(Cell, shape);
        Assert.Equal(user, crop);
    }

    [Fact]
    public void Cover_AfterUserCrop_CentersTheRemainingPart()
    {
        // 오른쪽 절반만 남기고(left 0.5) 칸(1:1)을 채운다: 남은 부분은 200×300 → 위아래를 더 자른다.
        var (_, crop) = PictureFitter.Fit(400, 300, new CropFractions(Left: 0.5f), new BoxF(0, 0, 100, 100), PictureFit.Cover);
        Assert.Equal(0.5f, crop.Left, 3);
        Assert.Equal(0f, crop.Right, 3);
        Assert.Equal((1 - 200f / 300f) / 2, crop.Top, 3);
    }

    [Fact]
    public void CropObject_RoundTrips()
    {
        var box = new BoxF(10, 20, 150, 50);
        var crop = new CropFractions(0.1f, 0.2f, 0.05f, 0.3f);
        var (pw, ph, ox, oy) = PictureFitter.ToCropObject(box, crop);
        var back = PictureFitter.FromCropObject(pw, ph, ox, oy, box.Width, box.Height);
        Assert.Equal(crop.Left, back.Left, 4);
        Assert.Equal(crop.Top, back.Top, 4);
        Assert.Equal(crop.Right, back.Right, 4);
        Assert.Equal(crop.Bottom, back.Bottom, 4);
    }

    [Fact]
    public void Split_TwoPicturesSideBySideWithGap()
    {
        var cells = PictureFitter.Split(Cell, 2, 3, null);
        Assert.Equal(2, cells.Count);
        Assert.Equal((281.52f - 3) / 2, cells[0].Width, 3);
        Assert.Equal(cells[0].Left + cells[0].Width + 3, cells[1].Left, 3);
        Assert.Equal(Cell.Left + Cell.Width, cells[1].Left + cells[1].Width, 3);
        Assert.All(cells, c => Assert.Equal(Cell.Height, c.Height));
        Assert.Equal(2, PictureFitter.Split(new BoxF(0, 0, 50, 200), 2, 4, null).Select(c => c.Top).Distinct().Count());
    }
}

public sealed class PptxPictureReaderTests
{
    private static byte[] Png(int w, int h)
    {
        using var bmp = new System.Drawing.Bitmap(w, h);
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        return ms.ToArray();
    }

    private static byte[] Jpeg(int w, int h)
    {
        using var bmp = new System.Drawing.Bitmap(w, h);
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
        return ms.ToArray();
    }

    [Fact]
    public void ReadImageSize_PngAndJpeg()
    {
        Assert.Equal((123, 45), PptxPictureReader.ReadImageSize(new MemoryStream(Png(123, 45))));
        Assert.Equal((321, 54), PptxPictureReader.ReadImageSize(new MemoryStream(Jpeg(321, 54))));
        Assert.Null(PptxPictureReader.ReadImageSize(new MemoryStream([1, 2, 3, 4])));
    }

    [Fact]
    public void ReadPixelSizes_FindsPictureBySlideIdAndShapeId()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bckppt-pixels-{Guid.NewGuid():N}.pptx");
        try
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                void Text(string name, string body) { using var w = new StreamWriter(zip.CreateEntry(name).Open()); w.Write(body); }
                void Bytes(string name, byte[] body) { using var w = zip.CreateEntry(name).Open(); w.Write(body); }
                const string rel = "http://schemas.openxmlformats.org/package/2006/relationships";
                Text("ppt/presentation.xml", """
                    <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <p:sldIdLst><p:sldId id="256" r:id="rId7"/><p:sldId id="300" r:id="rId8"/></p:sldIdLst>
                    </p:presentation>
                    """);
                Text("ppt/_rels/presentation.xml.rels", $"""
                    <Relationships xmlns="{rel}">
                      <Relationship Id="rId7" Type="slide" Target="slides/slide1.xml"/>
                      <Relationship Id="rId8" Type="slide" Target="slides/slide2.xml"/>
                    </Relationships>
                    """);
                Text("ppt/slides/slide2.xml", """
                    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                      <p:cSld><p:spTree>
                        <p:pic><p:nvPicPr><p:cNvPr id="5" name="사진"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></p:pic>
                        <p:pic><p:nvPicPr><p:cNvPr id="6" name="다른 사진"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></p:pic>
                      </p:spTree></p:cSld>
                    </p:sld>
                    """);
                Text("ppt/slides/_rels/slide2.xml.rels", $"""
                    <Relationships xmlns="{rel}">
                      <Relationship Id="rId2" Type="image" Target="../media/image1.png"/>
                      <Relationship Id="rId3" Type="image" Target="../media/image2.jpeg"/>
                    </Relationships>
                    """);
                Bytes("ppt/media/image1.png", Png(4032, 3024));
                Bytes("ppt/media/image2.jpeg", Jpeg(640, 480));
            }

            var sizes = PptxPictureReader.ReadPixelSizes(path, 300, [5, 6, 9]);
            Assert.Equal((4032, 3024), sizes[5]);
            Assert.Equal((640, 480), sizes[6]);
            Assert.False(sizes.ContainsKey(9));
            Assert.Empty(PptxPictureReader.ReadPixelSizes(path, 999, [5]));
        }
        finally { File.Delete(path); }
    }
}

[Collection("PptDiagnosticLog")]   // McpServerTests가 같은 로그 파일을 읽는다
public sealed class PictureMcpToolTests
{
    private const string DeckPath = @"C:\Decks\카드.pptx";
    private static readonly string LogDir = Path.Combine(Path.GetTempPath(), "bckppt-tests-logs");
    static PictureMcpToolTests() => PptDiagnosticLog.Directory = LogDir;

    private static (McpServer Server, FakePptComService Ppt, string Image) Create()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([
            new FakeShape("제목", 0, 0, 400, 40, 30, "카드"),
            new FakeShape("AS-IS 칸", 100, 200, 281.52f, 93.6f, 30),
            new FakeShape("[그림 삽입]", 100, 200, 281.52f, 93.6f, 30, "[그림 삽입]"),
            new FakeShape("TO-BE 칸", 400, 200, 281.52f, 93.6f, 30),
        ]);
        ppt.Active = new ActivePresentationInfo("카드.pptx", DeckPath, 1, 1);
        ppt.OpenWindows.Add(new OpenPresentationInfo("카드.pptx", DeckPath, 1, true, true));
        ppt.OpenInWindow = true;
        var image = Path.Combine(Path.GetTempPath(), $"bckppt-pic-{Guid.NewGuid():N}.png");
        File.WriteAllBytes(image, [0x89, (byte)'P', (byte)'N', (byte)'G']);
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

    [Fact]
    public async Task ToolsList_HasPictureAndRenderTools()
    {
        var (server, _, image) = Create();
        File.Delete(image);
        var reply = await server.HandleBodyAsync("""{"jsonrpc":"2.0","id":1,"method":"tools/list"}""", CancellationToken.None);
        var names = reply!["result"]!["tools"]!.AsArray().Select(t => t!["name"]!.GetValue<string>()).ToList();
        Assert.Contains("add_picture", names);
        Assert.Contains("add_pictures_in_box", names);
        Assert.Contains("replace_picture", names);
        Assert.Contains("render_slides", names);
    }

    [Fact]
    public async Task AddPicture_CoversTheCellAndReportsAddressSizeAndCrop()
    {
        var (server, ppt, image) = Create();
        try
        {
            var reply = await Call(server, "add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = image, ["left"] = 100, ["top"] = 200, ["width"] = 281.52, ["height"] = 93.6, ["name"] = "AS-IS 사진",
            });
            Assert.False(IsError(reply), Text(reply));
            Assert.Contains("AS-IS 사진 | 100,200 281.52×93.6", Text(reply));
            Assert.Contains("잘림 L0% T", Text(reply));
            Assert.Contains("원본 400×300px", Text(reply));
            Assert.Contains("watch 카드.pptx", ppt.Calls);
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task AddPicture_RejectsMissingFileAndBadCrop()
    {
        var (server, _, image) = Create();
        try
        {
            var missing = await Call(server, "add_picture", new JsonObject
            { ["slide"] = 1, ["path"] = image + ".none.png", ["left"] = 0, ["top"] = 0, ["width"] = 10, ["height"] = 10 });
            Assert.True(IsError(missing));
            Assert.Contains("그림 파일이 없습니다", Text(missing));
            var crop = await Call(server, "add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = image, ["left"] = 0, ["top"] = 0, ["width"] = 10, ["height"] = 10,
                ["crop"] = new JsonObject { ["left"] = 0.6, ["right"] = 0.5 },
            });
            Assert.True(IsError(crop));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task AddPicturesInBox_SplitsThePlaceholderAndReplacesItAtItsZOrder()
    {
        var (server, ppt, image) = Create();
        try
        {
            var reply = await Call(server, "add_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(image, image), ["box"] = "[그림 삽입]", ["replace_placeholder"] = true,
            });
            Assert.False(IsError(reply), Text(reply));
            var names = ppt.Slides[0].Select(s => s.Name).ToList();
            Assert.DoesNotContain("[그림 삽입]", names);
            // 지운 글상자 자리(2번째 뒤, TO-BE 칸 앞)에 두 장이 들어간다.
            Assert.Equal(["제목", "AS-IS 칸"], names.Take(2));
            Assert.All(ppt.Slides[0].Skip(2).Take(2), s => Assert.NotNull(s.PictureCrop));
            Assert.Equal("TO-BE 칸", names[4]);
            var pics = ppt.Slides[0].Where(s => s.PictureCrop is not null).ToList();
            Assert.Equal((281.52f - 3) / 2, pics[0].Width, 2);
            Assert.Equal(100 + (281.52f - 3) / 2 + 3, pics[1].Left, 2);
            Assert.Contains("「[그림 삽입]」는 지웠습니다", Text(reply));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task AddPicturesInBox_CellAsBox_PlaceholderGivenSeparately()
    {
        var (server, ppt, image) = Create();
        try
        {
            ppt.Slides[0].Single(s => s.Name == "[그림 삽입]").Height = 20;   // 글에 맞춰 줄어든 글상자
            var reply = await Call(server, "add_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(image), ["box"] = "AS-IS 칸", ["replace_placeholder"] = "[그림 삽입]",
            });
            Assert.False(IsError(reply), Text(reply));
            Assert.DoesNotContain(ppt.Slides[0], s => s.Name == "[그림 삽입]");
            Assert.Contains(ppt.Slides[0], s => s.Name == "AS-IS 칸");
            Assert.Equal(93.6f, ppt.Slides[0].Single(s => s.PictureCrop is not null).Height, 2);
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task ReplacePicture_KeepsPlaceSizeNameAndOrder()
    {
        var (server, ppt, image) = Create();
        try
        {
            await Call(server, "add_picture", new JsonObject
            { ["slide"] = 1, ["path"] = image, ["left"] = 400, ["top"] = 200, ["width"] = 281.52, ["height"] = 93.6, ["name"] = "TO-BE 사진" });
            ppt.Slides[0].Add(new FakeShape("맨 위 글", 0, 500, 100, 20, 30, "글"));
            var reply = await Call(server, "replace_picture", new JsonObject
            { ["slide"] = 1, ["shape"] = "TO-BE 사진", ["path"] = image, ["fit"] = "contain" });
            Assert.False(IsError(reply), Text(reply));
            var names = ppt.Slides[0].Select(s => s.Name).ToList();
            Assert.Equal(names.Count - 2, names.IndexOf("TO-BE 사진"));
            var pic = ppt.Slides[0].Single(s => s.Name == "TO-BE 사진");
            Assert.Equal(93.6f, pic.Height, 2);   // contain: 높이에 맞춤
            Assert.Contains("같은 자리·크기·Z순서", Text(reply));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task ListShapes_FiltersByBoxAndPictureType_AndShowsCropAndPixels()
    {
        var (server, _, image) = Create();
        try
        {
            await Call(server, "add_picture", new JsonObject
            { ["slide"] = 1, ["path"] = image, ["left"] = 100, ["top"] = 200, ["width"] = 281.52, ["height"] = 93.6, ["name"] = "AS-IS 사진" });
            await Call(server, "add_picture", new JsonObject
            { ["slide"] = 1, ["path"] = image, ["left"] = 400, ["top"] = 200, ["width"] = 281.52, ["height"] = 93.6, ["name"] = "TO-BE 사진" });

            var reply = await Call(server, "list_shapes", new JsonObject
            {
                ["slide"] = 1, ["type"] = "picture",
                ["in_box"] = new JsonObject { ["left"] = 90, ["top"] = 190, ["width"] = 300, ["height"] = 120 },
            });
            var text = Text(reply);
            Assert.False(IsError(reply), text);
            Assert.Contains("AS-IS 사진", text);
            Assert.DoesNotContain("TO-BE 사진", text);
            Assert.DoesNotContain("[그림 삽입]", text);
            Assert.Contains("문서에 든 그림 400×300px", text);
            Assert.Contains("걸러 봄", text);

            var byShape = Text(await Call(server, "list_shapes", new JsonObject { ["slide"] = 1, ["in_box"] = "TO-BE 칸" }));
            Assert.Contains("TO-BE 사진", byShape);
            Assert.DoesNotContain("AS-IS 사진", byShape);
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task OpenPresentation_ProtectedView_IsReported_AndUnblockReopensForEditing()
    {
        var (server, ppt, image) = Create();
        File.Delete(image);
        var deck = Path.Combine(Path.GetTempPath(), $"bckppt-open-{Guid.NewGuid():N}.pptx");
        File.WriteAllBytes(deck, []);
        try
        {
            ppt.OpenState = OpenedState.ProtectedView;
            ppt.DownloadMark = true;
            var first = Text(await Call(server, "open_presentation", new JsonObject { ["path"] = deck }));
            Assert.Contains("제한된 보기", first);
            Assert.Contains("unblock=true", first);

            var second = Text(await Call(server, "open_presentation", new JsonObject { ["path"] = deck, ["unblock"] = true }));
            Assert.Contains("편집 가능", second);
            Assert.Contains("받은 파일 표시를 지웠습니다", second);
            Assert.Contains($"open {Path.GetFileName(deck)} unblock=True", ppt.Calls);
        }
        finally { File.Delete(deck); }
    }

    [Fact]
    public async Task ToolOnDocumentInProtectedView_SaysWhyItCannotBeFound()
    {
        var (server, ppt, image) = Create();
        File.Delete(image);
        ppt.Active = null;
        ppt.OpenWindows.Clear();
        ppt.ProtectedViews.Add(new ProtectedViewInfo("메일첨부.pptx", @"C:\Downloads\메일첨부.pptx"));
        var reply = await Call(server, "list_slides", new JsonObject { ["file"] = "메일첨부" });
        Assert.True(IsError(reply));
        Assert.Contains("제한된 보기", Text(reply));
        var list = Text(await Call(server, "list_presentations", new JsonObject()));
        Assert.Contains("메일첨부.pptx · 「제한된 보기」", list);
    }

    [Fact]
    public async Task RenderSlides_ChecksSlideNumbers()
    {
        var (server, _, image) = Create();
        File.Delete(image);
        var reply = await Call(server, "render_slides", new JsonObject { ["slides"] = new JsonArray(1, 5) });
        Assert.True(IsError(reply));
        Assert.Contains("5쪽 없음", Text(reply));
    }
}
