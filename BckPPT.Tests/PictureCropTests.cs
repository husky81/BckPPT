using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json.Nodes;
using BckPPT.Mcp;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Tests;

public sealed class PictureFocusTests
{
    private const float CellRatio = 281.52f / 93.6f;   // 카드 그림 칸 ≈ 3:1

    [Fact]
    public void FocusCrop_Zoom1_IsTheLargestWindowOfTheCellRatio_CenteredOnFocus()
    {
        // 4:3 원본에 3:1 창: 너비 전체, 높이 = (4/3)/3 = 0.444
        var c = PictureFitter.FocusCrop(4032, 3024, null, CellRatio, new FocusPoint(0.5f, 0.3f), 1);
        Assert.Equal(0, c.Left, 3);
        Assert.Equal(0, c.Right, 3);
        float h = 1 - c.Top - c.Bottom;
        Assert.Equal(4f / 3 / CellRatio, h, 3);
        Assert.Equal(0.3f, c.Top + h / 2, 3);
    }

    [Fact]
    public void FocusCrop_Zoom2_HalvesTheWindow_AndPushesItInsideAtTheEdge()
    {
        var c = PictureFitter.FocusCrop(400, 300, null, CellRatio, new FocusPoint(0.95f, 0.02f), 2);
        Assert.Equal(0.5f, 1 - c.Left - c.Right, 3);
        Assert.Equal(0, c.Right, 3);     // 오른쪽 끝에 붙는다
        Assert.Equal(0, c.Top, 3);       // 위 끝에 붙는다
        Assert.True(c.Bottom > 0);
    }

    [Fact]
    public void FocusCrop_AfterUserCrop_StaysInsideTheRest()
    {
        var user = new CropFractions(Left: 0.5f);
        var c = PictureFitter.FocusCrop(400, 300, user, 1, new FocusPoint(0.1f, 0.5f), 1);
        Assert.True(c.Left >= 0.5f - 1e-4);
        // 남은 부분 200×300에서 1:1 가장 큰 창 = 200×200
        Assert.Equal(0.5f, 1 - c.Left - c.Right, 3);
        Assert.Equal(200f / 300, 1 - c.Top - c.Bottom, 3);
    }

    [Fact]
    public void FocusCrop_ZoomBelowOne_IsRefused() =>
        Assert.Throws<ArgumentException>(() => PictureFitter.FocusCrop(400, 300, null, 1, null, 0.5f));

    [Fact]
    public void Place_FocusWithContain_IsRefused() =>
        Assert.Throws<ArgumentException>(() => PictureFitter.Place(
            new PictureRequest("x.png", 0, 0, 100, 100, PictureFit.Contain) { Zoom = 2 }, 400, 300));

    [Fact]
    public void Recrop_KeepFrame_KeepsTheShapeAndChangesOnlyTheCrop()
    {
        var frame = new BoxF(100, 200, 300, 100);
        var (shape, crop) = PictureFitter.Recrop(frame, 300, 225, new CropFractions(Top: 0.2f, Bottom: 0.3556f),
            new PictureCropRequest(Focus: new FocusPoint(0.5f, 0.6f), Zoom: 1));
        Assert.Equal(frame, shape);
        Assert.Equal(0.6f, (crop.Top + 1 - crop.Bottom) / 2, 3);
    }

    [Fact]
    public void Recrop_NotKeepFrame_KeepsScaleAndPicturePosition_LikePowerPointCropTool()
    {
        // 전체 400×300pt 그림이 왼쪽 위 0,0에 있고 자르지 않은 상태. 오른쪽 절반만 남긴다.
        var (shape, crop) = PictureFitter.Recrop(new BoxF(0, 0, 400, 300), 400, 300, CropFractions.None,
            new PictureCropRequest(Crop: new CropFractions(Left: 0.5f), KeepFrame: false));
        Assert.Equal(new BoxF(200, 0, 200, 300), shape);
        Assert.Equal(0.5f, crop.Left, 4);
        // 다시 풀면 원래 자리로 돌아온다.
        var (back, none) = PictureFitter.Recrop(shape, 400, 300, crop, new PictureCropRequest(Reset: true, KeepFrame: false));
        Assert.Equal(new BoxF(0, 0, 400, 300), back);
        Assert.True(none.IsNone);
    }

    [Fact]
    public void Recrop_ResetKeepFrame_FitsTheWholePictureInsideTheFrame()
    {
        var (shape, crop) = PictureFitter.Recrop(new BoxF(0, 0, 300, 100), 300, 225, new CropFractions(Top: 0.3f, Bottom: 0.2556f),
            new PictureCropRequest(Reset: true));
        Assert.True(crop.IsNone);
        Assert.Equal(100, shape.Height, 2);
        Assert.Equal(100 * 300f / 225, shape.Width, 2);
    }

    [Fact]
    public void Grid_2x3_FillsRowByRowWithGaps()
    {
        var cells = PictureFitter.Grid(new BoxF(0, 0, 306, 203), 2, 3, 3);
        Assert.Equal(6, cells.Count);
        Assert.Equal(new BoxF(0, 0, 100, 100), cells[0]);
        Assert.Equal(new BoxF(206, 0, 100, 100), cells[2]);
        Assert.Equal(new BoxF(0, 103, 100, 100), cells[3]);
    }

    [Fact]
    public void AsFocus_RoundTripsThroughFocusCrop()
    {
        var crop = PictureFitter.FocusCrop(400, 300, null, 2, new FocusPoint(0.3f, 0.6f), 1.5f);
        var (focus, zoom) = PicturePreview.AsFocus(crop, 400, 300, 2);
        Assert.Equal(1.5f, zoom, 3);
        var again = PictureFitter.FocusCrop(400, 300, null, 2, focus, zoom);
        Assert.Equal(crop.Left, again.Left, 3);
        Assert.Equal(crop.Top, again.Top, 3);
    }

    [Fact]
    public void Preview_DrawsTheWholePictureWithMargins()
    {
        using var bmp = new Bitmap(800, 600);
        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Png);
        var (png, w, h) = PicturePreview.Render(ms.ToArray(), new CropFractions(0.1f, 0.2f, 0.1f, 0.2f), new CropFractions(Left: 0.5f));
        Assert.Equal((800, 600), (w, h));
        using var result = Image.FromStream(new MemoryStream(png));
        Assert.Equal(PicturePreview.MarginLeft + PicturePreview.LongSide + PicturePreview.MarginRight, result.Width);
    }
}

[Collection("PptDiagnosticLog")]
public sealed class PictureCropMcpToolTests
{
    private const string DeckPath = @"C:\Decks\카드.pptx";
    static PictureCropMcpToolTests() => PptDiagnosticLog.Directory = Path.Combine(Path.GetTempPath(), "bckppt-tests-logs");

    private static (McpServer Server, FakePptComService Ppt, string Image) Create()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([
            new FakeShape("제목", 0, 0, 400, 40, 30, "카드"),
            new FakeShape("그림 칸", 100, 200, 300, 200, 30),
        ]);
        ppt.Active = new ActivePresentationInfo("카드.pptx", DeckPath, 1, 1);
        ppt.OpenWindows.Add(new OpenPresentationInfo("카드.pptx", DeckPath, 1, true, true));
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
    private static bool HasImage(JsonNode reply) => reply["result"]!["content"]!.AsArray().Any(c => c!["type"]!.GetValue<string>() == "image");

    private static async Task<string> AddOne(McpServer server, string image, float width = 300, float height = 100)
    {
        var reply = await Call(server, "add_picture", new JsonObject
        {
            ["slide"] = 1, ["path"] = image, ["left"] = 100, ["top"] = 200, ["width"] = width, ["height"] = height, ["name"] = "사진",
        });
        Assert.False(IsError(reply), Text(reply));
        return "사진";
    }

    [Fact]
    public async Task ToolsList_HasCropAndPreviewTools()
    {
        var (server, _, image) = Create();
        File.Delete(image);
        var reply = await server.HandleBodyAsync("""{"jsonrpc":"2.0","id":1,"method":"tools/list"}""", CancellationToken.None);
        var tools = reply!["result"]!["tools"]!.AsArray();
        var names = tools.Select(t => t!["name"]!.GetValue<string>()).ToList();
        Assert.Contains("set_picture_crop", names);
        Assert.Contains("picture_preview", names);
        Assert.Contains("image_preview", names);
        var crop = tools.First(t => t!["name"]!.GetValue<string>() == "set_picture_crop")!;
        Assert.NotNull(crop["inputSchema"]!["properties"]!["keep_frame"]);
        Assert.NotNull(crop["inputSchema"]!["properties"]!["file"]);
    }

    [Fact]
    public async Task SetPictureCrop_FocusZoom_KeepsTheFrame()
    {
        var (server, ppt, image) = Create();
        try
        {
            await AddOne(server, image);
            var reply = await Call(server, "set_picture_crop", new JsonObject
            {
                ["slide"] = 1, ["shape"] = "사진", ["focus"] = new JsonObject { ["x"] = 0.7, ["y"] = 0.3 }, ["zoom"] = 2,
            });
            Assert.False(IsError(reply), Text(reply));
            var pic = ppt.Slides[0].Single(s => s.Name == "사진");
            Assert.Equal((100f, 200f, 300f, 100f), (pic.Left, pic.Top, pic.Width, pic.Height));
            Assert.Equal(0.5f, 1 - pic.PictureCrop!.Left - pic.PictureCrop.Right, 3);
            Assert.Contains("도형 자리·크기 그대로", Text(reply));
            Assert.Contains("후:", Text(reply));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task SetPictureCrop_ResetWithCrop_IsRefused_AndNonPictureIsRefused()
    {
        var (server, _, image) = Create();
        try
        {
            await AddOne(server, image);
            var both = await Call(server, "set_picture_crop", new JsonObject
            {
                ["slide"] = 1, ["shape"] = "사진", ["reset"] = true, ["crop"] = new JsonObject { ["left"] = 0.1 },
            });
            Assert.True(IsError(both));
            var notPicture = await Call(server, "set_picture_crop", new JsonObject { ["slide"] = 1, ["shape"] = "그림 칸", ["reset"] = true });
            Assert.True(IsError(notPicture));
            Assert.Contains("그림이 아닙니다", Text(notPicture));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task PicturePreview_GivesImageWithCurrentCropAndFocusEquivalent()
    {
        var (server, _, image) = Create();
        try
        {
            await AddOne(server, image);
            var reply = await Call(server, "picture_preview", new JsonObject
            {
                ["slide"] = 1, ["shape"] = "사진", ["focus"] = new JsonObject { ["x"] = 0.5, ["y"] = 0.8 },
            });
            Assert.False(IsError(reply), Text(reply));
            Assert.True(HasImage(reply));
            var text = Text(reply);
            Assert.Contains("원본 400×300px", text);
            Assert.Contains("지금 자름", text);
            Assert.Contains("zoom 1", text);
            Assert.Contains("바꿔 볼 영역", text);
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task ImagePreview_WithCellSize_ShowsTheAreaAddPictureWouldKeep()
    {
        var (server, _, image) = Create();
        try
        {
            var reply = await Call(server, "image_preview", new JsonObject
            {
                ["path"] = image, ["width"] = 281.52, ["height"] = 93.6, ["focus"] = new JsonObject { ["x"] = 0.5, ["y"] = 0.2 },
            });
            Assert.False(IsError(reply), Text(reply));
            Assert.True(HasImage(reply));
            Assert.Contains("원본 400×300px", Text(reply));
            Assert.Contains("y 0.00~", Text(reply));   // 위 끝으로 밀린다

            var noCell = await Call(server, "image_preview", new JsonObject { ["path"] = image, ["zoom"] = 2 });
            Assert.True(IsError(noCell));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task AddPicture_InsetOptions_ArePassedThrough_AndListShowsSourceFile()
    {
        var (server, ppt, image) = Create();
        try
        {
            var reply = await Call(server, "add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = image, ["left"] = 300, ["top"] = 250, ["width"] = 80, ["height"] = 60,
                ["z"] = "front", ["border"] = new JsonObject { ["color"] = "#ffffff", ["width"] = 2 }, ["shadow"] = true,
                ["focus"] = new JsonObject { ["x"] = 0.2, ["y"] = 0.2 }, ["zoom"] = 1.5, ["name"] = "inset",
            });
            Assert.False(IsError(reply), Text(reply));
            var spec = (PictureRequest)ppt.Slides[0].Single(s => s.Name == "inset").Spec!;
            Assert.Equal(PictureZ.Front, spec.Z);
            Assert.Equal(new PictureBorder("#ffffff", 2), spec.Border);
            Assert.True(spec.Shadow);
            Assert.Equal(1.5f, spec.Zoom);

            var listed = await Call(server, "list_shapes", new JsonObject { ["slide"] = 1, ["type"] = "picture" });
            Assert.Contains($"원본 파일 {Path.GetFileName(image)}", Text(listed));

            var bad = await Call(server, "add_picture", new JsonObject
            {
                ["slide"] = 1, ["path"] = image, ["left"] = 0, ["top"] = 0, ["width"] = 80, ["height"] = 60, ["fit"] = "contain", ["zoom"] = 2,
            });
            Assert.True(IsError(bad));
        }
        finally { File.Delete(image); }
    }

    [Fact]
    public async Task AddPicturesInBox_RowsColumns_LaysOutAGrid()
    {
        var (server, ppt, image) = Create();
        try
        {
            var reply = await Call(server, "add_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(Enumerable.Repeat((JsonNode)image, 6).Select(n => n.DeepClone()).ToArray()),
                ["box"] = new JsonObject { ["left"] = 0, ["top"] = 0, ["width"] = 306, ["height"] = 203 }, ["rows"] = 2, ["columns"] = 3,
            });
            Assert.False(IsError(reply), Text(reply));
            Assert.Contains("2×3 격자", Text(reply));
            var pics = ppt.Slides[0].Where(s => s.PictureCrop is not null).ToList();
            Assert.Equal(6, pics.Count);
            Assert.Equal((206f, 103f), (pics[5].Left, pics[5].Top));

            var tooMany = await Call(server, "add_pictures_in_box", new JsonObject
            {
                ["slide"] = 1, ["paths"] = new JsonArray(Enumerable.Repeat((JsonNode)image, 7).Select(n => n.DeepClone()).ToArray()),
                ["box"] = new JsonObject { ["left"] = 0, ["top"] = 0, ["width"] = 306, ["height"] = 203 }, ["rows"] = 2, ["columns"] = 3,
            });
            Assert.True(IsError(tooMany));
        }
        finally { File.Delete(image); }
    }
}
