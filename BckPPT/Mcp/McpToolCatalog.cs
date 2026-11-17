using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using BckPPT.Agent.Tools;
using BckPPT.Llm;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Mcp;

/// <summary>MCP 도구 결과 한 조각. 글 또는 그림(base64 PNG).</summary>
public sealed record McpContent(string Type, string? Text = null, string? Data = null, string? MimeType = null)
{
    public static McpContent FromText(string text) => new("text", Text: text);
    public static McpContent FromImage(byte[] png) => new("image", Data: Convert.ToBase64String(png), MimeType: "image/png");
}

public sealed record McpToolResult(IReadOnlyList<McpContent> Content, bool IsError)
{
    public static McpToolResult Ok(string text) => new([McpContent.FromText(text)], false);
    public static McpToolResult Error(string text) => new([McpContent.FromText(text)], true);
}

public sealed record McpTool(string Name, string Description, JsonObject InputSchema,
    Func<JsonElement, CancellationToken, Task<McpToolResult>> Call);

/// <summary>
/// 외부 에이전트(Claude Code 등)에 내보내는 도구. 판단은 그 에이전트가 하고 BckPPT는 PowerPoint를 정확히 다룬다.
/// 내부 에이전트의 PPT 도구를 그대로 쓰되, 대상 문서를 고르는 <c>file</c> 인자를 붙인다.
/// 열린 창의 문서는 그 자리에서 고치고 저장하지 않는다(저장은 사용자가 한다). 창이 없는 파일은 고친 뒤 저장한다.
/// LLM이 필요한 도구(카드 채우기·원문 요약)와 내부 루프 전용 도구(finish)는 내보내지 않는다.
/// </summary>
public sealed class McpToolCatalog
{
    private static readonly HashSet<string> Excluded = new(StringComparer.Ordinal) { "finish", "fill_cards", "render_slide" };
    private static readonly string[] PresentationExtensions = [".pptx", ".pptm", ".ppt"];

    private const string FileDescription =
        "대상 PPT. 절대 경로나 PowerPoint에 열린 문서 이름 조각. 비우면 PowerPoint에서 지금 보고 있는 문서";

    private readonly IPptComService _ppt;
    private readonly Dictionary<string, McpTool> _tools = new(StringComparer.Ordinal);
    private readonly ToolRegistry _agentTools;

    public McpToolCatalog(IPptComService ppt)
    {
        _ppt = ppt;
        _agentTools = new ToolRegistry(PptAgentTools.Create(includeSources: false).Where(t => !Excluded.Contains(t.Name)));

        Add("list_presentations", "PowerPoint 창에 열린 문서 목록(이름·경로·쪽 수·보고 있는 문서 표시)",
            Schema(), (_, _) => Task.FromResult(ListPresentations()));
        Add("open_presentation",
            "PPT 파일을 PowerPoint 창에 열고 문서가 올라올 때까지 기다린다(큰 파일도 다음 도구가 바로 찾게). " +
            "메일·인터넷에서 받은 파일이 「제한된 보기」로 열리면 결과에 알린다 — 그 상태에서는 다른 도구가 문서를 못 찾으니, " +
            "사용자에게 확인받고 unblock=true로 다시 부르면 받은 파일 표시를 지우고 편집 모드로 다시 연다",
            Schema(("path", "string", "PPT 절대 경로"),
                ("unblock", "boolean", "true면 인터넷에서 받은 파일 표시(Zone.Identifier)를 지우고, 제한된 보기 창이 있으면 닫고 편집 모드로 연다"),
                ("wait_seconds", "integer", "문서가 올라오길 기다릴 최대 초(기본 60)")), OpenPresentation, "path");
        Add("go_to_slide", "PowerPoint 창을 그 쪽으로 옮겨 사용자에게 보여 준다",
            WithFile(Schema(("slide", "integer", "쪽 번호(1부터)")), "slide"), GoToSlide);
        Add("list_fonts",
            "문서에 쓰인 글꼴 목록(글꼴마다 쓰인 횟수·쪽). font_name을 주면 그 글꼴을 쓰는 쪽·도형·글 앞부분을 모두 찾아 준다",
            WithFile(Schema(("font_name", "string", "찾을 글꼴 이름(정확히). 비우면 목록"))), ListFonts);
        Add("export_markdown", "문서를 슬라이드별 Markdown으로 내보낸다(쪽마다 PNG 미리보기를 _assets 폴더에 함께 저장)",
            WithFile(Schema(("output_path", "string", "저장할 .md 절대 경로. 비우면 PPT 옆에 같은 이름으로"))), ExportMarkdown);
        Add("render_slide", "슬라이드를 PNG로 렌더해 그림으로 돌려준다. 고친 뒤 넘침·겹침을 눈으로 확인할 때 쓴다",
            WithFile(Schema(("slide", "integer", "쪽 번호(1부터)")), "slide"), RenderSlide);
        Add("render_slides",
            "여러 쪽을 한 번에 렌더한다. slides(쪽 번호 배열) 또는 from~to. sheet=true면 여러 쪽을 한 장에 모은 축소판(쪽 번호 표시)으로 돌려준다. " +
            $"따로 받는 그림은 {MaxSeparateRenders}쪽까지, 넘으면 축소판으로 모은다(최대 {MaxSheetRenders}쪽)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slides"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "integer" }, ["description"] = "쪽 번호들(1부터)" },
                    ["from"] = new JsonObject { ["type"] = "integer", ["description"] = "slides 대신 범위의 첫 쪽" },
                    ["to"] = new JsonObject { ["type"] = "integer", ["description"] = "범위의 끝 쪽" },
                    ["sheet"] = new JsonObject { ["type"] = "boolean", ["description"] = "true면 한 장에 모은 축소판" },
                    ["columns"] = new JsonObject { ["type"] = "integer", ["description"] = "축소판 한 줄의 쪽 수(기본 4)" },
                },
            }), RenderSlides);

        var cropSchema = new JsonObject
        {
            ["type"] = "object",
            ["description"] = "원본 그림 기준으로 먼저 잘라 낼 비율(0~1). 핵심 부분만 크게 보이게 할 때. 예: {\"left\":0.2,\"right\":0.1}",
            ["properties"] = new JsonObject
            {
                ["left"] = new JsonObject { ["type"] = "number" }, ["top"] = new JsonObject { ["type"] = "number" },
                ["right"] = new JsonObject { ["type"] = "number" }, ["bottom"] = new JsonObject { ["type"] = "number" },
            },
        };
        var focusSchema = new JsonObject
        {
            ["type"] = "object",
            ["description"] = "가운데 둘 점(원본 기준 0~1, 0,0이 왼쪽 위). 칸 비율로 자를 창을 그 점 가운데에 두고 가장자리를 넘으면 안쪽으로 민다. picture_preview·image_preview의 눈금으로 정한다",
            ["properties"] = new JsonObject { ["x"] = new JsonObject { ["type"] = "number" }, ["y"] = new JsonObject { ["type"] = "number" } },
        };
        var keepSchema = new JsonObject
        {
            ["type"] = "object",
            ["description"] = "남길 영역(원본 기준 0~1) {x0,y0,x1,y1}. crop(잘라 낼 양)과 둘 중 하나. 예: {\"x0\":0.2,\"y0\":0.1,\"x1\":0.9,\"y1\":0.7}",
            ["properties"] = new JsonObject
            {
                ["x0"] = new JsonObject { ["type"] = "number" }, ["y0"] = new JsonObject { ["type"] = "number" },
                ["x1"] = new JsonObject { ["type"] = "number" }, ["y1"] = new JsonObject { ["type"] = "number" },
            },
        };
        JsonObject Str(string d) => new() { ["type"] = "string", ["description"] = d };
        const string SourceDescription = "그림 출처(URL·문서명). 태그로 남아 export_picture_credits가 쪽별 출처 목록을 만든다";
        const string LicenseDescription = "이용 조건(예: CC BY 4.0, 자체 촬영)";
        const string ZoomDescription = "확대(1 이상). 1 = 칸을 겨우 채움(cover와 같음), 2 = 두 배 확대(창이 절반). focus 없이 주면 가운데";
        const string CropFocusOrder = "crop과 focus·zoom을 같이 주면 crop을 먼저 적용하고 남은 부분에서 focus·zoom을 잡는다";
        JsonObject FitSchema(string note) => new()
        {
            ["type"] = "string", ["enum"] = new JsonArray("cover", "contain", "stretch"),
            ["description"] = "cover: 칸을 채우고 넘는 쪽을 가운데 기준으로 자른다. contain: 비율 유지, 칸 안 가운데. stretch: 칸 크기로 늘린다. " + note,
        };
        Add("add_picture",
            "그림 파일을 쪽에 넣는다(문서에 포함). 칸(left·top·width·height, pt)에 fit으로 맞춘다. 열린 창의 문서면 그 자리에 넣고 저장은 사용자가 한다. " +
            "자르기는 PowerPoint 자르기라 원본이 남고 set_picture_crop으로 다시 고칠 수 있다. focus·zoom으로 보여 줄 부분을 고른다(" + CropFocusOrder + "). " +
            "자리는 left·top·width·height 또는 anchor(기준 도형 모서리). 결과에 도형 주소·실제 자리·크기·자른 양·원본 픽셀·실제 해상도(dpi, 150 미만 경고)를 준다 (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["path"] = new JsonObject { ["type"] = "string", ["description"] = "그림 파일 절대 경로(png·jpg·gif·bmp·tif·emf·wmf·svg)" },
                    ["left"] = new JsonObject { ["type"] = "number" }, ["top"] = new JsonObject { ["type"] = "number" },
                    ["width"] = new JsonObject { ["type"] = "number" }, ["height"] = new JsonObject { ["type"] = "number" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["crop"] = cropSchema.DeepClone(),
                    ["focus"] = focusSchema.DeepClone(),
                    ["zoom"] = new JsonObject { ["type"] = "number", ["description"] = ZoomDescription },
                    ["keep"] = keepSchema.DeepClone(),
                    ["name"] = new JsonObject { ["type"] = "string", ["description"] = "도형 이름" },
                    ["source"] = Str(SourceDescription), ["license"] = Str(LicenseDescription),
                    ["anchor"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["description"] = "기준 도형 안 모서리에 붙여 넣는다(left·top·width·height 대신). 큰 그림 위 inset에. 높이는 그림 비율(crop·keep 반영)로 정한다",
                        ["properties"] = new JsonObject
                        {
                            ["shape"] = Str("기준 도형 주소(list_shapes의 값)"),
                            ["corner"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("top-left", "top-right", "bottom-left", "bottom-right", "center", "top", "bottom", "left", "right"), ["description"] = "기본 bottom-right" },
                            ["size"] = new JsonObject { ["type"] = "number", ["description"] = "기준 도형 너비 대비 너비(기본 0.4)" },
                            ["margin"] = new JsonObject { ["type"] = "number", ["description"] = "기준 도형 안쪽 여백 pt(기본 3)" },
                        },
                    },
                    ["z"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("front", "back"), ["description"] = "front: 맨 앞(기본, 큰 그림 위 inset), back: 맨 뒤" },
                    ["border"] = new JsonObject
                    {
                        ["type"] = "object", ["description"] = "테두리. 예: {\"color\":\"#FFFFFF\",\"width\":1.5}",
                        ["properties"] = new JsonObject { ["color"] = new JsonObject { ["type"] = "string", ["description"] = "#RRGGBB(기본 #FFFFFF)" }, ["width"] = new JsonObject { ["type"] = "number", ["description"] = "pt(기본 1.5)" } },
                    },
                    ["shadow"] = new JsonObject { ["type"] = "boolean", ["description"] = "true면 옅은 바깥 그림자(큰 그림 위에 올린 inset이 떠 보이게)" },
                },
            }, "slide", "path"), AddPicture);
        Add("add_pictures_in_box",
            "칸 하나를 그림 수만큼 나눠(가로로 긴 칸은 좌우, 세로로 긴 칸은 위아래) gap 간격으로 나란히 넣는다. rows·columns를 주면 격자(왼쪽 위부터 줄 단위)로 넣는다. " +
            "box는 {left,top,width,height} 또는 기존 도형 주소(그 도형 자리·크기를 쓴다). replace_placeholder=true면 그 도형(예: 「[그림 삽입]」 글상자)을, " +
            "도형 주소면 그 도형을 지우고 그 Z순서 자리에 넣는다 (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["paths"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "그림 파일 절대 경로들(넣는 순서: 왼쪽→오른쪽 또는 위→아래)" },
                    ["box"] = new JsonObject { ["description"] = "칸: {left,top,width,height}(pt) 또는 list_shapes의 도형 주소" },
                    ["gap"] = new JsonObject { ["type"] = "number", ["description"] = "그림 사이 간격 pt(기본 3)" },
                    ["direction"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("auto", "horizontal", "vertical"), ["description"] = "나누는 방향(기본 auto)" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["rows"] = new JsonObject { ["type"] = "integer", ["description"] = "격자 줄 수. columns만 주면 그림 수로 정한다" },
                    ["columns"] = new JsonObject { ["type"] = "integer", ["description"] = "격자 칸 수. rows만 주면 그림 수로 정한다. 예: 6장을 rows 2, columns 3" },
                    ["crops"] = new JsonObject { ["type"] = "array", ["items"] = cropSchema.DeepClone(), ["description"] = "그림마다 crop(paths 순서). 없으면 null" },
                    ["keeps"] = new JsonObject { ["type"] = "array", ["items"] = keepSchema.DeepClone(), ["description"] = "그림마다 남길 영역(crops 대신). 없으면 null" },
                    ["focuses"] = new JsonObject { ["type"] = "array", ["items"] = focusSchema.DeepClone(), ["description"] = "그림마다 focus(paths 순서). 없으면 null. " + CropFocusOrder },
                    ["sources"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "그림마다 출처(source보다 먼저)" },
                    ["source"] = Str(SourceDescription), ["license"] = Str(LicenseDescription),
                    ["zoom"] = new JsonObject { ["type"] = "number", ["description"] = "모든 그림의 zoom. " + ZoomDescription },
                    ["zooms"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "number" }, ["description"] = "그림마다 zoom(zoom보다 먼저)" },
                    ["names"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "그림마다 도형 이름" },
                    ["replace_placeholder"] = new JsonObject
                    {
                        ["description"] = "true면 box 도형을, 도형 주소(문자열)면 그 도형을 지우고 그 Z순서 자리에 넣는다. " +
                                          "「[그림 삽입]」 글상자가 글에 맞춰 줄어 칸보다 작으면 box에는 칸 도형, 여기에는 글상자 주소를 준다",
                    },
                },
            }, "slide", "paths", "box"), AddPicturesInBox);
        Add("replace_picture",
            "도형(보통 그림)을 다른 그림 파일로 바꾼다. 같은 자리·크기·Z순서를 지키고 이름·테두리·그림자도 이어받는다. fit 기본 cover. " +
            "BckPPT가 넣은 뒤 사용자가 바꾼 그림이면 멈추고 알린다(force=true로 덮어씀) (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["shape"] = new JsonObject { ["type"] = "string", ["description"] = "바꿀 도형 주소(list_shapes의 값)" },
                    ["path"] = new JsonObject { ["type"] = "string", ["description"] = "새 그림 파일 절대 경로" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["crop"] = cropSchema.DeepClone(),
                    ["focus"] = focusSchema.DeepClone(),
                    ["zoom"] = new JsonObject { ["type"] = "number", ["description"] = ZoomDescription },
                    ["keep"] = keepSchema.DeepClone(),
                    ["name"] = new JsonObject { ["type"] = "string", ["description"] = "새 이름. 비우면 원래 이름" },
                    ["source"] = Str(SourceDescription), ["license"] = Str(LicenseDescription),
                    ["force"] = new JsonObject { ["type"] = "boolean", ["description"] = "사용자가 자리·크기·자르기를 바꾼 그림도 덮어쓴다(사용자 확인 뒤)" },
                },
            }, "slide", "shape", "path"), ReplacePicture);
        JsonObject CropChange(bool forPreview) => new()
        {
            ["crop"] = new JsonObject
            {
                ["type"] = "object",
                ["description"] = "원본 기준 잘라 낼 비율(0~1) {left,top,right,bottom}. keep_frame이면 남은 부분을 칸 비율에 맞춰 가운데 기준으로 더 자른다",
                ["properties"] = cropSchema["properties"]!.DeepClone(),
            },
            ["keep"] = keepSchema.DeepClone(),
            ["focus"] = focusSchema.DeepClone(),
            ["zoom"] = new JsonObject { ["type"] = "number", ["description"] = ZoomDescription },
            ["reset"] = new JsonObject { ["type"] = "boolean", ["description"] = "true면 자르기를 없앤다(keep_frame이면 칸 안에 그림 전체가 들어가게 줄인다)" },
            ["keep_frame"] = new JsonObject
            {
                ["type"] = "boolean",
                ["description"] = "기본 true: 도형 자리·크기 그대로 자른 영역만 바꾼다. false: 지금 배율·그림 자리를 지키고 도형이 새 자른 영역을 따라간다(PowerPoint 자르기 도구처럼)" +
                                  "",
            },
        };
        Add("set_picture_crop",
            "이미 넣은 그림의 자른 영역을 바꾼다(PowerPoint 자르기, 원본 보존). crop(원본 기준 비율) 또는 focus+zoom(칸 비율로 그 점이 가운데 오게) 또는 reset. " +
            CropFocusOrder + ". 고치기 전에 picture_preview로 원본과 지금 자른 영역을 보고 값을 정한다 (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = Merge(new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["shape"] = new JsonObject { ["type"] = "string", ["description"] = "그림 도형 주소(list_shapes의 값)" },
                }, CropChange(false)),
            }, "slide", "shape"), SetPictureCrop);
        Add("picture_preview",
            "넣은 그림의 원본 전체를 줄인 PNG로 돌려준다. 원본 기준 0.1 간격 격자·눈금(0~1), 지금 자른 영역(빨간 실선, 바깥은 어둡게)을 그린다. " +
            "crop·focus·zoom·reset을 주면 set_picture_crop이 만들 영역을 노란 점선으로 미리 그린다(문서는 고치지 않는다). 결과에 원본 픽셀·지금 crop 비율·같은 영역의 focus/zoom 값",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = Merge(new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["shape"] = new JsonObject { ["type"] = "string", ["description"] = "그림 도형 주소(list_shapes의 값)" },
                }, CropChange(true)),
            }, "slide", "shape"), PicturePreviewTool);
        Add("image_preview",
            "넣기 전 그림 파일을 줄인 PNG로 돌려준다. 원본 기준 0.1 간격 격자·눈금(0~1)을 그린다. width·height(넣을 칸 pt)를 주면 add_picture가 " +
            "fit·crop·focus·zoom으로 잘라 낼 영역을 노란 점선으로 미리 그린다. 결과에 원본 픽셀",
            new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["path"] = new JsonObject { ["type"] = "string", ["description"] = "그림 파일 절대 경로" },
                    ["width"] = new JsonObject { ["type"] = "number", ["description"] = "넣을 칸 너비(pt, 비율만 쓴다)" },
                    ["height"] = new JsonObject { ["type"] = "number", ["description"] = "넣을 칸 높이(pt)" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["crop"] = cropSchema.DeepClone(),
                    ["keep"] = keepSchema.DeepClone(),
                    ["focus"] = focusSchema.DeepClone(),
                    ["zoom"] = new JsonObject { ["type"] = "number", ["description"] = ZoomDescription },
                },
                ["required"] = new JsonArray("path"),
            }, ImagePreviewTool);
        Add("replace_pictures_in_box",
            "칸 안(중심 기준)의 그림과 「[그림…]」 글상자를 모두 지우고 새 그림 목록을 넣는다(지우기·넣기 한 번에, 첫 그림의 Z순서 자리). " +
            "나누기·격자·crop/keep/focus 인자는 add_pictures_in_box와 같다. BckPPT가 넣은 뒤 사용자가 바꾼 그림이 칸에 있으면 멈추고 알린다 " +
            "(skip_user_modified=true면 칸을 건너뛰고, force=true면 덮어쓴다) (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["paths"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "새 그림 파일 절대 경로들" },
                    ["box"] = new JsonObject { ["description"] = "칸: {left,top,width,height}(pt) 또는 칸 도형 주소(그 도형은 지우지 않는다)" },
                    ["gap"] = new JsonObject { ["type"] = "number", ["description"] = "그림 사이 간격 pt(기본 3)" },
                    ["direction"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("auto", "horizontal", "vertical") },
                    ["rows"] = new JsonObject { ["type"] = "integer" }, ["columns"] = new JsonObject { ["type"] = "integer" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["crops"] = new JsonObject { ["type"] = "array", ["items"] = cropSchema.DeepClone() },
                    ["keeps"] = new JsonObject { ["type"] = "array", ["items"] = keepSchema.DeepClone() },
                    ["focuses"] = new JsonObject { ["type"] = "array", ["items"] = focusSchema.DeepClone() },
                    ["zoom"] = new JsonObject { ["type"] = "number", ["description"] = ZoomDescription },
                    ["zooms"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "number" } },
                    ["names"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" } },
                    ["sources"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" } },
                    ["source"] = Str(SourceDescription), ["license"] = Str(LicenseDescription),
                    ["remove_placeholders"] = new JsonObject { ["type"] = "boolean", ["description"] = "칸 안 「[그림…]」 글상자도 지운다(기본 true)" },
                    ["skip_user_modified"] = new JsonObject { ["type"] = "boolean", ["description"] = "사용자가 바꾼 그림이 있으면 칸을 건너뛴다" },
                    ["force"] = new JsonObject { ["type"] = "boolean", ["description"] = "사용자가 바꾼 그림도 지운다(사용자 확인 뒤)" },
                },
            }, "slide", "paths", "box"), ReplacePicturesInBox);
        Add("apply_picture_plan",
            "그림 계획을 칸별로 적용하고 결과 표와 바뀐 쪽 축소판을 돌려준다. 줄 형식 「쪽|칸|경로[@crop=l,t,r,b|@keep=x0,y0,x1,y1|@focus=x,y[,zoom]];경로…」 " +
            "(# 줄은 주석). 칸은 cells에 정의한 이름(상자 또는 도형 주소) 또는 그 쪽의 도형 이름. 칸마다 replace_pictures_in_box처럼 칸 안 그림·「[그림…]」 글상자를 " +
            "지우고 1장은 칸 전체, 여러 장은 나란히 넣는다. 사용자가 바꾼 그림이 있는 칸은 기본으로 건너뛴다. 상대 경로는 base_dir(기본: 계획 파일 폴더, 없으면 그 위 폴더) 기준 (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["plan_path"] = Str("계획 파일(UTF-8 텍스트) 절대 경로"),
                    ["plan"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "plan_path 대신 줄 배열" },
                    ["base_dir"] = Str("상대 경로의 기준 폴더"),
                    ["cells"] = new JsonObject
                    {
                        ["type"] = "object",
                        ["description"] = "칸 이름 → {left,top,width,height} 또는 도형 주소. 예: {\"AS-IS\":{\"left\":49.68,\"top\":220.32,\"width\":281.52,\"height\":93.6}}",
                    },
                    ["slides"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "integer" }, ["description"] = "이 쪽들의 줄만 적용" },
                    ["fit"] = FitSchema("기본 cover"),
                    ["gap"] = new JsonObject { ["type"] = "number", ["description"] = "그림 사이 간격 pt(기본 3)" },
                    ["skip_user_modified"] = new JsonObject { ["type"] = "boolean", ["description"] = "기본 true: 사용자가 바꾼 그림이 있는 칸은 건너뜀" },
                    ["force"] = new JsonObject { ["type"] = "boolean", ["description"] = "사용자가 바꾼 그림도 덮어쓴다(사용자 확인 뒤)" },
                    ["dry_run"] = new JsonObject { ["type"] = "boolean", ["description"] = "true면 고치지 않고 할 일만 표로" },
                    ["render"] = new JsonObject { ["type"] = "boolean", ["description"] = "바뀐 쪽 축소판을 돌려준다(기본 true)" },
                },
            }), ApplyPicturePlan);
        Add("image_sheet",
            "후보 그림 파일 여러 장을 번호·파일 이름·픽셀·비율·용량과 함께 한 장 축소판으로 돌려준다(후보 비교용, 문서와 상관없음). 최대 " + MaxSheetImages + "장",
            new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["paths"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "string" }, ["description"] = "그림 파일 절대 경로들" },
                    ["columns"] = new JsonObject { ["type"] = "integer", ["description"] = "한 줄의 장 수(기본 4)" },
                },
                ["required"] = new JsonArray("paths"),
            }, ImageSheet);
        Add("render_region",
            "쪽의 한 부분(box 또는 도형)만 크게 렌더한다. 칸 안 사진을 자세히 볼 때. scale은 pt당 픽셀(비우면 영역이 약 1200px 너비)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slide"] = new JsonObject { ["type"] = "integer", ["description"] = "쪽 번호(1부터)" },
                    ["box"] = new JsonObject { ["description"] = "{left,top,width,height}(pt) 또는 도형 주소" },
                    ["shape"] = Str("box 대신 도형 주소"),
                    ["scale"] = new JsonObject { ["type"] = "number", ["description"] = "pt당 픽셀(0.5~10)" },
                    ["margin"] = new JsonObject { ["type"] = "number", ["description"] = "영역 둘레 여백 pt(기본 6)" },
                },
            }, "slide"), RenderRegion);
        Add("document_size",
            "문서 파일 크기와 그림별 크기(디스크에 저장된 판): 그림 파일 합계, 쪽별 그림 크기, 큰 그림 순 목록(픽셀·자른 부분 포함 여부·같이 쓰는 곳)",
            WithFile(Schema(("top", "integer", "큰 그림 몇 개를 보일지(기본 15)"))), DocumentSizeTool);
        Add("compress_pictures",
            "그림을 줄여 문서 크기를 줄인다: 쪽에 보이는 크기 기준 dpi로 줄이고, delete_cropped면 잘린 부분을 버린다(그 뒤에는 자르기를 넓힐 수 없다). " +
            "confirm 없이 부르면 고치지 않고 줄일 그림과 예상 크기만 보여 준다. 사용자 문서는 사용자 확인 뒤에만 confirm=true (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["dpi"] = new JsonObject { ["type"] = "number", ["description"] = "목표 해상도(기본 200, 72~600)" },
                    ["delete_cropped"] = new JsonObject { ["type"] = "boolean", ["description"] = "잘린 부분 버리기(기본 true)" },
                    ["slides"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "integer" }, ["description"] = "이 쪽들만(비우면 전체)" },
                    ["jpeg_quality"] = new JsonObject { ["type"] = "integer", ["description"] = "JPEG 품질 50~100(기본 85). 투명한 그림은 PNG" },
                    ["confirm"] = new JsonObject { ["type"] = "boolean", ["description"] = "true면 적용(사용자 확인 뒤)" },
                },
            }), CompressPictures);
        Add("export_picture_credits",
            "쪽별 그림 출처 목록(마크다운). add_picture 등의 source·license로 남긴 기록을 읽는다(없으면 「출처 미기록」과 원본 파일 이름). " +
            "output_path에 저장하거나, write_notes=true면 쪽마다 발표자 메모에 「그림 출처: …」 문단을 쓴다(그때만 문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slides"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "integer" } },
                    ["output_path"] = Str("마크다운을 저장할 절대 경로"),
                    ["write_notes"] = new JsonObject { ["type"] = "boolean", ["description"] = "발표자 메모에 출처 문단 쓰기" },
                },
            }), ExportPictureCredits);
        Add("import_slides",
            "다른 PPT 파일의 쪽을 이 문서 after 쪽 뒤에 붙인다(클립보드를 쓰지 않는다). 붙인 쪽은 이 문서의 디자인(테마)을 따른다. " +
            "source가 창에 열려 있어도 디스크에 저장된 판을 읽으니, 고친 내용이 있으면 사용자에게 먼저 저장해 달라고 한다 (문서를 고친다)",
            WithFile(Schema(
                ("source", "string", "가져올 PPT. 절대 경로나 열린 문서 이름 조각"),
                ("from", "integer", "가져올 첫 쪽(1부터)"),
                ("to", "integer", "가져올 끝 쪽. 비우면 from 한 쪽"),
                ("after", "integer", "이 문서에서 이 쪽 뒤에 붙인다(0이면 맨 앞)")), "source", "from", "after"), ImportSlides);
        Add("delete_slides",
            "쪽들을 지운다. 지운 쪽의 번호·제목을 돌려준다. 지우기 전에 list_slides로 번호를 확인하고 사용자에게 어느 쪽을 지울지 확인받는다. " +
            "창에 열리지 않은 파일은 바로 저장되므로 confirm=true가 있어야 지운다 (문서를 고친다)",
            WithFile(new JsonObject
            {
                ["type"] = "object",
                ["properties"] = new JsonObject
                {
                    ["slides"] = new JsonObject { ["type"] = "array", ["items"] = new JsonObject { ["type"] = "integer" }, ["description"] = "지울 쪽 번호들(1부터, 지우기 전 번호)" },
                    ["confirm"] = new JsonObject { ["type"] = "boolean", ["description"] = "창에 열리지 않은 파일을 지울 때 사용자 확인을 받았으면 true" },
                },
            }, "slides"), DeleteSlides);
        Add("move_slide", "쪽을 옮긴다. to는 옮긴 뒤 그 쪽의 번호 (문서를 고친다)",
            WithFile(Schema(("slide", "integer", "옮길 쪽(1부터)"), ("to", "integer", "옮긴 뒤 번호")), "slide", "to"), MoveSlide);

        foreach (var tool in _agentTools.Tools)
        {
            var schema = JsonNode.Parse(tool.ParametersSchema)!.AsObject();
            var name = tool.Name;
            _tools[name] = new McpTool(name, tool.Description + (tool.Writes ? " (문서를 고친다)" : ""),
                WithFile(schema), (args, ct) => CallAgentTool(name, tool.Writes, args, ct));
        }
    }

    public IReadOnlyCollection<McpTool> Tools => _tools.Values;

    /// <summary>도구를 부르고 진단 로그에 호출(도구·파일·시간·오류)과 그 뒤 PowerPoint 상태를 남긴다.</summary>
    public async Task<McpToolResult> CallAsync(string name, JsonElement arguments, CancellationToken cancellationToken)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        string file = arguments.ValueKind == JsonValueKind.Object ? ToolArgs.OptionalString(arguments, "file") ?? "(보고 있는 문서)" : "";
        McpToolResult? result = null;
        try { return result = await CallCoreAsync(name, arguments, cancellationToken); }
        finally
        {
            string outcome = result is null ? "취소" : result.IsError ? "오류: " + Truncate(result.Content.FirstOrDefault()?.Text) : "ok";
            PptDiagnosticLog.Write("mcp", $"{name} file={file} {watch.ElapsedMilliseconds}ms {outcome}");
            string state;
            try { state = await Task.Run(_ppt.DescribeState); }
            catch (Exception ex) { state = "상태 읽기 실패: " + ex.GetBaseException().Message; }
            PptDiagnosticLog.Write("state", state);
        }

        static string Truncate(string? text) => text is null ? "" : text.Length <= 200 ? text : text[..200] + "…";
    }

    private async Task<McpToolResult> CallCoreAsync(string name, JsonElement arguments, CancellationToken cancellationToken)
    {
        if (!_tools.TryGetValue(name, out var tool))
            return McpToolResult.Error($"'{name}' 도구는 없습니다.");
        if (arguments.ValueKind != JsonValueKind.Object)
            return McpToolResult.Error("인자는 JSON 객체여야 합니다.");
        try { return await tool.Call(arguments, cancellationToken); }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) when (ex.GetBaseException() is PptUserInterventionException)
        {
            return McpToolResult.Error("사용자가 PowerPoint에서 도형이나 글을 선택해 작업을 멈췄습니다. 사용자에게 확인한 뒤 다시 부르세요.");
        }
        catch (AgentToolException ex) { return McpToolResult.Error(ex.Message); }
        catch (Exception ex) { return McpToolResult.Error($"{name} 실행 실패: {ex.GetBaseException().Message}"); }
    }

    // ── 대상 문서 ────────────────────────────────────────────

    /// <summary>file 인자를 문서 경로로 바꾼다. 비우면 보고 있는 문서, 경로면 그 파일, 아니면 열린 문서 이름 조각.</summary>
    public string ResolveFile(JsonElement args) =>
        ResolveFile(ToolArgs.OptionalString(args, "file"), ToolArgs.OptionalBool(args, "show_window") == true);

    /// <summary>open_presentation 직후에는 문서가 아직 안 보일 수 있어 잠깐 다시 찾는다.</summary>
    private DateTime _lastOpen = DateTime.MinValue;
    public static TimeSpan RecentOpenWindow { get; set; } = TimeSpan.FromSeconds(120);
    public static TimeSpan RetryAfterOpen { get; set; } = TimeSpan.FromSeconds(10);

    public string ResolveFile(string? file, bool showWindow = false)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        while (true)
        {
            try
            {
                var path = ResolveFileOnce(file);
                // 경로로 준 문서가 창 없이 남아 있으면 창을 띄워 달라고 했을 때 띄운다.
                if (showWindow && !_ppt.IsOpenInWindow(path) && Windowless(path).Count == 1) ShowWindow(path);
                return path;
            }
            catch (AgentToolException) when (DateTime.UtcNow - _lastOpen < RecentOpenWindow && watch.Elapsed < RetryAfterOpen)
            {
                Thread.Sleep(500);
            }
            catch (AgentToolException ex) when (ProtectedViewHint(file) is { } hint)
            {
                throw new AgentToolException(ex.Message + " " + hint);
            }
            catch (AgentToolException ex) when (Windowless(file) is { Count: > 0 } hidden)
            {
                if (showWindow && hidden.Count == 1)
                {
                    ShowWindow(hidden[0].Path);
                    return hidden[0].Path;
                }
                throw new AgentToolException(ex.Message + $" 창 없이 PowerPoint에 남아 있는 문서가 있습니다: {string.Join(", ", hidden.Select(h => h.Path))}. " +
                                             "창을 닫았는데 문서가 메모리에 남은 것입니다(저장 안 된 고침이 있을 수 있음). " +
                                             "open_presentation(path)로 창을 다시 띄우거나 이 도구에 show_window=true를 주세요.");
            }
        }
    }

    /// <summary>창 없이 남은 문서 중 file(비우면 전부, 경로 또는 이름 조각)에 맞는 것.</summary>
    private IReadOnlyList<OpenPresentationInfo> Windowless(string? file)
    {
        IReadOnlyList<OpenPresentationInfo> all;
        try { all = _ppt.ListWindowlessPresentations(); } catch { return []; }
        file = file?.Trim().Trim('"');
        if (string.IsNullOrEmpty(file)) return all;
        return all.Where(o => string.Equals(o.Path, file, StringComparison.OrdinalIgnoreCase) ||
                              !Path.IsPathFullyQualified(file) && Path.GetFileName(o.Path).Contains(file, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    private void ShowWindow(string path)
    {
        PptDiagnosticLog.Event("show-window", new { path });
        _lastOpen = DateTime.UtcNow;
        var r = _ppt.OpenAndWait(path, unblock: false, TimeSpan.FromSeconds(30));
        if (r.State is OpenedState.NotYet or OpenedState.ProtectedView)
            throw new AgentToolException($"창 없이 남은 문서 {Path.GetFileName(path)}에 창을 띄우지 못했습니다. open_presentation으로 다시 해 보세요.");
    }

    private string? ProtectedViewHint(string? file)
    {
        IReadOnlyList<ProtectedViewInfo> views;
        try { views = _ppt.ListProtectedViews(); } catch { return null; }
        file = file?.Trim().Trim('"');
        var hits = views.Where(v => string.IsNullOrEmpty(file) || v.Name.Contains(Path.GetFileName(file), StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(v.Path, file, StringComparison.OrdinalIgnoreCase)).ToList();
        return hits.Count == 0 ? null
            : $"「제한된 보기」로 열린 창이 있습니다({string.Join(", ", hits.Select(v => v.Name))}). 이 창의 문서는 도구가 읽지 못합니다. " +
              "사용자에게 확인받고 open_presentation(path, unblock=true)로 편집 모드로 다시 여세요.";
    }

    private string ResolveFileOnce(string? file)
    {
        file = file?.Trim().Trim('"');
        if (string.IsNullOrEmpty(file))
            return _ppt.GetActivePresentation()?.Path is { Length: > 0 } active
                ? active
                : throw new AgentToolException("PowerPoint에서 보고 있는 문서가 없습니다. file에 PPT 경로를 주거나 open_presentation으로 여세요.");

        if (Path.IsPathFullyQualified(file))
        {
            var full = Path.GetFullPath(file);
            if (!PresentationExtensions.Contains(Path.GetExtension(full), StringComparer.OrdinalIgnoreCase))
                throw new AgentToolException("PPT 파일(.pptx·.pptm·.ppt)이 아닙니다.");
            return File.Exists(full) ? full : throw new AgentToolException($"파일이 없습니다: {full}");
        }

        var open = _ppt.ListOpenWindows().Where(o => o.IsSaved).ToList();
        var hits = open.Where(o => Path.GetFileName(o.Path).Contains(file, StringComparison.OrdinalIgnoreCase)).ToList();
        var exact = hits.Where(o => string.Equals(Path.GetFileNameWithoutExtension(o.Path), file, StringComparison.OrdinalIgnoreCase)).ToList();
        if (exact.Count == 1) return exact[0].Path;
        return hits.Count switch
        {
            1 => hits[0].Path,
            0 => throw new AgentToolException($"'{file}' 이름이 든 열린 문서가 없습니다. 열린 문서: " +
                                              (open.Count == 0 ? "없음" : string.Join(", ", open.Select(o => o.Name)))),
            _ => throw new AgentToolException($"'{file}' 이름이 든 문서가 여럿입니다: {string.Join(", ", hits.Select(o => o.Name))}. 더 구체적으로 주세요.")
        };
    }

    private async Task<McpToolResult> CallAgentTool(string name, bool writes, JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        if (!writes)
        {
            var read = await _agentTools.ExecuteAsync(new AgentToolContext(_ppt, path),
                new LlmToolCall(Guid.NewGuid().ToString("N"), name, args.GetRawText()), ct);
            return new([McpContent.FromText(read.Content)], read.IsError);
        }
        return await WriteAsync(path, async () =>
        {
            var outcome = await _agentTools.ExecuteAsync(new AgentToolContext(_ppt, path),
                new LlmToolCall(Guid.NewGuid().ToString("N"), name, args.GetRawText()), ct);
            return new McpToolResult([McpContent.FromText(outcome.Content)], outcome.IsError);
        });
    }

    /// <summary>
    /// 쓰기 도구 공통. 창에 열린 문서는 사용자가 보고 있으니 저장하지 않고 고친다(그동안 사용자가 도형을 고르면 멈춘다).
    /// 읽기 전용으로 열린 창이면 고친 내용을 그 이름으로 저장할 수 없으니 결과에 알린다.
    /// CellBox 동기화 폴더의 큰 PPT가 읽기 전용으로 열려 고친 내용을 저장하지 못한 일이 있었다(2026-11-03).
    /// </summary>
    private async Task<McpToolResult> WriteAsync(string path, Func<Task<McpToolResult>> write)
    {
        bool inWindow = _ppt.IsOpenInWindow(path);
        using var watch = inWindow ? _ppt.WatchUserActivity(path) : null;
        var result = await write();
        if (result.IsError || !inWindow || !IsReadOnlyWindow(path)) return result;
        return result with
        {
            Content = [.. result.Content, McpContent.FromText(
                "주의: 이 문서는 PowerPoint에 읽기 전용으로 열려 있습니다. 고친 내용은 창에 보이지만 같은 이름으로 저장되지 않으니, " +
                "사용자에게 '다른 이름으로 저장'하거나 파일을 쓰기 가능하게 다시 열어 달라고 알리세요.")],
        };
    }

    private bool IsReadOnlyWindow(string path) =>
        _ppt.ListOpenWindows().Any(o => o.IsReadOnly && string.Equals(o.Path, path, StringComparison.OrdinalIgnoreCase));

    // ── MCP 전용 도구 ────────────────────────────────────────

    private McpToolResult ListPresentations()
    {
        var open = _ppt.ListOpenWindows();
        var protectedViews = _ppt.ListProtectedViews();
        var hidden = Windowless(null);
        if (open.Count == 0 && protectedViews.Count == 0 && hidden.Count == 0) return McpToolResult.Ok("PowerPoint에 열린 문서가 없습니다.");
        var sb = new StringBuilder($"열린 문서 {open.Count}개\n");
        foreach (var h in hidden)
            sb.AppendLine($"  {h.Name} · {h.SlideCount}쪽 · 창 없음(창을 닫았는데 메모리에 남음) — open_presentation(path)로 창 띄우기 · {h.Path}");
        foreach (var pv in protectedViews)
            sb.AppendLine($"  {pv.Name} · 「제한된 보기」 — 도구가 읽지 못함. open_presentation(path, unblock=true)로 편집 모드로 다시 연다" +
                          (pv.Path.Length > 0 ? $" · {pv.Path}" : ""));
        foreach (var o in open)
            sb.AppendLine($"{(o.IsActive ? "▶" : " ")} {o.Name} · {o.SlideCount}쪽 · {(o.IsSaved ? o.Path : "저장 안 된 새 문서")}" +
                          (o.IsReadOnly ? " · 읽기 전용(고쳐도 같은 이름으로 저장 안 됨)" : ""));
        if (_ppt.GetActivePresentation() is { CurrentSlide: int current } active)
            sb.AppendLine($"보고 있는 쪽: {active.Name} {current}쪽");
        return McpToolResult.Ok(sb.ToString().TrimEnd());
    }

    /// <summary>
    /// 연 뒤 문서가 올라올 때까지 기다린다. 26MB PPT를 연 직후 다음 도구가 「열린 문서: 없음」을 받았고(셸로 열면 바로 돌아온다),
    /// 메일로 받은 파일은 「제한된 보기」로 열려 다른 도구가 문서를 찾지 못했다(2026-11-17).
    /// </summary>
    private async Task<McpToolResult> OpenPresentation(JsonElement args, CancellationToken ct)
    {
        var raw = ToolArgs.OptionalString(args, "path")?.Trim().Trim('"') ?? throw new AgentToolException("인자 'path'가 필요합니다.");
        if (!Path.IsPathFullyQualified(raw)) throw new AgentToolException("path는 PPT 절대 경로여야 합니다.");
        var path = ResolveFileOnce(raw);
        bool unblock = ToolArgs.OptionalBool(args, "unblock") == true;
        var timeout = TimeSpan.FromSeconds(Math.Clamp(ToolArgs.OptionalInt(args, "wait_seconds") ?? 60, 1, 300));
        _lastOpen = DateTime.UtcNow;
        var r = await Task.Run(() => _ppt.OpenAndWait(path, unblock, timeout, ct), ct);
        _lastOpen = DateTime.UtcNow;
        string mark = r.Unblocked ? " 인터넷에서 받은 파일 표시를 지웠습니다." : "";
        return r.State switch
        {
            OpenedState.Editable => McpToolResult.Ok($"열었습니다: {path} · {r.SlideCount}쪽 · 편집 가능 ({r.Waited.TotalSeconds:0.#}초).{mark}"),
            OpenedState.ReadOnly => McpToolResult.Ok($"열었습니다: {path} · {r.SlideCount}쪽 · 읽기 전용 ({r.Waited.TotalSeconds:0.#}초).{mark} " +
                                                     "고친 내용을 같은 이름으로 저장할 수 없으니, 고칠 일이 있으면 사용자에게 다른 곳에서 열었는지 확인하세요."),
            OpenedState.ProtectedView => McpToolResult.Ok(
                $"「제한된 보기」로 열렸습니다: {path}" + (r.HadDownloadMark ? " (메일·인터넷에서 받은 파일 표시가 있음)" : "") + ". " +
                "이 상태에서는 다른 도구가 이 문서를 찾지 못합니다. 사용자에게 확인받은 뒤 unblock=true로 다시 부르면 표시를 지우고 편집 모드로 다시 엽니다."),
            _ => McpToolResult.Ok($"여는 중입니다: {path} — {r.Waited.TotalSeconds:0}초 안에 문서가 올라오지 않았습니다. " +
                                  "잠시 뒤 list_presentations로 확인하세요(PowerPoint에 대화상자가 떠 있을 수도 있습니다)."),
        };
    }

    private Task<McpToolResult> GoToSlide(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        _ppt.NavigateToSlide(path, slide);
        return Task.FromResult(McpToolResult.Ok($"{Path.GetFileName(path)} {slide}쪽으로 옮겼습니다."));
    }

    private async Task<McpToolResult> ListFonts(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        var fontName = ToolArgs.OptionalString(args, "font_name")?.Trim();
        // 전체 분석(도형마다 줄 측정)은 큰 문서에서 COM 45초 제한을 넘는다. 글꼴만 쪽마다 따로 읽는다.
        var fonts = await Task.Run(() =>
        {
            using var keep = _ppt.IsOpenInWindow(path) ? null : _ppt.KeepOpen(path, readOnly: true);
            var all = new List<PptFontInfo>();
            int count = _ppt.ListSlides(path).Count;
            for (int slide = 1; slide <= count; slide++)
            {
                ct.ThrowIfCancellationRequested();
                all.AddRange(_ppt.ListSlideFonts(path, slide));
            }
            return all;
        }, ct);
        var sb = new StringBuilder();
        if (string.IsNullOrEmpty(fontName))
        {
            var groups = fonts.GroupBy(f => f.FontName).OrderByDescending(g => g.Count()).ToList();
            sb.AppendLine($"글꼴 {groups.Count}종");
            foreach (var g in groups)
            {
                var slides = g.Select(f => f.SlideIndex).Distinct().Order().ToList();
                sb.AppendLine($"- {g.Key}: {g.Count()}곳 · {slides.Count}쪽 ({string.Join(",", slides.Take(30))}{(slides.Count > 30 ? ",…" : "")})" +
                              (g.Any(f => !f.IsAllowed) ? " · 허용 목록 밖" : ""));
            }
            return McpToolResult.Ok(sb.ToString().TrimEnd());
        }

        var hits = fonts.Where(f => string.Equals(f.FontName, fontName, StringComparison.OrdinalIgnoreCase)).ToList();
        if (hits.Count == 0)
            return McpToolResult.Ok($"'{fontName}' 글꼴을 쓰는 곳이 없습니다. 쓰인 글꼴: {string.Join(", ", fonts.Select(f => f.FontName).Distinct())}");
        sb.AppendLine($"'{fontName}' {hits.Count}곳");
        foreach (var f in hits.OrderBy(f => f.SlideIndex))
            sb.AppendLine($"- {f.SlideIndex}쪽 [{f.ShapeName}] {f.FontSize}pt{(f.IsBold ? " 굵게" : "")} · {f.TextPreview}");
        return McpToolResult.Ok(sb.ToString().TrimEnd());
    }

    private async Task<McpToolResult> ExportMarkdown(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        var output = ToolArgs.OptionalString(args, "output_path") is { Length: > 0 } o ? o : Path.ChangeExtension(path, ".md");
        if (!Path.IsPathFullyQualified(output)) throw new AgentToolException("output_path는 절대 경로여야 합니다.");
        var result = await Task.Run(() => _ppt.ExportMarkdown(path, output, ct), ct);
        return McpToolResult.Ok($"{result.SlideCount}쪽을 내보냈습니다: {result.OutputPath}" +
                                (result.WarningCount > 0 ? $" (경고 {result.WarningCount}건)" : ""));
    }

    private Task<McpToolResult> ImportSlides(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        var source = ResolveFile(ToolArgs.OptionalString(args, "source") ?? throw new AgentToolException("인자 'source'가 필요합니다."));
        if (string.Equals(source, path, StringComparison.OrdinalIgnoreCase))
            throw new AgentToolException("같은 문서의 쪽은 duplicate_slide로 복제하세요.");
        int from = ToolArgs.Int(args, "from");
        int to = ToolArgs.OptionalInt(args, "to") ?? from;
        int after = ToolArgs.Int(args, "after");
        if (from < 1 || to < from) throw new AgentToolException("from은 1 이상, to는 from 이상이어야 합니다.");
        int sourceCount = _ppt.ListSlides(source).Count;
        if (to > sourceCount) throw new AgentToolException($"{Path.GetFileName(source)}는 {sourceCount}쪽까지 있습니다.");
        return WriteAsync(path, async () =>
        {
            var (first, count) = await Task.Run(() => _ppt.ImportSlides(source, from, to, path, after), ct);
            return McpToolResult.Ok($"{Path.GetFileName(source)} {from}~{to}쪽을 {first}~{first + count - 1}쪽으로 붙였습니다. " +
                                    $"뒤 쪽 번호는 {count}씩 밀렸습니다. 디자인은 이 문서 것을 따르니 render_slide로 확인하세요.");
        });
    }

    private Task<McpToolResult> DeleteSlides(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        var slides = args.TryGetProperty("slides", out var s) && s.ValueKind == JsonValueKind.Array
            ? s.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Number).Select(e => e.GetInt32()).ToList()
            : throw new AgentToolException("인자 'slides'(쪽 번호 배열)가 필요합니다.");
        if (!_ppt.IsOpenInWindow(path) && ToolArgs.OptionalBool(args, "confirm") != true)
            throw new AgentToolException("창에 열리지 않은 파일은 지우는 즉시 저장됩니다. 사용자에게 지울 쪽을 확인받은 뒤 confirm=true로 다시 부르세요.");
        return WriteAsync(path, async () =>
        {
            var deleted = await Task.Run(() => _ppt.DeleteSlides(path, slides), ct);
            return McpToolResult.Ok($"{deleted.Count}쪽을 지웠습니다(지우기 전 번호): " +
                                    string.Join(", ", deleted.Select(d => $"{d.Index}쪽 「{(d.Title.Length > 0 ? d.Title : "제목 없음")}」")));
        });
    }

    private Task<McpToolResult> MoveSlide(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide"), to = ToolArgs.Int(args, "to");
        return WriteAsync(path, async () =>
        {
            await Task.Run(() => _ppt.MoveSlide(path, slide, to), ct);
            return McpToolResult.Ok($"{slide}쪽을 {to}쪽 자리로 옮겼습니다.");
        });
    }

    private async Task<McpToolResult> RenderSlide(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        var png = await _ppt.GetSlideThumbnailAsync(path, slide) ?? throw new AgentToolException($"{slide}쪽을 렌더하지 못했습니다.");
        try
        {
            return new([McpContent.FromText($"{Path.GetFileName(path)} {slide}쪽"), McpContent.FromImage(await File.ReadAllBytesAsync(png, ct))], false);
        }
        finally { try { File.Delete(png); } catch { } }
    }

    // ── 여러 쪽 렌더 ─────────────────────────────────────────

    public const int MaxSeparateRenders = 12, MaxSheetRenders = 60;

    private async Task<McpToolResult> RenderSlides(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        List<int> slides;
        if (args.TryGetProperty("slides", out var s) && s.ValueKind == JsonValueKind.Array)
            slides = s.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Number).Select(e => e.GetInt32()).Distinct().ToList();
        else if (ToolArgs.OptionalInt(args, "from") is int from)
            slides = Enumerable.Range(from, Math.Max(1, (ToolArgs.OptionalInt(args, "to") ?? from) - from + 1)).ToList();
        else throw new AgentToolException("slides(쪽 번호 배열)나 from~to를 주세요.");
        if (slides.Count == 0) throw new AgentToolException("렌더할 쪽이 없습니다.");
        int count = _ppt.ListSlides(path).Count;
        if (slides.Where(n => n < 1 || n > count).ToList() is { Count: > 0 } bad)
            throw new AgentToolException($"{Path.GetFileName(path)}는 1~{count}쪽입니다({string.Join(",", bad)}쪽 없음).");
        bool sheet = ToolArgs.OptionalBool(args, "sheet") ?? slides.Count > MaxSeparateRenders;
        if (slides.Count > (sheet ? MaxSheetRenders : MaxSeparateRenders))
            throw new AgentToolException(sheet ? $"축소판은 {MaxSheetRenders}쪽까지입니다." : $"따로 받는 그림은 {MaxSeparateRenders}쪽까지입니다. sheet=true로 모으세요.");

        var pngs = new List<(int Slide, string? Png)>();
        try
        {
            using (var keep = _ppt.IsOpenInWindow(path) ? null : _ppt.KeepOpen(path, readOnly: true))
                foreach (var n in slides)
                {
                    ct.ThrowIfCancellationRequested();
                    pngs.Add((n, await _ppt.GetSlideThumbnailAsync(path, n)));
                }
            var failed = pngs.Where(p => p.Png is null).Select(p => p.Slide).ToList();
            var ok = pngs.Where(p => p.Png is not null).ToList();
            if (ok.Count == 0) throw new AgentToolException("렌더하지 못했습니다.");
            string note = failed.Count > 0 ? $" (렌더 실패: {string.Join(",", failed)}쪽)" : "";
            if (sheet)
            {
                int columns = Math.Clamp(ToolArgs.OptionalInt(args, "columns") ?? 4, 1, 8);
                var bytes = ContactSheet.Compose(ok.Select(p => (p.Slide, p.Png!)).ToList(), columns);
                return new([McpContent.FromText($"{Path.GetFileName(path)} {string.Join(",", ok.Select(p => p.Slide))}쪽 축소판{note}"), McpContent.FromImage(bytes)], false);
            }
            var content = new List<McpContent> { McpContent.FromText($"{Path.GetFileName(path)} {ok.Count}쪽{note}") };
            foreach (var (n, png) in ok)
            {
                content.Add(McpContent.FromText($"{n}쪽"));
                content.Add(McpContent.FromImage(await File.ReadAllBytesAsync(png!, ct)));
            }
            return new(content, false);
        }
        finally
        {
            foreach (var (_, png) in pngs) if (png is not null) try { File.Delete(png); } catch { }
        }
    }

    // ── 그림 ────────────────────────────────────────────────

    private static readonly string[] ImageExtensions = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg"];

    private static string ImagePath(string? path)
    {
        path = path?.Trim().Trim('"');
        if (string.IsNullOrEmpty(path)) throw new AgentToolException("그림 파일 경로가 필요합니다.");
        if (!Path.IsPathFullyQualified(path)) throw new AgentToolException($"그림 경로는 절대 경로여야 합니다: {path}");
        var full = Path.GetFullPath(path);
        if (!ImageExtensions.Contains(Path.GetExtension(full), StringComparer.OrdinalIgnoreCase))
            throw new AgentToolException($"그림 파일이 아닙니다({string.Join("·", ImageExtensions)}): {full}");
        return File.Exists(full) ? full : throw new AgentToolException($"그림 파일이 없습니다: {full}");
    }

    private static PictureFit ParseFit(string? fit) => fit?.Trim().ToLowerInvariant() switch
    {
        null or "" or "cover" => PictureFit.Cover,
        "contain" => PictureFit.Contain,
        "stretch" => PictureFit.Stretch,
        _ => throw new AgentToolException("fit은 cover·contain·stretch 중 하나입니다."),
    };

    private static CropFractions CheckCrop(CropFractions c, string what = "crop")
    {
        if (new[] { c.Left, c.Top, c.Right, c.Bottom }.Any(v => v is < 0 or >= 1))
            throw new AgentToolException($"{what} 값은 0~1 비율입니다(원본 기준, 0.1 = 10%).");
        if (c.Left + c.Right >= 0.99f || c.Top + c.Bottom >= 0.99f)
            throw new AgentToolException($"{what}이 그림을 거의 다 잘라 냅니다.");
        return c;
    }

    private static CropFractions? ParseCrop(JsonElement crop)
    {
        if (crop.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined) return null;
        if (crop.ValueKind != JsonValueKind.Object) throw new AgentToolException("crop은 {left,top,right,bottom}(0~1) 객체입니다.");
        float F(string n) => ToolArgs.OptionalFloat(crop, n) ?? 0f;
        return CheckCrop(new CropFractions(F("left"), F("top"), F("right"), F("bottom")));
    }

    /// <summary>남길 영역 {x0,y0,x1,y1}(원본 기준 0~1)을 잘라 낼 비율로.</summary>
    private static CropFractions? ParseKeep(JsonElement keep)
    {
        if (keep.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined) return null;
        if (keep.ValueKind != JsonValueKind.Object) throw new AgentToolException("keep은 남길 영역 {x0,y0,x1,y1}(원본 기준 0~1) 객체입니다.");
        float F(string n, float d) => ToolArgs.OptionalFloat(keep, n) ?? d;
        float x0 = F("x0", 0), y0 = F("y0", 0), x1 = F("x1", 1), y1 = F("y1", 1);
        if (x0 < 0 || y0 < 0 || x1 > 1 || y1 > 1 || x1 <= x0 || y1 <= y0)
            throw new AgentToolException("keep은 0 ≤ x0 < x1 ≤ 1, 0 ≤ y0 < y1 ≤ 1 입니다(남길 영역, 원본 기준).");
        return CheckCrop(PictureFitter.FromKeep(x0, y0, x1, y1), "keep");
    }

    /// <summary>crop(잘라 낼 양)이나 keep(남길 영역) 중 하나.</summary>
    private static CropFractions? ParseCropOrKeep(JsonElement args, string cropName = "crop", string keepName = "keep")
    {
        var crop = args.TryGetProperty(cropName, out var c) ? ParseCrop(c) : null;
        var keep = args.TryGetProperty(keepName, out var k) ? ParseKeep(k) : null;
        if (crop is not null && keep is not null) throw new AgentToolException($"{cropName}과 {keepName}은 둘 중 하나만 줍니다.");
        return crop ?? keep;
    }

    private static CropFractions? ParseCrop(JsonElement args, string name) => ParseCropOrKeep(args, name, "keep");

    /// <summary>배열 인자(crops·keeps)를 그림마다 하나로 합친다.</summary>
    private static List<CropFractions?> ParseCropArrays(JsonElement args, int count)
    {
        List<CropFractions?> Read(string name, Func<JsonElement, CropFractions?> parse) =>
            args.TryGetProperty(name, out var a) && a.ValueKind == JsonValueKind.Array ? a.EnumerateArray().Select(parse).ToList() : [];
        var crops = Read("crops", ParseCrop);
        var keeps = Read("keeps", ParseKeep);
        var result = new List<CropFractions?>(count);
        for (int i = 0; i < count; i++)
        {
            var c = i < crops.Count ? crops[i] : null;
            var k = i < keeps.Count ? keeps[i] : null;
            if (c is not null && k is not null) throw new AgentToolException($"{i + 1}번째 그림에 crop과 keep을 둘 다 주었습니다.");
            result.Add(c ?? k);
        }
        return result;
    }

    private static (string? Source, string? License) ParseCredit(JsonElement args) =>
        (ToolArgs.OptionalString(args, "source")?.Trim() is { Length: > 0 } s ? s : null,
         ToolArgs.OptionalString(args, "license")?.Trim() is { Length: > 0 } l ? l : null);

    /// <summary>실제 인쇄 해상도 글. 150dpi 미만이면 경고를 붙인다.</summary>
    private static string DpiText(int? pixelWidth, float shapeWidth, CropFractions crop)
    {
        if (pixelWidth is not int px) return "";
        float dpi = PictureFitter.EffectiveDpi(px, shapeWidth, crop);
        return dpi <= 0 ? "" : $" | 약 {dpi:0}dpi" + (dpi < PictureFitter.LowDpi ? " ⚠ 해상도 낮음(150dpi 미만, 인쇄·확대하면 흐림)" : "");
    }

    private static string Describe(PlacedPicture p) =>
        $"  {p.Address} | {p.Left:0.##},{p.Top:0.##} {p.Width:0.##}×{p.Height:0.##}" +
        (p.Crop.IsNone ? " | 자르지 않음" : $" | 잘림 {p.Crop}") +
        (p.PixelWidth is int w && p.PixelHeight is int h ? $" | 원본 {w}×{h}px" : "") +
        DpiText(p.PixelWidth, p.Width, p.Crop);

    /// <summary>넣은 그림 목록 + 해상도가 낮은 그림 요약.</summary>
    private static string DescribeAll(IReadOnlyList<PlacedPicture> placed)
    {
        var text = string.Join("\n", placed.Select(Describe));
        var low = placed.Where(p => p.PixelWidth is int px && PictureFitter.EffectiveDpi(px, p.Width, p.Crop) is > 0 and < PictureFitter.LowDpi).ToList();
        return low.Count == 0 ? text
            : text + $"\n주의: 해상도가 낮은 그림 {low.Count}장({string.Join(", ", low.Select(p => p.Address))}) — 더 큰 원본을 쓰거나 덜 자르거나 칸을 줄이세요.";
    }

    /// <summary>BckPPT가 넣은 뒤 사용자가 자리·크기·자르기를 바꾼 그림이면 무엇이 바뀌었는지.</summary>
    public static string? UserChanges(ShapeEntry e) =>
        e.Picture?.Placed?.Differences(new BoxF(e.Left, e.Top, e.Width, e.Height), e.Picture.Crop);

    /// <summary>그림 파일의 가로/세로 비율(crop 반영). 머리로 못 읽으면 그려 본다.</summary>
    private static float ImageAspect(string file, CropFractions? crop)
    {
        (int W, int H)? px;
        using (var s = File.OpenRead(file)) px = PptxPictureReader.ReadImageSize(s);
        if (px is null)
        {
            try { using var img = System.Drawing.Image.FromFile(file); px = (img.Width, img.Height); }
            catch (Exception ex) when (ex is ArgumentException or OutOfMemoryException) { throw new AgentToolException($"그림 크기를 읽지 못했습니다: {file}"); }
        }
        var c = crop ?? CropFractions.None;
        return px.Value.W * (1 - c.Left - c.Right) / (px.Value.H * (1 - c.Top - c.Bottom));
    }

    private static PictureZ ParseZ(JsonElement args, PictureZ fallback = PictureZ.None) =>
        ToolArgs.OptionalString(args, "z")?.Trim().ToLowerInvariant() switch
        {
            null or "" => fallback, "front" => PictureZ.Front, "back" => PictureZ.Back,
            _ => throw new AgentToolException("z는 front·back 중 하나입니다."),
        };

    private async Task<McpToolResult> AddPicture(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        string file = ImagePath(ToolArgs.OptionalString(args, "path"));
        var crop = ParseCrop(args, "crop");
        var focus = ParseFocus(args);
        var zoom = ParseZoom(args);
        var (source, license) = ParseCredit(args);
        BoxF box;
        string anchorNote = "";
        bool anchored = args.TryGetProperty("anchor", out var anchor) && anchor.ValueKind == JsonValueKind.Object;
        if (anchored)
        {
            string refAddress = ToolArgs.OptionalString(anchor, "shape") ?? throw new AgentToolException("anchor.shape(기준 도형 주소)가 필요합니다.");
            var reference = await FindShapeAsync(path, slide, refAddress, ct);
            string corner = ToolArgs.OptionalString(anchor, "corner") ?? "bottom-right";
            float size = ToolArgs.OptionalFloat(anchor, "size") ?? 0.4f, margin = ToolArgs.OptionalFloat(anchor, "margin") ?? 3f;
            // focus·zoom이면 남은 부분 비율이 칸 비율을 따르니 원본 비율(crop 반영)로 칸을 잡는다.
            float aspect = ImageAspect(file, crop);
            try { box = PictureFitter.Anchor(new BoxF(reference.Left, reference.Top, reference.Width, reference.Height), corner, size, margin, aspect); }
            catch (ArgumentException ex) { throw new AgentToolException(ex.Message); }
            if (ToolArgs.OptionalFloat(args, "width") is float aw && ToolArgs.OptionalFloat(args, "height") is float ah)
                box = box with { Width = aw, Height = ah, Left = corner.Contains("right") ? box.Left + box.Width - aw : box.Left,
                                 Top = corner.Contains("bottom") ? box.Top + box.Height - ah : box.Top };
            anchorNote = $" 「{reference.Address}」 {corner}에 붙임(너비 {size:0.##}배, 여백 {margin:0.##}pt).";
        }
        else
        {
            box = new BoxF(
                ToolArgs.OptionalFloat(args, "left") ?? throw new AgentToolException("인자 'left'가 필요합니다(또는 anchor)."),
                ToolArgs.OptionalFloat(args, "top") ?? throw new AgentToolException("인자 'top'이 필요합니다."),
                ToolArgs.OptionalFloat(args, "width") ?? throw new AgentToolException("인자 'width'가 필요합니다."),
                ToolArgs.OptionalFloat(args, "height") ?? throw new AgentToolException("인자 'height'가 필요합니다."));
        }
        if (box.Width <= 0 || box.Height <= 0) throw new AgentToolException("width·height는 0보다 커야 합니다.");
        var request = new PictureRequest(file, box.Left, box.Top, box.Width, box.Height,
            ParseFit(ToolArgs.OptionalString(args, "fit")), crop, ToolArgs.OptionalString(args, "name"))
        {
            Z = ParseZ(args, anchored ? PictureZ.Front : PictureZ.None),
            Border = ParseBorder(args),
            Shadow = ToolArgs.OptionalBool(args, "shadow") == true,
            Credit = source, License = license,
        };
        request = WithFocus(request, focus, zoom);
        return await WriteAsync(path, async () =>
        {
            var placed = await Task.Run(() => _ppt.PlacePictures(path, slide, [request]), ct);
            return McpToolResult.Ok($"{slide}쪽에 그림을 넣었습니다({FitName(request.Fit)}).{anchorNote}\n" + DescribeAll(placed));
        });
    }

    /// <summary>칸을 그림 수만큼 나누는 자리. rows·columns가 있으면 격자.</summary>
    private static (IReadOnlyList<BoxF> Cells, string Layout) LayoutBox(BoxF box, int count, float gap, int? rows, int? columns, bool? horizontal)
    {
        try
        {
            if (rows is not null || columns is not null)
            {
                int r = rows ?? (count + columns!.Value - 1) / Math.Max(1, columns.Value);
                int cols = columns ?? (count + r - 1) / Math.Max(1, r);
                if (r < 1 || cols < 1) throw new AgentToolException("rows·columns는 1 이상입니다.");
                if (count > r * cols) throw new AgentToolException($"그림 {count}장이 {r}×{cols} 격자({r * cols}칸)에 다 들어가지 않습니다.");
                return (PictureFitter.Grid(box, r, cols, gap), $"{r}×{cols} 격자로");
            }
            return (PictureFitter.Split(box, count, gap, horizontal), "나란히");
        }
        catch (ArgumentException ex) { throw new AgentToolException(ex.Message); }
    }

    /// <summary>add_pictures_in_box·replace_pictures_in_box 공통: 인자에서 그림마다 넣기 요청을 만든다.</summary>
    private static (List<PictureRequest> Requests, string Layout, PictureFit Fit, float Gap) BoxRequests(JsonElement args, IReadOnlyList<string> files, BoxF box)
    {
        float gap = ToolArgs.OptionalFloat(args, "gap") ?? 3f;
        bool? horizontal = ToolArgs.OptionalString(args, "direction")?.Trim().ToLowerInvariant() switch
        {
            null or "" or "auto" => null, "horizontal" => true, "vertical" => false,
            _ => throw new AgentToolException("direction은 auto·horizontal·vertical 중 하나입니다."),
        };
        var (cells, layout) = LayoutBox(box, files.Count, gap, ToolArgs.OptionalInt(args, "rows"), ToolArgs.OptionalInt(args, "columns"), horizontal);
        var fit = ParseFit(ToolArgs.OptionalString(args, "fit"));
        var crops = ParseCropArrays(args, files.Count);
        List<string?> Strings(string name) => args.TryGetProperty(name, out var n) && n.ValueKind == JsonValueKind.Array
            ? n.EnumerateArray().Select(e => e.ValueKind == JsonValueKind.String ? e.GetString() : null).ToList() : [];
        var names = Strings("names");
        var sources = Strings("sources");
        var (source, license) = ParseCredit(args);
        var focuses = args.TryGetProperty("focuses", out var fo) && fo.ValueKind == JsonValueKind.Array
            ? fo.EnumerateArray().Select(e => e.ValueKind is JsonValueKind.Null ? (FocusPoint?)null : ParseFocusValue(e)).ToList() : [];
        float? zoomAll = ParseZoom(args);
        var zooms = args.TryGetProperty("zooms", out var zo) && zo.ValueKind == JsonValueKind.Array
            ? zo.EnumerateArray().Select(e => e.ValueKind == JsonValueKind.Number ? CheckZoom(e.GetSingle()) : (float?)null).ToList() : [];
        string? At(List<string?> list, int i) => i < list.Count && !string.IsNullOrWhiteSpace(list[i]) ? list[i]!.Trim() : null;
        var requests = files.Select((f, i) => WithFocus(
            new PictureRequest(f, cells[i].Left, cells[i].Top, cells[i].Width, cells[i].Height, fit, crops[i], At(names, i))
            {
                Credit = At(sources, i) ?? source, License = license,
            },
            i < focuses.Count ? focuses[i] : null, (i < zooms.Count ? zooms[i] : null) ?? zoomAll)).ToList();
        return (requests, layout, fit, gap);
    }

    private async Task<McpToolResult> AddPicturesInBox(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        var files = ToolArgs.StringArray(args, "paths").Select(ImagePath).ToList();
        if (files.Count == 0) throw new AgentToolException("paths에 그림이 한 장 이상 있어야 합니다.");
        if (!args.TryGetProperty("box", out var boxArg)) throw new AgentToolException("인자 'box'가 필요합니다.");
        var (box, address) = await ResolveBoxAsync(path, slide, boxArg, ct);
        string? remove = null;
        if (args.TryGetProperty("replace_placeholder", out var rp))
        {
            if (rp.ValueKind == JsonValueKind.String && rp.GetString()!.Trim() is { Length: > 0 } other && !bool.TryParse(other, out _))
                remove = (await FindShapeAsync(path, slide, other, ct)).Address;
            else if (ToolArgs.OptionalBool(args, "replace_placeholder") == true)
                remove = address ?? throw new AgentToolException("replace_placeholder=true는 box가 도형 주소일 때만 씁니다. 지울 도형 주소를 replace_placeholder에 직접 주세요.");
        }
        var (requests, layout, fit, gap) = BoxRequests(args, files, box);

        return await WriteAsync(path, async () =>
        {
            var placed = await Task.Run(() => _ppt.PlacePictures(path, slide, requests, remove), ct);
            return McpToolResult.Ok(
                $"{slide}쪽 칸 {box.Left:0.##},{box.Top:0.##} {box.Width:0.##}×{box.Height:0.##}에 {placed.Count}장을 {layout} 넣었습니다({FitName(fit)}, 간격 {gap:0.##}pt)" +
                (remove is not null ? $". 「{remove}」는 지웠습니다" : "") + ".\n" + DescribeAll(placed));
        });
    }

    private async Task<McpToolResult> ReplacePicture(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        string shape = ToolArgs.String(args, "shape");
        string file = ImagePath(ToolArgs.OptionalString(args, "path"));
        var entry = await FindShapeAsync(path, slide, shape, ct);
        if (UserChanges(entry) is { } changed && ToolArgs.OptionalBool(args, "force") != true)
            throw new AgentToolException($"{slide}쪽 「{entry.Address}」는 BckPPT가 넣은 뒤 사용자가 {changed}를 바꾼 그림입니다. " +
                                         "사용자에게 확인받고 덮어쓰려면 force=true로 다시 부르세요.");
        var (source, license) = ParseCredit(args);
        var request = WithFocus(new PictureRequest(file, entry.Left, entry.Top, entry.Width, entry.Height,
            ParseFit(ToolArgs.OptionalString(args, "fit")), ParseCrop(args, "crop"), ToolArgs.OptionalString(args, "name"))
            {
                Credit = source, License = license,
            },
            ParseFocus(args), ParseZoom(args));
        return await WriteAsync(path, async () =>
        {
            var placed = await Task.Run(() => _ppt.PlacePictures(path, slide, [request], entry.Address, keepReplacedName: true), ct);
            return McpToolResult.Ok($"{slide}쪽 「{shape}」를 바꿨습니다(같은 자리·크기·Z순서·테두리·그림자, {FitName(request.Fit)})" +
                                    (entry.Picture is null ? ". 원래 도형은 그림이 아니었습니다" : "") + ".\n" + DescribeAll(placed));
        });
    }

    // ── 칸 안 그림 통째로 바꾸기·계획 적용 ─────────────────────────

    private static bool CenterIn(ShapeEntry e, BoxF box)
    {
        float cx = e.Left + e.Width / 2, cy = e.Top + e.Height / 2;
        return cx >= box.Left && cx <= box.Left + box.Width && cy >= box.Top && cy <= box.Top + box.Height;
    }

    private static bool IsPlaceholderText(ShapeEntry e) =>
        e.Picture is null && e.HasText && e.TextPreview.Trim().StartsWith("[그림", StringComparison.Ordinal);

    private enum BoxOutcome { Replaced, Skipped }

    /// <summary>
    /// 칸 안(중심 기준) 그림을 모두 지우고 새 그림을 넣는다. 「[그림 삽입]」 글상자도 지운다(removePlaceholders).
    /// 사용자가 바꾼 그림이 있으면 skipUserModified면 칸을 건너뛰고, force가 아니면 오류. WriteAsync 안에서 부른다.
    /// </summary>
    private async Task<(BoxOutcome Outcome, string Message, IReadOnlyList<PlacedPicture> Placed)> ReplaceInBoxAsync(
        string path, int slide, BoxF box, string? boxAddress, List<PictureRequest> requests,
        bool force, bool skipUserModified, bool removePlaceholders, CancellationToken ct)
    {
        var shapes = await Task.Run(() => _ppt.ListShapes(path, slide), ct);
        var old = shapes.Where(e => e.Address != boxAddress && CenterIn(e, box) && (e.Picture is not null || removePlaceholders && IsPlaceholderText(e))).ToList();
        var changed = old.Where(e => UserChanges(e) is not null).ToList();
        if (changed.Count > 0 && !force)
        {
            string list = string.Join(", ", changed.Select(e => $"「{e.Address}」({UserChanges(e)})"));
            if (skipUserModified) return (BoxOutcome.Skipped, $"사용자가 바꾼 그림이 있어 건너뜀: {list}", Array.Empty<PlacedPicture>());
            throw new AgentToolException($"{slide}쪽 칸에 사용자가 바꾼 그림이 있습니다: {list}. 건너뛰려면 skip_user_modified=true, " +
                                         "사용자에게 확인받고 덮어쓰려면 force=true.");
        }
        // 그림 하나를 Z순서 자리로 삼고(새 그림이 그 자리에 들어간다) 나머지는 같이 지운다.
        var anchor = old.FirstOrDefault(e => e.Picture is not null) ?? old.FirstOrDefault();
        var rest = old.Where(e => e != anchor).Select(e => e.Address).ToList();
        var placed = await Task.Run(() => _ppt.PlacePictures(path, slide, requests, anchor?.Address, deleteAlso: rest), ct);
        string removed = old.Count == 0 ? "지운 것 없음"
            : $"지움 {old.Count}개({string.Join(", ", old.Select(e => e.Picture is null ? $"{e.Address}(글상자)" : e.Address))})";
        return (BoxOutcome.Replaced, $"{removed}, {placed.Count}장 넣음", placed);
    }

    private async Task<McpToolResult> ReplacePicturesInBox(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        var files = ToolArgs.StringArray(args, "paths").Select(ImagePath).ToList();
        if (files.Count == 0) throw new AgentToolException("paths에 그림이 한 장 이상 있어야 합니다.");
        if (!args.TryGetProperty("box", out var boxArg)) throw new AgentToolException("인자 'box'가 필요합니다.");
        var (box, address) = await ResolveBoxAsync(path, slide, boxArg, ct);
        var (requests, layout, fit, gap) = BoxRequests(args, files, box);
        bool force = ToolArgs.OptionalBool(args, "force") == true;
        bool skip = ToolArgs.OptionalBool(args, "skip_user_modified") == true;
        bool placeholders = ToolArgs.OptionalBool(args, "remove_placeholders") ?? true;
        return await WriteAsync(path, async () =>
        {
            var (outcome, message, placed) = await ReplaceInBoxAsync(path, slide, box, address, requests, force, skip, placeholders, ct);
            if (outcome == BoxOutcome.Skipped) return McpToolResult.Ok($"{slide}쪽 칸을 건너뛰었습니다 — {message}.");
            return McpToolResult.Ok($"{slide}쪽 칸 {box.Left:0.##},{box.Top:0.##} {box.Width:0.##}×{box.Height:0.##}: {message} " +
                                    $"({layout}, {FitName(fit)}, 간격 {gap:0.##}pt).\n" + DescribeAll(placed));
        });
    }

    /// <summary>계획 한 줄: 쪽|칸|경로[@crop=..|@keep=..|@focus=x,y[,zoom]];경로…</summary>
    public sealed record PlanLine(int LineNo, int Slide, string Cell, IReadOnlyList<PlanImage> Images);
    public sealed record PlanImage(string Path, CropFractions? Crop, FocusPoint? Focus, float? Zoom);

    private static readonly System.Text.RegularExpressions.Regex PlanSpec =
        new(@"^(?<path>.*?)\s*@\s*(?<kind>crop|keep|focus)\s*[=:]\s*(?<nums>[-\d.,\s]+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    public static List<PlanLine> ParsePlan(IEnumerable<string> lines, string? baseDir)
    {
        var plan = new List<PlanLine>();
        int no = 0;
        foreach (var raw in lines)
        {
            no++;
            string line = raw.Trim();
            if (line.Length == 0 || line.StartsWith('#')) continue;
            var parts = line.Split('|');
            if (parts.Length != 3 || !int.TryParse(parts[0].Trim(), out int slide) || slide < 1)
                throw new AgentToolException($"계획 {no}줄 「{line}」: 「쪽|칸|경로;경로…」 형식이 아닙니다.");
            var images = new List<PlanImage>();
            foreach (var item in parts[2].Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                string file = item;
                CropFractions? crop = null; FocusPoint? focus = null; float? zoom = null;
                if (PlanSpec.Match(item) is { Success: true } m)
                {
                    file = m.Groups["path"].Value.Trim();
                    var nums = m.Groups["nums"].Value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .Select(v => float.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var f) ? f
                            : throw new AgentToolException($"계획 {no}줄: 숫자가 아닙니다: {v}")).ToArray();
                    switch (m.Groups["kind"].Value.ToLowerInvariant())
                    {
                        case "crop" when nums.Length == 4: crop = CheckCrop(new CropFractions(nums[0], nums[1], nums[2], nums[3])); break;
                        case "keep" when nums.Length == 4:
                            if (nums[2] <= nums[0] || nums[3] <= nums[1]) throw new AgentToolException($"계획 {no}줄: keep은 x0,y0,x1,y1(x0<x1, y0<y1)입니다.");
                            crop = CheckCrop(PictureFitter.FromKeep(nums[0], nums[1], nums[2], nums[3]), "keep"); break;
                        case "focus" when nums.Length is 2 or 3:
                            if (nums[0] is < 0 or > 1 || nums[1] is < 0 or > 1) throw new AgentToolException($"계획 {no}줄: focus는 0~1입니다.");
                            focus = new FocusPoint(nums[0], nums[1]);
                            if (nums.Length == 3) zoom = CheckZoom(nums[2]);
                            break;
                        default: throw new AgentToolException($"계획 {no}줄: @crop=l,t,r,b · @keep=x0,y0,x1,y1 · @focus=x,y[,zoom] 중 하나입니다.");
                    }
                }
                file = file.Trim('"');
                if (!Path.IsPathFullyQualified(file))
                {
                    if (baseDir is null) throw new AgentToolException($"계획 {no}줄: 상대 경로 「{file}」는 base_dir가 있어야 합니다.");
                    var here = Path.GetFullPath(Path.Combine(baseDir, file));
                    // 계획 파일이 _scripts 같은 하위 폴더에 있으면 그 위 폴더 기준일 수 있다.
                    var up = Directory.GetParent(baseDir) is { } parent ? Path.GetFullPath(Path.Combine(parent.FullName, file)) : here;
                    file = File.Exists(here) || !File.Exists(up) ? here : up;
                }
                images.Add(new PlanImage(file, crop, focus, zoom));
            }
            if (images.Count == 0) throw new AgentToolException($"계획 {no}줄: 그림 경로가 없습니다.");
            plan.Add(new PlanLine(no, slide, parts[1].Trim(), images));
        }
        return plan;
    }

    private async Task<McpToolResult> ApplyPicturePlan(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        string? planPath = ToolArgs.OptionalString(args, "plan_path")?.Trim().Trim('"');
        string? baseDir = ToolArgs.OptionalString(args, "base_dir")?.Trim().Trim('"');
        List<string> lines;
        if (!string.IsNullOrEmpty(planPath))
        {
            if (!Path.IsPathFullyQualified(planPath) || !File.Exists(planPath)) throw new AgentToolException($"계획 파일이 없습니다: {planPath}");
            lines = File.ReadAllLines(planPath).ToList();
            baseDir ??= Path.GetDirectoryName(Path.GetFullPath(planPath));
        }
        else lines = args.TryGetProperty("plan", out var pl) && pl.ValueKind == JsonValueKind.Array ? ToolArgs.StringArray(args, "plan").ToList() : [];
        if (lines.Count == 0) throw new AgentToolException("plan_path나 plan(줄 배열)을 주세요.");
        var plan = ParsePlan(lines, baseDir);
        if (args.TryGetProperty("slides", out var only) && only.ValueKind == JsonValueKind.Array)
        {
            var keep = only.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Number).Select(e => e.GetInt32()).ToHashSet();
            plan = plan.Where(p => keep.Contains(p.Slide)).ToList();
        }
        if (plan.Count == 0) throw new AgentToolException("적용할 줄이 없습니다.");
        var cells = args.TryGetProperty("cells", out var c) && c.ValueKind == JsonValueKind.Object ? c : default;
        bool force = ToolArgs.OptionalBool(args, "force") == true;
        bool skip = ToolArgs.OptionalBool(args, "skip_user_modified") ?? true;
        bool dryRun = ToolArgs.OptionalBool(args, "dry_run") == true;
        bool render = ToolArgs.OptionalBool(args, "render") ?? true;
        float gap = ToolArgs.OptionalFloat(args, "gap") ?? 3f;
        var fit = ParseFit(ToolArgs.OptionalString(args, "fit"));
        int slideCount = _ppt.ListSlides(path).Count;

        var rows = new List<string>();
        var changedSlides = new SortedSet<int>();
        var lowDpi = new List<string>();
        async Task<McpToolResult> Run()
        {
            foreach (var line in plan)
            {
                string tag = $"{line.Slide}쪽 {line.Cell}";
                try
                {
                    if (line.Slide > slideCount) throw new AgentToolException($"문서는 {slideCount}쪽까지입니다.");
                    var missing = line.Images.Where(i => !File.Exists(i.Path)).Select(i => i.Path).ToList();
                    if (missing.Count > 0) throw new AgentToolException("파일 없음: " + string.Join(", ", missing));
                    // 칸: cells에 정의(상자 또는 도형 주소)가 있으면 그것, 없으면 쪽의 도형 이름.
                    JsonElement cellArg = cells.ValueKind == JsonValueKind.Object && cells.TryGetProperty(line.Cell, out var def)
                        ? def : JsonDocument.Parse(JsonSerializer.Serialize(line.Cell)).RootElement;
                    var (box, address) = await ResolveBoxAsync(path, line.Slide, cellArg, ct);
                    var (boxes, layout) = LayoutBox(box, line.Images.Count, gap, null, null, null);
                    var requests = line.Images.Select((img, i) => WithFocus(
                        new PictureRequest(ImagePath(img.Path), boxes[i].Left, boxes[i].Top, boxes[i].Width, boxes[i].Height, fit, img.Crop),
                        img.Focus, img.Zoom)).ToList();
                    if (dryRun)
                    {
                        rows.Add($"| {line.Slide} | {line.Cell} | 할 일 | {line.Images.Count}장 {layout} → 칸 {box.Left:0.#},{box.Top:0.#} {box.Width:0.#}×{box.Height:0.#} |");
                        continue;
                    }
                    var (outcome, message, placed) = await ReplaceInBoxAsync(path, line.Slide, box, address, requests, force, skip, true, ct);
                    if (outcome == BoxOutcome.Replaced) changedSlides.Add(line.Slide);
                    foreach (var p in placed)
                        if (p.PixelWidth is int px && PictureFitter.EffectiveDpi(px, p.Width, p.Crop) is var dpi and > 0 and < PictureFitter.LowDpi)
                            lowDpi.Add($"{tag} {p.Address} 약 {dpi:0}dpi");
                    rows.Add($"| {line.Slide} | {line.Cell} | {(outcome == BoxOutcome.Replaced ? "넣음" : "건너뜀")} | {message} |");
                }
                catch (Exception ex) when (ex is AgentToolException or InvalidOperationException or IOException)
                {
                    rows.Add($"| {line.Slide} | {line.Cell} | 오류 | {ex.Message.Replace("|", "/")} |");
                }
            }
            var sb = new StringBuilder();
            sb.AppendLine($"{Path.GetFileName(path)} 그림 계획 {plan.Count}줄{(dryRun ? " (dry_run: 문서를 고치지 않음)" : "")}");
            sb.AppendLine("| 쪽 | 칸 | 결과 | 내용 |");
            sb.AppendLine("|---|---|---|---|");
            foreach (var r in rows) sb.AppendLine(r);
            if (lowDpi.Count > 0) sb.AppendLine("해상도 낮음(150dpi 미만): " + string.Join("; ", lowDpi));
            var content = new List<McpContent> { McpContent.FromText(sb.ToString().TrimEnd()) };
            if (render && changedSlides.Count > 0)
            {
                var pngs = new List<(int, string)>();
                try
                {
                    foreach (var n in changedSlides.Take(MaxSheetRenders))
                        if (await _ppt.GetSlideThumbnailAsync(path, n) is { } png) pngs.Add((n, png));
                    if (pngs.Count > 0) content.Add(McpContent.FromImage(ContactSheet.Compose(pngs, Math.Min(4, pngs.Count))));
                }
                finally { foreach (var (_, png) in pngs) try { File.Delete(png); } catch { } }
            }
            return new McpToolResult(content, false);
        }
        return dryRun ? await Run() : await WriteAsync(path, Run);
    }

    // ── 후보 사진 축소판·영역 렌더 ─────────────────────────────

    public const int MaxSheetImages = 40;

    private async Task<McpToolResult> ImageSheet(JsonElement args, CancellationToken ct)
    {
        var files = ToolArgs.StringArray(args, "paths").Select(ImagePath).ToList();
        if (files.Count == 0) throw new AgentToolException("paths에 그림이 한 장 이상 있어야 합니다.");
        if (files.Count > MaxSheetImages) throw new AgentToolException($"한 번에 {MaxSheetImages}장까지입니다.");
        int columns = Math.Clamp(ToolArgs.OptionalInt(args, "columns") ?? 4, 1, 8);
        var items = new List<(string, string, byte[]?)>();
        var lines = new List<string>();
        for (int i = 0; i < files.Count; i++)
        {
            var bytes = await File.ReadAllBytesAsync(files[i], ct);
            (int W, int H)? px;
            using (var ms = new MemoryStream(bytes)) px = PptxPictureReader.ReadImageSize(ms);
            if (px is null) try { using var img = System.Drawing.Image.FromStream(new MemoryStream(bytes)); px = (img.Width, img.Height); } catch (ArgumentException) { }
            string size = px is { } p ? $"{p.W}×{p.H}px · {(float)p.W / p.H:0.##}" : "크기 모름";
            string detail = $"{size} · {Mb(bytes.Length)}";
            items.Add(($"{i + 1}. {Path.GetFileName(files[i])}", detail, bytes));
            lines.Add($"{i + 1}. {files[i]} — {detail}");
        }
        var png = ContactSheet.ComposeImages(items, columns);
        return new([McpContent.FromText($"후보 {files.Count}장(번호는 paths 순서)\n" + string.Join("\n", lines)), McpContent.FromImage(png)], false);
    }

    private async Task<McpToolResult> RenderRegion(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        JsonElement target = args.TryGetProperty("box", out var b) ? b
            : args.TryGetProperty("shape", out var s) ? s : throw new AgentToolException("box({left,top,width,height}) 또는 shape(도형 주소)를 주세요.");
        var (box, address) = await ResolveBoxAsync(path, slide, target, ct);
        float margin = ToolArgs.OptionalFloat(args, "margin") ?? 6f;
        // scale: pt당 픽셀. 비우면 영역이 1200px 너비가 되게(1~8).
        float scale = ToolArgs.OptionalFloat(args, "scale") ?? Math.Clamp(1200f / (box.Width + 2 * margin), 1f, 8f);
        if (scale is < 0.5f or > 10f) throw new AgentToolException("scale은 0.5~10(pt당 픽셀)입니다.");
        var rendered = await _ppt.RenderSlideAsync(path, slide, scale) ?? throw new AgentToolException($"{slide}쪽을 렌더하지 못했습니다.");
        try
        {
            using var page = System.Drawing.Image.FromFile(rendered.PngPath);
            float k = page.Width / rendered.SlideWidth;
            var rect = System.Drawing.Rectangle.Round(new System.Drawing.RectangleF(
                (box.Left - margin) * k, (box.Top - margin) * k, (box.Width + 2 * margin) * k, (box.Height + 2 * margin) * k));
            rect.Intersect(new System.Drawing.Rectangle(0, 0, page.Width, page.Height));
            if (rect.Width < 2 || rect.Height < 2) throw new AgentToolException("영역이 쪽 밖입니다.");
            using var part = new System.Drawing.Bitmap(rect.Width, rect.Height);
            using (var g = System.Drawing.Graphics.FromImage(part))
                g.DrawImage(page, new System.Drawing.Rectangle(0, 0, rect.Width, rect.Height), rect, System.Drawing.GraphicsUnit.Pixel);
            using var ms = new MemoryStream();
            part.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            string what = address is not null ? $"「{address}」" : $"{box.Left:0.##},{box.Top:0.##} {box.Width:0.##}×{box.Height:0.##}";
            return new([McpContent.FromText($"{Path.GetFileName(path)} {slide}쪽 {what} (여백 {margin:0.#}pt, {rect.Width}×{rect.Height}px, pt당 {k:0.##}px)"),
                        McpContent.FromImage(ms.ToArray())], false);
        }
        finally { try { File.Delete(rendered.PngPath); } catch { } }
    }

    // ── 문서 크기·그림 압축 ─────────────────────────────────────

    private static string Mb(long bytes) => bytes >= 1024 * 1024 ? $"{bytes / 1024.0 / 1024.0:0.##}MB" : $"{bytes / 1024.0:0.#}KB";

    private Task<McpToolResult> DocumentSizeTool(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int top = Math.Clamp(ToolArgs.OptionalInt(args, "top") ?? 15, 1, 100);
        var size = PptxPictureReader.ReadDocumentSize(path) ?? throw new AgentToolException("pptx를 읽지 못했습니다(.pptx만, 디스크에 저장된 판).");
        var sb = new StringBuilder();
        sb.AppendLine($"{Path.GetFileName(path)} {Mb(size.FileBytes)} (디스크에 저장된 판 기준 — 창에서 고치고 저장하지 않은 것은 빠진다)");
        sb.AppendLine($"그림 파일 {size.MediaCount}개 {Mb(size.MediaBytes)} ({(size.FileBytes > 0 ? size.MediaBytes * 100.0 / size.FileBytes : 0):0}%)");
        var byMedia = size.Pictures.GroupBy(p => p.Media).Select(g => g.First()).ToList();
        var cropped = byMedia.Where(p => p.HasCrop).ToList();
        sb.AppendLine($"자른 부분까지 든 그림 {cropped.Count}개 {Mb(cropped.Sum(p => p.Bytes))} — compress_pictures(delete_cropped=true)로 잘린 부분을 버리면 준다");
        sb.AppendLine("쪽별 그림 크기: " + string.Join(", ", size.Pictures.GroupBy(p => p.Slide)
            .Select(g => (Slide: g.Key, Bytes: g.GroupBy(p => p.Media).Sum(m => m.First().Bytes)))
            .OrderByDescending(x => x.Bytes).Take(10).Select(x => $"{x.Slide}쪽 {Mb(x.Bytes)}")));
        sb.AppendLine($"큰 그림 {Math.Min(top, size.Pictures.Count)}개:");
        foreach (var p in size.Pictures.OrderByDescending(p => p.Bytes).Take(top))
            sb.AppendLine($"  {p.Slide}쪽 「{p.ShapeName}」 @{p.ShapeId} | {Mb(p.Bytes)} | " +
                          (p.PixelWidth is int w && p.PixelHeight is int h ? $"{w}×{h}px" : "크기 모름") +
                          (p.HasCrop ? " | 자른 부분 포함" : "") + (p.SharedBy > 1 ? $" | {p.SharedBy}곳에서 같이 씀" : "") +
                          $" | {Path.GetFileName(p.Media)}");
        return Task.FromResult(McpToolResult.Ok(sb.ToString().TrimEnd()));
    }

    /// <summary>한 그림을 줄인 결과(적용 전). Bytes가 null이면 건너뛴 까닭이 Reason.</summary>
    private sealed record Compressed(int Slide, ShapeEntry Entry, long OldBytes, byte[]? Bytes, string Ext, string Reason, PictureSource? Source);

    private static Compressed CompressOne(int slide, ShapeEntry entry, PictureSource src, float dpi, bool deleteCropped, int quality)
    {
        var info = src.Info;
        var crop = info.Crop;
        if (deleteCropped && new[] { crop.Left, crop.Top, crop.Right, crop.Bottom }.Any(v => v < -0.001f))
            return new(slide, entry, src.Bytes.Length, null, "", "그림 밖 여백이 있어 건너뜀", src);
        using var image = System.Drawing.Image.FromStream(new MemoryStream(src.Bytes));
        var region = deleteCropped && !crop.IsNone
            ? System.Drawing.Rectangle.Round(new System.Drawing.RectangleF(crop.Left * image.Width, crop.Top * image.Height,
                (1 - crop.Left - crop.Right) * image.Width, (1 - crop.Top - crop.Bottom) * image.Height))
            : new System.Drawing.Rectangle(0, 0, image.Width, image.Height);
        region.Intersect(new System.Drawing.Rectangle(0, 0, image.Width, image.Height));
        float shownWidth = deleteCropped ? src.Frame.Width : info.FullWidth;   // 이 그림(영역)이 쪽에서 차지하는 너비(pt)
        float targetW = shownWidth / 72f * dpi;
        float scale = Math.Min(1f, targetW / region.Width);
        bool cropsAway = region.Width < image.Width || region.Height < image.Height;
        if (scale > 0.95f && !cropsAway) return new(slide, entry, src.Bytes.Length, null, "", $"이미 {dpi:0}dpi 안팎", src);
        int w = Math.Max(1, (int)Math.Round(region.Width * scale)), h = Math.Max(1, (int)Math.Round(region.Height * scale));
        bool alpha = System.Drawing.Image.IsAlphaPixelFormat(image.PixelFormat);
        using var bmp = new System.Drawing.Bitmap(w, h, alpha ? System.Drawing.Imaging.PixelFormat.Format32bppArgb : System.Drawing.Imaging.PixelFormat.Format24bppRgb);
        using (var g = System.Drawing.Graphics.FromImage(bmp))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            g.DrawImage(image, new System.Drawing.Rectangle(0, 0, w, h), region, System.Drawing.GraphicsUnit.Pixel);
        }
        using var ms = new MemoryStream();
        string ext;
        if (alpha) { bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png); ext = ".png"; }
        else
        {
            var codec = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders().First(e => e.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
            using var ps = new System.Drawing.Imaging.EncoderParameters(1);
            ps.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)quality);
            bmp.Save(ms, codec, ps);
            ext = ".jpg";
        }
        var bytes = ms.ToArray();
        if (bytes.Length >= src.Bytes.Length * 0.9) return new(slide, entry, src.Bytes.Length, null, "", "줄여도 크기가 거의 같음", src);
        return new(slide, entry, src.Bytes.Length, bytes, ext, $"{image.Width}×{image.Height} → {w}×{h}px" + (cropsAway ? ", 잘린 부분 버림" : ""), src);
    }

    private async Task<McpToolResult> CompressPictures(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        float dpi = ToolArgs.OptionalFloat(args, "dpi") ?? 200f;
        if (dpi is < 72 or > 600) throw new AgentToolException("dpi는 72~600입니다.");
        bool deleteCropped = ToolArgs.OptionalBool(args, "delete_cropped") ?? true;
        int quality = Math.Clamp(ToolArgs.OptionalInt(args, "jpeg_quality") ?? 85, 50, 100);
        bool confirm = ToolArgs.OptionalBool(args, "confirm") == true;
        int count = _ppt.ListSlides(path).Count;
        var slides = args.TryGetProperty("slides", out var sl) && sl.ValueKind == JsonValueKind.Array
            ? sl.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Number).Select(e => e.GetInt32()).Where(n => n >= 1 && n <= count).Distinct().ToList()
            : Enumerable.Range(1, count).ToList();

        var results = new List<Compressed>();
        foreach (var slide in slides)
        {
            ct.ThrowIfCancellationRequested();
            foreach (var e in (await Task.Run(() => _ppt.ListShapes(path, slide), ct)).Where(e => e.Picture is not null))
            {
                try
                {
                    var src = await Task.Run(() => _ppt.GetPictureSource(path, slide, e.Address), ct);
                    results.Add(CompressOne(slide, e, src, dpi, deleteCropped, quality));
                }
                catch (Exception ex) when (ex is InvalidOperationException or ArgumentException or IOException or System.Runtime.InteropServices.COMException)
                {
                    results.Add(new(slide, e, 0, null, "", "읽지 못함: " + ex.Message, null));
                }
            }
        }
        var todo = results.Where(r => r.Bytes is not null).ToList();
        long before = todo.Sum(r => r.OldBytes), after = todo.Sum(r => (long)r.Bytes!.Length);
        var sb = new StringBuilder();
        sb.AppendLine($"{Path.GetFileName(path)} 그림 {results.Count}개 중 {todo.Count}개를 줄일 수 있습니다: {Mb(before)} → {Mb(after)} " +
                      $"(목표 {dpi:0}dpi{(deleteCropped ? ", 잘린 부분 버림" : "")}, JPEG 품질 {quality})");
        foreach (var r in results.OrderBy(r => r.Slide))
            sb.AppendLine($"  {r.Slide}쪽 {r.Entry.Address} | " + (r.Bytes is null ? $"그대로 — {r.Reason}" : $"{Mb(r.OldBytes)} → {Mb(r.Bytes.Length)} | {r.Reason}"));
        if (!confirm || todo.Count == 0)
        {
            if (todo.Count > 0)
                sb.Append("문서는 아직 그대로입니다. 적용하면 원본(자르기 전)을 문서에서 버려 나중에 자르기를 넓힐 수 없습니다. " +
                          "사용자에게 확인받은 뒤 confirm=true로 다시 부르세요.");
            return McpToolResult.Ok(sb.ToString().TrimEnd());
        }
        return await WriteAsync(path, async () =>
        {
            var temp = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), $"bckppt-compress-{Guid.NewGuid():N}"));
            int done = 0;
            var failed = new List<string>();
            try
            {
                foreach (var r in todo)
                {
                    var file = Path.Combine(temp.FullName, $"{r.Slide}-{r.Entry.Id}{r.Ext}");
                    await File.WriteAllBytesAsync(file, r.Bytes!, ct);
                    var e = r.Entry;
                    var request = new PictureRequest(file, e.Left, e.Top, e.Width, e.Height, PictureFit.Stretch, deleteCropped ? null : e.Picture!.Crop)
                    {
                        RecordedSource = e.Picture!.SourcePath ?? r.Source?.SourcePath, Credit = e.Picture.Credit, License = e.Picture.License,
                    };
                    try { await Task.Run(() => _ppt.PlacePictures(path, r.Slide, [request], e.Address, keepReplacedName: true), ct); done++; }
                    catch (Exception ex) when (ex is InvalidOperationException or System.Runtime.InteropServices.COMException) { failed.Add($"{r.Slide}쪽 {e.Address}: {ex.Message}"); }
                }
            }
            finally { try { temp.Delete(true); } catch { } }
            sb.AppendLine($"적용: {done}개를 바꿨습니다(같은 자리·크기·Z순서·이름·테두리). 파일 크기는 저장한 뒤 document_size로 확인하세요.");
            if (failed.Count > 0) sb.AppendLine("실패: " + string.Join("; ", failed));
            return McpToolResult.Ok(sb.ToString().TrimEnd());
        });
    }

    // ── 출처 ────────────────────────────────────────────────

    private const string CreditMarker = "그림 출처:";

    private async Task<McpToolResult> ExportPictureCredits(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int count = _ppt.ListSlides(path).Count;
        var slides = args.TryGetProperty("slides", out var sl) && sl.ValueKind == JsonValueKind.Array
            ? sl.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Number).Select(e => e.GetInt32()).Where(n => n >= 1 && n <= count).Distinct().Order().ToList()
            : Enumerable.Range(1, count).ToList();
        bool notes = ToolArgs.OptionalBool(args, "write_notes") == true;
        string? output = ToolArgs.OptionalString(args, "output_path")?.Trim().Trim('"');
        if (output is { Length: > 0 } && !Path.IsPathFullyQualified(output)) throw new AgentToolException("output_path는 절대 경로여야 합니다.");

        var md = new StringBuilder($"# {Path.GetFileNameWithoutExtension(path)} 그림 출처\n");
        var perSlide = new Dictionary<int, List<string>>();
        int missing = 0;
        foreach (var slide in slides)
        {
            var pics = (await Task.Run(() => _ppt.ListShapes(path, slide), ct)).Where(e => e.Picture is not null).ToList();
            if (pics.Count == 0) continue;
            md.AppendLine($"\n## {slide}쪽");
            var credits = new List<string>();
            foreach (var e in pics)
            {
                var p = e.Picture!;
                string file = p.SourcePath is { } sp ? Path.GetFileName(sp) : "";
                if (p.Credit is null) missing++;
                string credit = p.Credit is null ? "출처 미기록" : p.Credit + (p.License is { } l ? $" ({l})" : "");
                md.AppendLine($"- {e.Address}: {credit}" + (file.Length > 0 ? $" · 파일 {file}" : ""));
                if (p.Credit is not null && !credits.Contains(credit)) credits.Add(credit);
            }
            perSlide[slide] = credits;
        }
        if (perSlide.Count == 0) return McpToolResult.Ok("그림이 있는 쪽이 없습니다.");
        if (missing > 0) md.AppendLine($"\n(출처 미기록 {missing}개 — add_picture 등의 source·license 인자로 넣으면 기록된다)");
        if (output is { Length: > 0 })
        {
            Directory.CreateDirectory(Path.GetDirectoryName(output)!);
            await File.WriteAllTextAsync(output, md.ToString(), new UTF8Encoding(false), ct);
        }
        string tail = output is { Length: > 0 } ? $"\n(저장: {output})" : "";
        if (!notes) return McpToolResult.Ok(md.ToString().TrimEnd() + tail);
        return await WriteAsync(path, async () =>
        {
            int written = 0;
            foreach (var (slide, credits) in perSlide.Where(kv => kv.Value.Count > 0))
            {
                await Task.Run(() => _ppt.SetNotesParagraph(path, slide, CreditMarker, $"{CreditMarker} " + string.Join("; ", credits)), ct);
                written++;
            }
            return McpToolResult.Ok(md.ToString().TrimEnd() + tail + $"\n발표자 메모 {written}쪽에 「{CreditMarker} …」 문단을 썼습니다(다시 부르면 그 문단을 바꾼다).");
        });
    }

    private static FocusPoint ParseFocusValue(JsonElement f)
    {
        if (f.ValueKind != JsonValueKind.Object) throw new AgentToolException("focus는 {x,y}(원본 기준 0~1) 객체입니다.");
        float x = ToolArgs.OptionalFloat(f, "x") ?? throw new AgentToolException("focus.x가 필요합니다.");
        float y = ToolArgs.OptionalFloat(f, "y") ?? throw new AgentToolException("focus.y가 필요합니다.");
        if (x is < 0 or > 1 || y is < 0 or > 1) throw new AgentToolException("focus의 x·y는 0~1(원본 기준 비율)입니다.");
        return new FocusPoint(x, y);
    }

    private static FocusPoint? ParseFocus(JsonElement args) =>
        args.TryGetProperty("focus", out var f) && f.ValueKind is not (JsonValueKind.Null or JsonValueKind.Undefined) ? ParseFocusValue(f) : null;

    private static float CheckZoom(float zoom) => zoom is >= 1 and <= 20 ? zoom
        : throw new AgentToolException("zoom은 1~20입니다(1 = 칸을 겨우 채움, 2 = 두 배 확대).");

    private static float? ParseZoom(JsonElement args) => ToolArgs.OptionalFloat(args, "zoom") is float z ? CheckZoom(z) : null;

    private static PictureRequest WithFocus(PictureRequest r, FocusPoint? focus, float? zoom)
    {
        if (focus is null && zoom is null) return r;
        if (r.Fit != PictureFit.Cover) throw new AgentToolException("focus·zoom은 fit=cover일 때만 씁니다.");
        return r with { Focus = focus, Zoom = zoom };
    }

    private static PictureBorder? ParseBorder(JsonElement args)
    {
        if (!args.TryGetProperty("border", out var b) || b.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined) return null;
        if (b.ValueKind == JsonValueKind.True) return new PictureBorder("#FFFFFF", 1.5f);
        if (b.ValueKind != JsonValueKind.Object) throw new AgentToolException("border는 {color,width} 객체입니다.");
        string color = ToolArgs.OptionalString(b, "color")?.Trim() is { Length: > 0 } c ? (c.StartsWith('#') ? c : "#" + c) : "#FFFFFF";
        if (!System.Text.RegularExpressions.Regex.IsMatch(color, "^#[0-9A-Fa-f]{6}$")) throw new AgentToolException("border.color는 #RRGGBB입니다.");
        float width = ToolArgs.OptionalFloat(b, "width") ?? 1.5f;
        if (width is <= 0 or > 20) throw new AgentToolException("border.width는 0보다 크고 20pt 이하입니다.");
        return new PictureBorder(color, width);
    }

    /// <summary>set_picture_crop·picture_preview의 crop/focus/zoom/reset/keep_frame. 아무것도 없으면 null.</summary>
    private static PictureCropRequest? ParseCropRequest(JsonElement args)
    {
        var crop = ParseCrop(args, "crop");
        var focus = ParseFocus(args);
        var zoom = ParseZoom(args);
        bool reset = ToolArgs.OptionalBool(args, "reset") == true;
        bool keepFrame = ToolArgs.OptionalBool(args, "keep_frame") ?? true;
        if (reset && (crop is not null || focus is not null || zoom is not null))
            throw new AgentToolException("reset은 crop·focus·zoom과 같이 쓰지 않습니다.");
        if (!reset && crop is null && focus is null && zoom is null) return null;
        return new PictureCropRequest(crop, focus, zoom, reset, keepFrame);
    }

    private async Task<ShapeEntry> FindPictureAsync(string path, int slide, string address, CancellationToken ct)
    {
        var entry = await FindShapeAsync(path, slide, address, ct);
        return entry.Picture is null
            ? throw new AgentToolException($"{slide}쪽 「{entry.Address}」는 그림이 아닙니다. list_shapes(type:\"picture\")로 그림 주소를 확인하세요.")
            : entry;
    }

    private async Task<McpToolResult> SetPictureCrop(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        var request = ParseCropRequest(args) ?? throw new AgentToolException("crop, focus·zoom, reset 중 하나를 주세요.");
        var entry = await FindPictureAsync(path, slide, ToolArgs.String(args, "shape"), ct);
        return await WriteAsync(path, async () =>
        {
            var placed = await Task.Run(() => _ppt.SetPictureCrop(path, slide, entry.Address, request), ct);
            return McpToolResult.Ok(
                $"{slide}쪽 「{placed.Address}」의 자르기를 바꿨습니다({(request.KeepFrame ? "도형 자리·크기 그대로" : "도형이 자른 영역을 따라감")}).\n" +
                $"  전: {DescribeCrop(entry.Picture!.Crop)}\n  후: {DescribeCrop(placed.Crop)}\n" + Describe(placed) +
                "\npicture_preview로 원본 위의 자른 영역을, render_slide로 쪽을 확인하세요.");
        });
    }

    private static string DescribeCrop(CropFractions c) => c.IsNone ? "자르지 않음" : $"{c} (남은 영역 {PicturePreview.Range(c)})";

    private async Task<McpToolResult> PicturePreviewTool(JsonElement args, CancellationToken ct)
    {
        var path = ResolveFile(args);
        int slide = ToolArgs.Int(args, "slide");
        var request = ParseCropRequest(args);
        var entry = await FindPictureAsync(path, slide, ToolArgs.String(args, "shape"), ct);
        var source = await Task.Run(() => _ppt.GetPictureSource(path, slide, entry.Address), ct);
        var info = source.Info;
        CropFractions? proposed = null;
        if (request is not null)
        {
            try { proposed = PictureFitter.Recrop(source.Frame, info.FullWidth, info.FullHeight, info.Crop, request).Crop; }
            catch (ArgumentException ex) { throw new AgentToolException(ex.Message); }
        }
        var (png, pw, ph) = PicturePreview.Render(source.Bytes, info.Crop, proposed);
        bool exported = source.Origin.StartsWith("PowerPoint");
        int? pixelW = exported ? entry.Picture?.PixelWidth : pw, pixelH = exported ? entry.Picture?.PixelHeight : ph;
        float ratio = source.Frame.Width / source.Frame.Height;
        var (focus, zoom) = PicturePreview.AsFocus(info.Crop, info.FullWidth, info.FullHeight, ratio);
        var sb = new StringBuilder();
        sb.AppendLine($"{Path.GetFileName(path)} {slide}쪽 「{entry.Address}」 원본 전체 — 출처: {source.Origin}{(source.SourcePath is { } sp ? $" ({sp})" : "")}");
        sb.AppendLine(pixelW is int w && pixelH is int h ? $"원본 {w}×{h}px" : "원본 픽셀: 모름(저장하지 않은 그림을 PowerPoint에서 뽑았다)");
        sb.AppendLine($"도형 {source.Frame.Left:0.##},{source.Frame.Top:0.##} {source.Frame.Width:0.##}×{source.Frame.Height:0.##}pt (가로:세로 {ratio:0.##})");
        sb.AppendLine($"지금 자름(빨간 실선, 바깥 어둡게): {DescribeCrop(info.Crop)}");
        sb.AppendLine($"  같은 영역을 focus·zoom으로: focus {{x:{focus.X:0.###}, y:{focus.Y:0.###}}}, zoom {zoom:0.##}");
        if (proposed is not null)
            sb.AppendLine($"바꿔 볼 영역(노란 점선): {DescribeCrop(proposed)} — set_picture_crop에 같은 인자를 주면 이렇게 됩니다");
        sb.Append("격자: 원본 기준 0.1 간격, 가운데 0.5는 굵은 선. 눈금은 위(x)·왼쪽(y)");
        return new([McpContent.FromText(sb.ToString()), McpContent.FromImage(png)], false);
    }

    private async Task<McpToolResult> ImagePreviewTool(JsonElement args, CancellationToken ct)
    {
        string file = ImagePath(ToolArgs.OptionalString(args, "path"));
        var bytes = await File.ReadAllBytesAsync(file, ct);
        float? width = ToolArgs.OptionalFloat(args, "width"), height = ToolArgs.OptionalFloat(args, "height");
        var crop = ParseCrop(args, "crop");
        var focus = ParseFocus(args);
        var zoom = ParseZoom(args);
        var fit = ParseFit(ToolArgs.OptionalString(args, "fit"));
        int pw, ph;
        using (var ms = new MemoryStream(bytes))
            (pw, ph) = PptxPictureReader.ReadImageSize(ms) ?? (0, 0);
        CropFractions? proposed = crop;
        string? note = null;
        if (width is float bw && height is float bh)
        {
            if (bw <= 0 || bh <= 0) throw new AgentToolException("width·height는 0보다 커야 합니다.");
            if (pw <= 0 || ph <= 0)
            {
                // 머리에서 크기를 못 읽는 형식(emf·tif 등)은 그려 보며 잰다.
                (_, pw, ph) = PicturePreview.Render(bytes, null, null);
            }
            var request = WithFocus(new PictureRequest(file, 0, 0, bw, bh, fit, crop), focus, zoom);
            try { proposed = PictureFitter.Place(request, pw, ph).Crop; }
            catch (ArgumentException ex) { throw new AgentToolException(ex.Message); }
            note = $"칸 {bw:0.##}×{bh:0.##}pt(가로:세로 {bw / bh:0.##})에 {FitName(fit)}" + (focus is not null || zoom is not null ? $", focus·zoom" : "");
        }
        else if (focus is not null || zoom is not null)
            throw new AgentToolException("focus·zoom을 미리 보려면 넣을 칸의 width·height를 주세요.");
        var (png, rw, rh) = PicturePreview.Render(bytes, null, proposed is { IsNone: false } ? proposed : null);
        var sb = new StringBuilder();
        sb.AppendLine($"{Path.GetFileName(file)} 원본 {rw}×{rh}px (가로:세로 {(float)rw / rh:0.##})");
        if (proposed is not null)
            sb.AppendLine((note is null ? "" : note + " → ") + $"잘라 낼 영역(노란 점선): {DescribeCrop(proposed)}");
        sb.Append("격자: 원본 기준 0.1 간격, 가운데 0.5는 굵은 선. 눈금은 위(x)·왼쪽(y)");
        return new([McpContent.FromText(sb.ToString()), McpContent.FromImage(png)], false);
    }

    private static string FitName(PictureFit fit) => fit switch
    {
        PictureFit.Contain => "contain: 비율 유지",
        PictureFit.Stretch => "stretch: 칸에 늘림",
        _ => "cover: 칸 채우고 넘는 쪽 자름",
    };

    private async Task<ShapeEntry> FindShapeAsync(string path, int slide, string address, CancellationToken ct)
    {
        var shapes = await Task.Run(() => _ppt.ListShapes(path, slide), ct);
        return shapes.FirstOrDefault(s => s.Address == address) ?? shapes.FirstOrDefault(s => s.Name == address)
            ?? throw new AgentToolException($"{slide}쪽에서 도형 '{address}'를 찾지 못했습니다. list_shapes로 주소를 확인하세요.");
    }

    /// <summary>box 인자: {left,top,width,height} 또는 도형 주소. 주소면 그 도형 자리와 주소를 돌려준다.</summary>
    private async Task<(BoxF Box, string? Address)> ResolveBoxAsync(string path, int slide, JsonElement box, CancellationToken ct)
    {
        if (box.ValueKind == JsonValueKind.String)
        {
            string address = box.GetString()!.Trim();
            var e = await FindShapeAsync(path, slide, address, ct);
            return (new BoxF(e.Left, e.Top, e.Width, e.Height), e.Address);
        }
        if (box.ValueKind != JsonValueKind.Object) throw new AgentToolException("box는 {left,top,width,height} 또는 도형 주소입니다.");
        float Need(string n) => ToolArgs.OptionalFloat(box, n) ?? throw new AgentToolException($"box.{n}가 필요합니다.");
        var b = new BoxF(Need("left"), Need("top"), Need("width"), Need("height"));
        if (b.Width <= 0 || b.Height <= 0) throw new AgentToolException("box의 width·height는 0보다 커야 합니다.");
        return (b, null);
    }

    // ── 스키마 ──────────────────────────────────────────────

    private static JsonObject Merge(JsonObject a, JsonObject b)
    {
        foreach (var (k, v) in b.ToList()) { b.Remove(k); a[k] = v; }
        return a;
    }

    private void Add(string name, string description, JsonObject schema,
        Func<JsonElement, CancellationToken, Task<McpToolResult>> call, params string[] required)
    {
        if (required.Length > 0) schema["required"] = new JsonArray(required.Select(r => (JsonNode)r).ToArray());
        _tools[name] = new McpTool(name, description, schema, call);
    }

    private static JsonObject Schema(params (string Name, string Type, string Description)[] properties)
    {
        var props = new JsonObject();
        foreach (var (name, type, description) in properties)
            props[name] = new JsonObject { ["type"] = type, ["description"] = description };
        return new JsonObject { ["type"] = "object", ["properties"] = props };
    }

    private static JsonObject WithFile(JsonObject schema, params string[] required)
    {
        if (schema["properties"] is not JsonObject props) schema["properties"] = props = new JsonObject();
        props["file"] = new JsonObject { ["type"] = "string", ["description"] = FileDescription };
        props["show_window"] = new JsonObject
        {
            ["type"] = "boolean",
            ["description"] = "대상 문서가 창 없이 PowerPoint 메모리에만 남아 있으면(창을 닫은 뒤) 창을 다시 띄우고 이어서 처리한다",
        };
        if (required.Length > 0) schema["required"] = new JsonArray(required.Select(r => (JsonNode)r).ToArray());
        return schema;
    }
}
