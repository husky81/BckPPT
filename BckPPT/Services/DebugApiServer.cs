using System.IO;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Windows.Threading;
using BckPPT.Models;
using BckPPT.ViewModels;

namespace BckPPT.Services;

/// <summary>localhost:38700 HTTP API — ViewModel 직접 제어로 디버깅 루프 지원.</summary>
public class DebugApiServer
{
    private readonly BckPptViewModel _vm;
    private readonly Dispatcher _dispatcher;
    private readonly SpdxService _spdx;
    private HttpListener? _listener;
    private CancellationTokenSource _cts = new();

    // SPDX 추출 결과 캐시 (apply 시 entries 참조용)
    private List<SpdxEntry>? _lastSpdxEntries;
    private int _lastSpdxSlide;

    private static readonly JsonSerializerOptions _json = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };

    public DebugApiServer(BckPptViewModel vm, Dispatcher dispatcher)
    {
        _vm = vm;
        _dispatcher = dispatcher;
        _spdx = new SpdxService(vm, new PptComService());
    }

    public void Start()
    {
        try
        {
            _listener = new HttpListener();
            _listener.Prefixes.Add("http://localhost:38700/");
            _listener.Start();
            Task.Run(ListenLoop);
            System.Diagnostics.Debug.WriteLine("[DebugAPI] Started on http://localhost:38700");
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[DebugAPI] Start failed: {ex.Message}");
        }
    }

    public void Stop()
    {
        _cts.Cancel();
        _listener?.Stop();
    }

    private async Task ListenLoop()
    {
        while (!_cts.IsCancellationRequested)
        {
            try
            {
                var ctx = await _listener!.GetContextAsync();
                _ = HandleRequest(ctx);
            }
            catch (HttpListenerException) { break; }
            catch (ObjectDisposedException) { break; }
        }
    }

    private async Task HandleRequest(HttpListenerContext ctx)
    {
        var path = ctx.Request.Url?.AbsolutePath ?? "";
        var method = ctx.Request.HttpMethod;
        System.Diagnostics.Debug.WriteLine($"[DebugAPI] {method} {path}");

        try
        {
            object result = (method, path) switch
            {
                ("GET", "/api/status") => await UI(() => GetStatus()),
                ("GET", "/api/slides") => await UI(() => GetSlides()),
                ("GET", "/api/tree")   => await UI(() => GetTree(ctx.Request.QueryString["slide"])),
                ("POST", "/api/select-slide") => await HandleSelectSlide(ctx),
                ("POST", "/api/move")         => await HandleMove(ctx),
                ("POST", "/api/set-role")     => await HandleSetRole(ctx),
                ("POST", "/api/save")         => await UI(() => { _vm.SaveTreeStructure(); return new { ok = true }; }),
                ("POST", "/api/refresh")      => await UI(() => { _vm.RefreshCommand.Execute(null); return new { ok = true }; }),
                ("GET",  "/api/spdx")           => await HandleSpdxExtract(ctx),
                ("POST", "/api/spdx/apply")    => await HandleSpdxApply(ctx),
                ("POST", "/api/spdx/undo")     => await HandleSpdxUndo(ctx),
                ("POST", "/api/spdx/apply-multi") => await HandleSpdxApplyMulti(ctx),
                ("GET",  "/api/images")          => await HandleImageList(ctx),
                ("POST", "/api/images/replace")  => await HandleImageReplace(ctx),
                ("POST", "/api/set-shape-text")  => await HandleSetShapeText(ctx),
                // ── prefix 라우팅 ───────────────────────────────────────────
                ("GET",  var p) when p.StartsWith("/api/node/")
                    => await HandleNodeDetail(ctx, p),
                ("GET",  var p) when p.StartsWith("/api/slide/") && p.EndsWith("/summary")
                    => await HandleSlideSummary(ctx, p),
                // ────────────────────────────────────────────────────────────
                _ => new { error = $"unknown: {method} {path}" },
            };

            await WriteJson(ctx.Response, result);
        }
        catch (Exception ex)
        {
            await WriteJson(ctx.Response, new { error = ex.Message }, 500);
        }
    }

    // ── GET handlers ─────────────────────────────────────────────

    private object GetStatus() => new
    {
        presentation = _vm.SelectedPresentation?.Name,
        slideCount = _vm.SlideItems.Count(s => !s.IsAll),
        nodeCount = FlatCount(_vm.RootNodes),
    };

    private object GetSlides() => _vm.SlideItems
        .Where(s => !s.IsAll)
        .Select(s => new { s.SlideIndex, s.Title, s.ShapeCount })
        .ToList();

    private object GetTree(string? slideFilter)
    {
        var nodes = _vm.RootNodes.AsEnumerable();
        if (int.TryParse(slideFilter, out int slide))
            nodes = nodes.Where(n => n.Shape.SlideIndex == slide);
        return new { nodes = nodes.Select(MapNode).ToList() };
    }

    // ── POST handlers ────────────────────────────────────────────

    private async Task<object> HandleSelectSlide(HttpListenerContext ctx)
    {
        var body = await ReadBody(ctx);
        if (!body.TryGetProperty("slide", out var slideProp)) return new { error = "missing 'slide'" };
        int slideIdx = slideProp.GetInt32();

        return await UI(() =>
        {
            var slide = _vm.SlideItems.FirstOrDefault(s => s.SlideIndex == slideIdx);
            if (slide is null) return (object)new { error = $"slide {slideIdx} not found" };
            _vm.SelectedSlide = slide;
            return new { ok = true, slide = slideIdx };
        });
    }

    private async Task<object> HandleMove(HttpListenerContext ctx)
    {
        var body = await ReadBody(ctx);

        var nodeIds = body.TryGetProperty("nodeIds", out var nArr)
            ? nArr.EnumerateArray().Select(e => e.GetString()!).ToList()
            : [];
        var targetId = body.TryGetProperty("targetId", out var tProp) ? tProp.GetString() : null;
        var posStr = body.TryGetProperty("position", out var pProp) ? pProp.GetString() : "below";

        if (nodeIds.Count == 0 || targetId is null)
            return new { error = "missing nodeIds or targetId" };

        var pos = posStr switch
        {
            "above" => BckPptViewModel.DropPosition.Above,
            "inside" => BckPptViewModel.DropPosition.Inside,
            _ => BckPptViewModel.DropPosition.Below,
        };

        return await UI(() =>
        {
            var target = _vm.FindNode(targetId);
            if (target is null) return (object)new { error = $"target '{targetId}' not found" };

            var moved = new List<string>();
            foreach (var id in nodeIds)
            {
                var source = _vm.FindNode(id);
                if (source is null) continue;
                _vm.MoveNode(source, target, pos);
                moved.Add(id);
            }

            _vm.SaveTreeStructure();
            return new { ok = true, moved, target = targetId, position = posStr };
        });
    }

    private async Task<object> HandleSetRole(HttpListenerContext ctx)
    {
        var body = await ReadBody(ctx);

        var nodeIds = body.TryGetProperty("nodeIds", out var nArr)
            ? nArr.EnumerateArray().Select(e => e.GetString()!).ToList()
            : [];
        var roleStr = body.TryGetProperty("role", out var rProp) ? rProp.GetString() : null;

        if (nodeIds.Count == 0 || roleStr is null)
            return new { error = "missing nodeIds or role" };

        if (!Enum.TryParse<ShapeRole>(roleStr, true, out var role))
            return new { error = $"unknown role '{roleStr}'" };

        return await UI(() =>
        {
            var changed = new List<string>();
            foreach (var id in nodeIds)
            {
                var node = _vm.FindNode(id);
                if (node is null) continue;
                node.SetRole(role);
                changed.Add(id);
            }
            _vm.SaveTreeStructure();
            return new { ok = true, role = roleStr, changed };
        });
    }

    // ── SPDX handlers ─────────────────────────────────────────────

    private async Task<object> HandleSpdxExtract(HttpListenerContext ctx)
    {
        if (_vm.SelectedPresentation is null)
            return new { error = "프레젠테이션이 선택되지 않았습니다." };

        var slideStr = ctx.Request.QueryString["slide"];
        if (!int.TryParse(slideStr, out int slideIndex))
            return new { error = "slide 파라미터가 필요합니다. 예: /api/spdx?slide=2" };

        var path = _vm.SelectedPresentation.Path;
        var (tsv, entries) = await Task.Run(() => _spdx.ExtractSpdx(path, slideIndex));

        _lastSpdxEntries = entries;
        _lastSpdxSlide = slideIndex;

        return new
        {
            spdx = tsv,
            slideIndex,
            shapeCount = entries.Count(e => e.Type.StartsWith("shape:")),
            lineCount = entries.Count(e => e.Type.StartsWith("line")),
        };
    }

    private async Task<object> HandleSpdxApply(HttpListenerContext ctx)
    {
        if (_vm.SelectedPresentation is null)
            return new { error = "프레젠테이션이 선택되지 않았습니다." };

        var body = await ReadBody(ctx);
        if (!body.TryGetProperty("slide", out var slideProp))
            return new { error = "missing 'slide'" };
        if (!body.TryGetProperty("patch", out var patchProp))
            return new { error = "missing 'patch'" };

        int slideIndex = slideProp.GetInt32();
        string patchTsv = patchProp.GetString() ?? "";
        bool smartTrim = body.TryGetProperty("smartTrim", out var stProp) && stProp.GetBoolean();
        System.Diagnostics.Debug.WriteLine($"[SPDX] patch length={patchTsv.Length}, smartTrim={smartTrim}");

        // entries가 없으면 먼저 추출
        if (_lastSpdxEntries is null || _lastSpdxSlide != slideIndex)
        {
            var path = _vm.SelectedPresentation.Path;
            var (_, entries) = await Task.Run(() => _spdx.ExtractSpdx(path, slideIndex));
            _lastSpdxEntries = entries;
            _lastSpdxSlide = slideIndex;
        }

        var result = await Task.Run(() =>
            _spdx.ApplyPatch(_vm.SelectedPresentation.Path, slideIndex, patchTsv, _lastSpdxEntries, smartTrim));

        // 적용 후 UI 새로고침
        if (result.Ok)
            await UI(() => { _vm.RefreshCommand.Execute(null); return true; });

        return new
        {
            result.Ok,
            result.Message,
            changes = result.Changes.Select(c => new
            {
                c.ShapeName,
                originalPreview = c.OriginalText.Length > 50
                    ? c.OriginalText[..50] + "..." : c.OriginalText,
                newPreview = c.NewText.Length > 50
                    ? c.NewText[..50] + "..." : c.NewText,
            }),
            result.Violations,
            overflowEntries = result.OverflowEntries.Select(o => new
            {
                o.Id,
                o.ShapeName,
                o.MaxHW,
                o.ActualHW,
                o.Text,
            }),
            undoAvailable = _spdx.GetUndoCount(_vm.SelectedPresentation.Path, slideIndex),
        };
    }

    private async Task<object> HandleSpdxUndo(HttpListenerContext ctx)
    {
        if (_vm.SelectedPresentation is null)
            return new { error = "프레젠테이션이 선택되지 않았습니다." };

        var body = await ReadBody(ctx);
        if (!body.TryGetProperty("slide", out var slideProp))
            return new { error = "missing 'slide'" };

        int slideIndex = slideProp.GetInt32();
        var result = await Task.Run(() =>
            _spdx.Undo(_vm.SelectedPresentation.Path, slideIndex));

        if (result.Ok)
            await UI(() => { _vm.RefreshCommand.Execute(null); return true; });

        return new { result.Ok, result.Message, result.RemainingUndoCount };
    }

    private async Task<object> HandleSpdxApplyMulti(HttpListenerContext ctx)
    {
        if (_vm.SelectedPresentation is null)
            return new { error = "프레젠테이션이 선택되지 않았습니다." };

        var body = await ReadBody(ctx);
        if (!body.TryGetProperty("slides", out var slidesArr))
            return new { error = "missing 'slides'" };

        bool smartTrim = body.TryGetProperty("smartTrim", out var stProp) && stProp.GetBoolean();
        var presentationPath = _vm.SelectedPresentation.Path;
        var results = new List<object>();

        foreach (var slideObj in slidesArr.EnumerateArray())
        {
            if (!slideObj.TryGetProperty("slide", out var sProp) ||
                !slideObj.TryGetProperty("patch", out var pProp))
            {
                results.Add(new { slide = 0, ok = false, error = "missing slide or patch" });
                continue;
            }

            int slideIndex = sProp.GetInt32();
            string patchTsv = pProp.GetString() ?? "";

            // 각 슬라이드별 entries 추출
            var (_, entries) = await Task.Run(() => _spdx.ExtractSpdx(presentationPath, slideIndex));
            var result = await Task.Run(() =>
                _spdx.ApplyPatch(presentationPath, slideIndex, patchTsv, entries, smartTrim));

            results.Add(new
            {
                slide = slideIndex,
                ok = result.Ok,
                message = result.Message,
                violations = result.Violations,
                overflowEntries = result.OverflowEntries.Select(o => new
                {
                    o.Id, o.ShapeName, o.MaxHW, o.ActualHW, o.Text,
                }),
            });
        }

        // 하나라도 성공하면 UI 새로고침
        if (results.Any(r => ((dynamic)r).ok))
            await UI(() => { _vm.RefreshCommand.Execute(null); return true; });

        return new { results };
    }

    // ── Image handlers ────────────────────────────────────────────

    private async Task<object> HandleImageList(HttpListenerContext ctx)
    {
        if (_vm.SelectedPresentation is null)
            return new { error = "프레젠테이션이 선택되지 않았습니다." };

        var slideStr = ctx.Request.QueryString["slide"];
        if (!int.TryParse(slideStr, out int slideIndex))
            return new { error = "slide 파라미터가 필요합니다." };

        var path = _vm.SelectedPresentation.Path;
        var images = await Task.Run(() => new PptComService().GetImageShapes(path, slideIndex));

        return new { images };
    }

    private async Task<object> HandleImageReplace(HttpListenerContext ctx)
    {
        if (_vm.SelectedPresentation is null)
            return new { error = "프레젠테이션이 선택되지 않았습니다." };

        var body = await ReadBody(ctx);
        if (!body.TryGetProperty("slide", out var slideProp))
            return new { error = "missing 'slide'" };
        if (!body.TryGetProperty("replacements", out var repArr))
            return new { error = "missing 'replacements'" };

        int slideIndex = slideProp.GetInt32();
        var presentationPath = _vm.SelectedPresentation.Path;
        var com = new PptComService();
        var results = new List<object>();

        var reps = repArr.EnumerateArray().ToList();
        foreach (var rep in reps)
        {
            float left = (float)rep.GetProperty("left").GetDouble();
            float top = (float)rep.GetProperty("top").GetDouble();
            string imageUrl = rep.GetProperty("imageUrl").GetString() ?? "";
            if (string.IsNullOrEmpty(imageUrl)) continue;

            try
            {
                string imagePath;
                bool isTempFile = false;

                if (File.Exists(imageUrl))
                {
                    // 로컬 파일 경로
                    imagePath = imageUrl;
                }
                else
                {
                    // URL → 다운로드
                    using var http = new HttpClient();
                    http.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                    var bytes = await http.GetByteArrayAsync(imageUrl);
                    var ext = imageUrl.Contains(".png", StringComparison.OrdinalIgnoreCase) ? ".png" : ".jpg";
                    imagePath = Path.Combine(Path.GetTempPath(), $"bckppt_{Guid.NewGuid():N}{ext}");
                    await File.WriteAllBytesAsync(imagePath, bytes);
                    isTempFile = true;
                }

                await Task.Run(() => com.ReplaceImage(presentationPath, slideIndex, left, top, imagePath));
                results.Add(new { left, top, ok = true });

                if (isTempFile) try { File.Delete(imagePath); } catch { }
            }
            catch (Exception ex)
            {
                results.Add(new { left, top, ok = false, error = ex.Message });
            }
        }

        // UI 새로고침
        await UI(() => { _vm.RefreshCommand.Execute(null); return true; });

        return new { ok = true, results };
    }

    // ── set-shape-text ────────────────────────────────────────────

    /// <summary>
    /// POST /api/set-shape-text
    /// Body: { "slide": 1, "shapeName": "Table 1 [R2C2]", "text": "{대표자}" }
    /// 표 셀(shapeName에 [RxCy] 포함)과 일반 텍스트 상자 모두 지원.
    /// </summary>
    private async Task<object> HandleSetShapeText(HttpListenerContext ctx)
    {
        if (_vm.SelectedPresentation is null)
            return new { error = "프레젠테이션이 선택되지 않았습니다." };

        var body = await ReadBody(ctx);
        if (!body.TryGetProperty("slide",     out var slideProp))     return new { error = "missing 'slide'" };
        if (!body.TryGetProperty("shapeName", out var nameProp))      return new { error = "missing 'shapeName'" };
        if (!body.TryGetProperty("text",      out var textProp))      return new { error = "missing 'text'" };

        int    slide     = slideProp.GetInt32();
        string shapeName = nameProp.GetString() ?? "";
        string text      = textProp.GetString() ?? "";
        string path      = _vm.SelectedPresentation.Path;

        await Task.Run(() => new PptComService().SetShapeText(path, slide, shapeName, text));
        await UI(() => { _vm.RefreshCommand.Execute(null); return true; });

        return new { ok = true, slide, shapeName, text };
    }

    // ── node detail / slide summary ──────────────────────────────

    private async Task<object> HandleNodeDetail(HttpListenerContext ctx, string path)
    {
        var encoded = path["/api/node/".Length..];
        var nodeId  = Uri.UnescapeDataString(encoded);

        return await UI(() =>
        {
            var node = _vm.FindNode(nodeId);
            if (node is null)
                return (object)new { error = $"node '{nodeId}' not found" };
            return MapNodeDetail(node);
        });
    }

    private async Task<object> HandleSlideSummary(HttpListenerContext ctx, string path)
    {
        // path = "/api/slide/2/summary"
        var seg = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (seg.Length != 4 || !int.TryParse(seg[2], out int slideNum))
            return new { error = "invalid slide summary path" };

        return await UI(() =>
        {
            var slideItem = _vm.SlideItems.FirstOrDefault(s => s.SlideIndex == slideNum);
            if (slideItem is null)
                return (object)new { error = $"slide {slideNum} not found" };

            var roots = _vm.RootNodes
                .Where(n => n.Shape.SlideIndex == slideNum)
                .ToList();

            var shapes = new List<object>();
            FlattenNodes(roots, 0, shapes);

            return new { slide = slideNum, title = slideItem.Title, shapes };
        });
    }

    private static object MapNodeDetail(PptNode node)
    {
        var s = node.Shape;

        int? maxLines = (s.LineHeight > 0f && s.BoxHeight > 0f)
            ? (int?)Math.Floor((double)s.BoxHeight / s.LineHeight)
            : null;

        int? remainingLines = maxLines.HasValue
            ? maxLines.Value - s.LineCount
            : null;

        string layoutHint = s.IsImage    ? "image"
            : s.HasBullet               ? "bullet_list"
            : s.LineCount > 1           ? "paragraph"
            :                             "single_line";

        return new
        {
            id          = node.NodeId,
            slide       = s.SlideIndex,
            shapeName   = s.ShapeName,
            role        = s.Role.ToString(),
            roleName    = s.Role.ToDisplayName(),
            layoutHint,
            text        = s.FullText,
            lineCount   = s.LineCount,
            maxLines,
            remainingLines,
            lineHeight  = s.LineHeight,
            fontSize    = s.FontSize,
            boundHeight = s.BoundHeight,
            boxHeight   = s.BoxHeight,
            isOverflow  = s.IsOverflow,
            hasBullet   = s.HasBullet,
            maxLineHw   = s.MaxLineHw,
            parent      = node.Parent?.NodeId,
            children    = node.Children.Select(c => c.NodeId).ToList(),
        };
    }

    private static void FlattenNodes(IEnumerable<PptNode> nodes, int depth, List<object> result)
    {
        foreach (var node in nodes)
        {
            var s = node.Shape;

            int? maxLines = (s.LineHeight > 0f && s.BoxHeight > 0f)
                ? (int?)Math.Floor((double)s.BoxHeight / s.LineHeight)
                : null;

            string layoutHint = s.IsImage    ? "image"
                : s.HasBullet               ? "bullet_list"
                : s.LineCount > 1           ? "paragraph"
                :                             "single_line";

            result.Add(new
            {
                id             = node.NodeId,
                role           = s.Role.ToString(),
                roleName       = s.Role.ToDisplayName(),
                layoutHint,
                lineCount      = s.LineCount,
                maxLines,
                remainingLines = maxLines.HasValue ? maxLines.Value - s.LineCount : (int?)null,
                fontSize       = s.FontSize,
                maxLineHw      = s.MaxLineHw,
                isOverflow     = s.IsOverflow,
                hasBullet      = s.HasBullet,
                depth,
                parent         = node.Parent?.NodeId,
                children       = node.Children.Select(c => c.NodeId).ToList(),
            });

            FlattenNodes(node.Children, depth + 1, result);
        }
    }

    // ── helpers ───────────────────────────────────────────────────

    private async Task<T> UI<T>(Func<T> func)
        => await _dispatcher.InvokeAsync(func);

    private static object MapNode(PptNode node) => new
    {
        id = node.NodeId,
        role = node.Shape.Role.ToAbbr(),
        roleName = node.Shape.Role.ToDisplayName(),
        text = node.Shape.TextPreview,
        slide = node.Shape.SlideIndex,
        visible = node.IsVisible,
        children = node.Children.Select(MapNode).ToList(),
    };

    private static int FlatCount(IEnumerable<PptNode> nodes)
        => nodes.Sum(n => 1 + FlatCount(n.Children));

    private static async Task<JsonElement> ReadBody(HttpListenerContext ctx)
    {
        using var reader = new StreamReader(ctx.Request.InputStream, System.Text.Encoding.UTF8);
        var json = await reader.ReadToEndAsync();
        return JsonSerializer.Deserialize<JsonElement>(json);
    }

    private static async Task WriteJson(HttpListenerResponse resp, object data, int status = 200)
    {
        resp.StatusCode = status;
        resp.ContentType = "application/json; charset=utf-8";
        var json = JsonSerializer.Serialize(data, _json);
        var bytes = System.Text.Encoding.UTF8.GetBytes(json);
        resp.ContentLength64 = bytes.Length;
        await resp.OutputStream.WriteAsync(bytes);
        resp.Close();
    }
}
