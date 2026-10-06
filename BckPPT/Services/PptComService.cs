using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using BckPPT.Models;
using Microsoft.Win32;

namespace BckPPT.Services;

public class PptComService : IPptComService
{
    private const int MsoTrue  = -1;
    private const int MsoFalse =  0;

    // ── Marshal.GetActiveObject (.NET 5+ P/Invoke) ───────────────

    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        ref Guid rclsid, IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    private static object OleGetActiveObject(string progId)
    {
        int hr = CLSIDFromProgID(progId, out var clsid);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        hr = GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
        return obj;
    }

    // ── COM 안전 해제 헬퍼 ────────────────────────────────────────

    /// <summary>presentation.Close() + ReleaseComObject를 각각 try-catch로 감싸 이중 해제 방지.</summary>
    private static void SafeClose(dynamic? prs)
    {
        if (prs is null) return;
        try { prs.Close(); }    catch { }
        try { Marshal.ReleaseComObject(prs); } catch { }
    }

    // ── STA 스레드 래퍼 (StaDispatcher 싱글턴 사용) ─────────────

    private static T RunOnSta<T>(Func<T> func)   => StaDispatcher.Instance.Invoke(func);
    private static void RunOnSta(Action action)   => StaDispatcher.Instance.Invoke(action);

    // ── 분석 헬퍼 ─────────────────────────────────────────────────

    /// <summary>FullText에서 줄별 반각 HW를 계산하여 최대값 반환.</summary>
    private static int CalcMaxLineHw(string fullText)
    {
        if (string.IsNullOrEmpty(fullText)) return 0;
        int max = 0;
        foreach (var line in fullText.Split('\r'))
        {
            int hw = PptSlideInfo.CalcHalfWidth(line.TrimEnd('\n', '\v'));
            if (hw > max) max = hw;
        }
        return max;
    }

    /// <summary>
    /// slide.Shapes 또는 group.GroupItems를 재귀적으로 평탄화하여 개별 shape를 반환.
    /// msoGroup(6) shape는 직접 반환하지 않고 GroupItems를 재귀 탐색.
    /// </summary>
    private static IEnumerable<dynamic> FlattenShapes(dynamic shapeCollection)
    {
        foreach (dynamic shape in shapeCollection)
        {
            int t = 0;
            try { t = (int)shape.Type; } catch { }
            if (t == 6) // msoGroup
            {
                IEnumerable<dynamic> children;
                try { children = FlattenShapes(shape.GroupItems); }
                catch { continue; }
                foreach (var child in children)
                    yield return child;
            }
            else
            {
                yield return shape;
            }
        }
    }

    /// <summary>FlattenShapes와 같은 순서로 펼치되 각 도형의 조상 그룹 이름을 함께 돌려준다.</summary>
    private static List<(dynamic Shape, IReadOnlyList<string> Path)> FlattenShapesWithPath(
        dynamic shapeCollection, IReadOnlyList<string>? parent = null)
    {
        parent ??= Array.Empty<string>();
        var result = new List<(dynamic, IReadOnlyList<string>)>();
        foreach (dynamic shape in shapeCollection)
        {
            int t = 0;
            try { t = (int)shape.Type; } catch { }
            if (t == 6) // msoGroup
            {
                var path = parent.Append((string)shape.Name).ToArray();
                try { result.AddRange(FlattenShapesWithPath(shape.GroupItems, path)); }
                catch { }
            }
            else
            {
                result.Add((shape, parent));
            }
        }
        return result;
    }

    /// <summary><see cref="ShapeAddress"/> 규약으로 도형을 찾는다. 그룹 안 도형과 겹치는 이름도 구분한다.</summary>
    private static dynamic? ResolveShape(dynamic slide, ShapeAddress address)
    {
        List<(dynamic Shape, IReadOnlyList<string> Path)> shapes = FlattenShapesWithPath(slide.Shapes);
        if (address.Id is int id)
        {
            foreach (var (shape, _) in shapes)
            {
                try { if ((int)shape.Id == id) return shape; } catch { }
            }
            return null;
        }
        var candidates = shapes
            .Select(s => (s.Path, (string)s.Shape.Name))
            .ToList();
        int index = address.Select(candidates);
        return index < 0 ? null : shapes[index].Shape;
    }

    private static dynamic RequireShape(dynamic slide, int slideIndex, string address)
        => ResolveShape(slide, ShapeAddress.Parse(address))
           ?? throw new InvalidOperationException($"{slideIndex}쪽에서 도형 '{address}'를 찾지 못했습니다.");

    /// <summary>
    /// 주소(<see cref="ShapeAddress"/>)에 맞는 도형을 0개 또는 1개 담아 돌려준다.
    /// 기존 foreach/break 흐름에 그대로 끼워 그룹 안 도형과 겹치는 이름을 지원하기 위한 것.
    /// </summary>
    private static IEnumerable<dynamic> MatchingShape(dynamic slide, string address)
    {
        dynamic? shape = ResolveShape(slide, ShapeAddress.Parse(address));
        return shape is null ? Array.Empty<object>() : new object[] { shape };
    }

    /// <summary>TextFrame2의 첫 5문단 중 불릿이 하나라도 있으면 true.</summary>
    private static bool ReadHasBullet(dynamic tf)
    {
        try
        {
            dynamic paras = tf.TextRange.Paragraphs;
            int count = Math.Min((int)paras.Count, 5);
            for (int i = 1; i <= count; i++)
            {
                if ((int)paras(i).ParagraphFormat.Bullet.Visible == MsoTrue)
                    return true;
            }
        }
        catch { }
        return false;
    }

    // ── 공개 API ──────────────────────────────────────────────────

    public bool IsPowerPointRunning
    {
        get
        {
            try { OleGetActiveObject("PowerPoint.Application"); return true; }
            catch (COMException) { return false; }
        }
    }

    public IReadOnlyList<PptPresentation> GetOpenPresentations()
        => RunOnSta(() =>
        {
            dynamic? pptApp = TryGetActiveInstance();
            if (pptApp is null) return Array.Empty<PptPresentation>();

            var list = new List<PptPresentation>();
            foreach (dynamic prs in pptApp.Presentations)
            {
                try { list.Add(new PptPresentation((string)prs.Name, (string)prs.FullName, (int)prs.Slides.Count)); }
                catch { }
            }
            return (IReadOnlyList<PptPresentation>)list;
        });

    // ── 분석 ──────────────────────────────────────────────────────

    public PresentationAnalysis AnalyzePresentation(
        string presentationPath,
        IProgress<SlideAnalysisProgress>? slideProgress = null,
        CancellationToken cancellationToken = default)
        => RunOnSta(() =>
        {
            var rules  = PptRules.Current;
            dynamic pptApp = GetOrCreateInstance();

            // 이미 열려 있으면 그 인스턴스 사용 (미저장 변경 반영)
            // 열려 있지 않으면 읽기 전용으로 새로 열기
            dynamic? existing = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = existing is null;
            dynamic prs = existing
                ?? pptApp.Presentations.Open(presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                var slides = new List<SlideListItem>();
                var shapes = new List<PptSlideInfo>();
                var fonts  = new List<PptFontInfo>();

                int totalSlides = (int)prs.Slides.Count;

                // ── Throttle 도구 (STA 스레드 전용) ──────────────────
                var batchShapes = new List<PptSlideInfo>(20);
                var batchFonts  = new List<PptFontInfo>(20);
                var sw = Stopwatch.StartNew();

                void FlushBatch(int slideIdx)
                {
                    if (batchShapes.Count == 0 && batchFonts.Count == 0) return;
                    slideProgress?.Report(new SlideAnalysisProgress(
                        AnalysisProgressKind.ShapeBatch, slideIdx, totalSlides,
                        ShapeBatch: batchShapes.ToList(),
                        FontBatch:  batchFonts.ToList()));
                    batchShapes.Clear();
                    batchFonts.Clear();
                    sw.Restart();
                }

                foreach (dynamic slide in prs.Slides)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    int    slideIdx      = (int)slide.SlideIndex;
                    string slideTitle    = "";
                    int    overflowCount = 0;
                    int    shapeCount    = 0;

                    // ★ Skeleton: 슬라이드 진입 즉시 보고
                    slideProgress?.Report(new SlideAnalysisProgress(
                        AnalysisProgressKind.SlideStart, slideIdx, totalSlides));

                    foreach (dynamic shape in FlattenShapes(slide.Shapes))
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        // ── 표(Table) 처리 ──
                        try
                        {
                            if ((int)shape.HasTable == MsoTrue)
                            {
                                string tableName = (string)shape.Name;
                                dynamic table = shape.Table;
                                int rows = (int)table.Rows.Count;
                                int cols = (int)table.Columns.Count;

                                for (int r = 1; r <= rows; r++)
                                for (int c = 1; c <= cols; c++)
                                {
                                    try
                                    {
                                        dynamic cell  = table.Cell(r, c);
                                        dynamic ctf   = cell.Shape.TextFrame2;
                                        string  ctext = (string)ctf.TextRange.Text;
                                        string  cellName = $"{tableName} [R{r}C{c}]";

                                        PptFontInfo? cellFont = null;
                                        try
                                        {
                                            string cfn = (string)ctf.TextRange.Font.Name;
                                            float  csz = (float)ctf.TextRange.Font.Size;
                                            bool   cb  = (int)ctf.TextRange.Font.Bold == MsoTrue;
                                            string cp  = string.IsNullOrWhiteSpace(ctext) ? "(빈 셀)"
                                                : ctext[..Math.Min(40, ctext.Length)].Replace('\r', ' ').Replace('\n', ' ');
                                            cellFont = new PptFontInfo(slideIdx, cellName, cp,
                                                cfn, (float)Math.Round(csz, 1), cb,
                                                rules.AllowedFontsSet.Contains(cfn));
                                            fonts.Add(cellFont);
                                            batchFonts.Add(cellFont);
                                        }
                                        catch { }

                                        if (string.IsNullOrWhiteSpace(ctext))
                                        {
                                            var emptyCell = new PptSlideInfo(slideIdx, cellName, "", "(빈 셀)",
                                                0f, 0f, 0, false) { IsEmpty = true };
                                            shapes.Add(emptyCell);
                                            batchShapes.Add(emptyCell);
                                            continue;
                                        }

                                        float cbH = 0f, ccH = 0f; int cln = 0; float cfs = 0f;
                                        try { cbH = (float)Math.Round((float)ctf.TextRange.BoundHeight, 1); } catch { }
                                        try { ccH = (float)Math.Round((float)cell.Shape.Height, 1); }         catch { }
                                        try { cln = (int)ctf.TextRange.Lines.Count; }                         catch { }
                                        try { cfs = (float)Math.Round((float)ctf.TextRange.Font.Size, 1); }   catch { }

                                        bool cov = cbH > ccH + rules.OverflowTolerance;
                                        if (cov) overflowCount++;

                                        string cprev = ctext[..Math.Min(60, ctext.Length)]
                                            .Replace('\r', ' ').Replace('\n', ' ');
                                        var cellShape = new PptSlideInfo(slideIdx, cellName, ctext, cprev,
                                            cbH, ccH, cln, cov, cfs)
                                        {
                                            MaxLineHw = CalcMaxLineHw(ctext),
                                            HasBullet = ReadHasBullet(ctf),
                                        };
                                        shapes.Add(cellShape);
                                        batchShapes.Add(cellShape);
                                        shapeCount++;
                                    }
                                    catch { }

                                    // ★ Throttle 체크 (표 셀 루프 내)
                                    if (batchShapes.Count >= 20 || sw.ElapsedMilliseconds >= 33)
                                        FlushBatch(slideIdx);
                                }
                                continue;
                            }
                        }
                        catch { }

                        // ── 이미지 shape 처리 ──
                        int shapeType = 0;
                        try { shapeType = (int)shape.Type; } catch { }
                        if ((int)shape.HasTextFrame != MsoTrue)
                        {
                            // 이미지 shape 수집 (msoLinkedPicture=11, msoPicture=13)
                            if (shapeType == 13 || shapeType == 11)
                            {
                                string imgName = (string)shape.Name;
                                float imgH = 0f;
                                try { imgH = (float)Math.Round((float)shape.Height, 1); } catch { }
                                var imgShape = new PptSlideInfo(slideIdx, imgName, "", "(이미지)",
                                    imgH, imgH, 0, false) { IsImage = true };
                                shapes.Add(imgShape);
                                batchShapes.Add(imgShape);
                            }
                            continue;
                        }

                        // ── 일반 텍스트 shape 처리 ──

                        var    tf        = shape.TextFrame2;
                        string text      = (string)tf.TextRange.Text;
                        string shapeName = (string)shape.Name;
                        bool   isEmpty   = string.IsNullOrWhiteSpace(text);

                        try
                        {
                            string fontName = (string)tf.TextRange.Font.Name;
                            float  fontSz   = (float)tf.TextRange.Font.Size;
                            bool   isBold   = (int)tf.TextRange.Font.Bold == MsoTrue;
                            bool   allowed  = rules.AllowedFontsSet.Contains(fontName);
                            string fprev    = isEmpty ? "(빈 상자)"
                                : text[..Math.Min(40, text.Length)].Replace('\r', ' ').Replace('\n', ' ');
                            var fontInfo = new PptFontInfo(slideIdx, shapeName, fprev,
                                fontName, (float)Math.Round(fontSz, 1), isBold, allowed);
                            fonts.Add(fontInfo);
                            batchFonts.Add(fontInfo);
                        }
                        catch { }

                        if (isEmpty)
                        {
                            var emptyShape = new PptSlideInfo(slideIdx, shapeName, "", "(빈 상자)",
                                0f, (float)Math.Round((float)shape.Height, 1),
                                0, false) { IsEmpty = true };
                            shapes.Add(emptyShape);
                            batchShapes.Add(emptyShape);
                            continue;
                        }

                        shapeCount++;

                        if (string.IsNullOrEmpty(slideTitle))
                        {
                            if (shapeName.Contains("Title", StringComparison.OrdinalIgnoreCase) ||
                                shapeName.Contains("제목", StringComparison.Ordinal))
                                slideTitle = text.Split('\r', '\n')[0];
                        }
                        if (string.IsNullOrEmpty(slideTitle))
                            slideTitle = text.Split('\r', '\n')[0];

                        float boundH    = (float)tf.TextRange.BoundHeight;
                        float boxH      = (float)shape.Height;
                        int   lineCount = (int)tf.TextRange.Lines.Count;
                        bool  overflow  = boundH > boxH + rules.OverflowTolerance;
                        if (overflow) overflowCount++;

                        float fontSize = 0f;
                        try { fontSize = (float)Math.Round((float)tf.TextRange.Font.Size, 1); } catch { }

                        string preview = text[..Math.Min(60, text.Length)]
                            .Replace('\r', ' ').Replace('\n', ' ');

                        var shapeInfo = new PptSlideInfo(slideIdx, shapeName, text, preview,
                            (float)Math.Round(boundH, 1), (float)Math.Round(boxH, 1),
                            lineCount, overflow, fontSize)
                        {
                            MaxLineHw = CalcMaxLineHw(text),
                            HasBullet = ReadHasBullet(tf),
                        };
                        shapes.Add(shapeInfo);
                        batchShapes.Add(shapeInfo);

                        // ★ Throttle 체크 (shape 루프)
                        if (batchShapes.Count >= 20 || sw.ElapsedMilliseconds >= 33)
                            FlushBatch(slideIdx);
                    }

                    // ★ 슬라이드 마지막 batch flush
                    FlushBatch(slideIdx);

                    if (slideTitle.Length > 30)
                        slideTitle = slideTitle[..27] + "...";

                    var slideItem = new SlideListItem(slideIdx, slideTitle, shapeCount, overflowCount);
                    slides.Add(slideItem);

                    // ★ SlideCompleted: skeleton → 실제 항목
                    slideProgress?.Report(new SlideAnalysisProgress(
                        AnalysisProgressKind.SlideCompleted, slideIdx, totalSlides,
                        Slide: slideItem));
                }

                return new PresentationAnalysis(slides, shapes, fonts);
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    // ── Markdown 내보내기 ────────────────────────────────────────

    private sealed record MarkdownParagraph(string Text, int Indent, bool HasBullet);
    private sealed record MarkdownPart(
        string Name, float Top, float Left,
        IReadOnlyList<MarkdownParagraph>? Paragraphs = null,
        IReadOnlyList<IReadOnlyList<string>>? Table = null);

    /// <summary>
    /// 그룹 도형을 반드시 재귀적으로 순회한다. 이 순서는 PowerPoint의 그룹 항목 순서를
    /// 보존하므로 카드형 슬라이드의 제목·문제점·성과물이 뒤섞이지 않는다.
    /// </summary>
    public MarkdownExportResult ExportMarkdown(string presentationPath, string outputPath,
        CancellationToken cancellationToken = default)
        => RunOnSta(() =>
        {
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Markdown 출력 경로가 없습니다.", nameof(outputPath));

            outputPath = Path.GetFullPath(outputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            var assetDirectoryName = Path.GetFileNameWithoutExtension(outputPath) + "_assets";
            var assetDirectory = Path.Combine(Path.GetDirectoryName(outputPath)!, assetDirectoryName);
            Directory.CreateDirectory(assetDirectory);

            dynamic pptApp = GetOrCreateInstance();
            dynamic? existing = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = existing is null;
            dynamic prs = existing ?? pptApp.Presentations.Open(presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                var markdown = new StringBuilder();
                var warnings = new List<string>();
                markdown.AppendLine($"# {Path.GetFileNameWithoutExtension(presentationPath)}");
                markdown.AppendLine();
                markdown.AppendLine($"- 원본: `{Path.GetFileName(presentationPath)}`");
                markdown.AppendLine($"- 슬라이드 수: {(int)prs.Slides.Count}");
                markdown.AppendLine("- 생성: BckPPT PowerPoint 구조 추출");

                for (int slideIndex = 1; slideIndex <= (int)prs.Slides.Count; slideIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    dynamic slide = prs.Slides[slideIndex];
                    var parts = new List<MarkdownPart>();
                    CollectMarkdownParts(slide.Shapes, parts, warnings, slideIndex, cancellationToken);
                    var title = FindMarkdownTitle(parts, slideIndex);

                    markdown.AppendLine();
                    markdown.AppendLine("---");
                    markdown.AppendLine();
                    markdown.AppendLine($"## 슬라이드 {slideIndex:00}. {EscapeMarkdown(title)}");
                    markdown.AppendLine();

                    string imageName = $"slide-{slideIndex:00}.png";
                    string imagePath = Path.Combine(assetDirectory, imageName);
                    try
                    {
                        slide.Export(imagePath, "PNG", 1920, 1080);
                        markdown.AppendLine($"![슬라이드 {slideIndex:00} 미리보기](<{assetDirectoryName}/{imageName}>)");
                        markdown.AppendLine();
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"슬라이드 {slideIndex:00}: 렌더 이미지 생성 실패 ({ex.Message})");
                    }

                    AppendMarkdownParts(markdown, parts, title);
                }

                if (warnings.Count > 0)
                {
                    markdown.AppendLine();
                    markdown.AppendLine("---");
                    markdown.AppendLine();
                    markdown.AppendLine("## 추출 경고");
                    markdown.AppendLine();
                    foreach (var warning in warnings.Distinct())
                        markdown.AppendLine($"- {EscapeMarkdown(warning)}");
                }

                File.WriteAllText(outputPath, markdown.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                return new MarkdownExportResult(outputPath, (int)prs.Slides.Count, warnings.Count);
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    private static void CollectMarkdownParts(dynamic shapes, List<MarkdownPart> parts,
        List<string> warnings, int slideIndex, CancellationToken cancellationToken)
    {
        foreach (dynamic shape in shapes)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int shapeType = 0;
            try { shapeType = (int)shape.Type; } catch { }
            if (shapeType == 6) // msoGroup
            {
                try { CollectMarkdownParts(shape.GroupItems, parts, warnings, slideIndex, cancellationToken); }
                catch { warnings.Add($"슬라이드 {slideIndex:00}: 그룹 도형을 읽지 못했습니다."); }
                continue;
            }

            string name = "도형";
            float top = 0, left = 0;
            try { name = (string)shape.Name; } catch { }
            try { top = (float)shape.Top; left = (float)shape.Left; } catch { }

            try
            {
                if ((int)shape.HasTable == MsoTrue)
                {
                    dynamic table = shape.Table;
                    var rows = new List<IReadOnlyList<string>>();
                    for (int r = 1; r <= (int)table.Rows.Count; r++)
                    {
                        var row = new List<string>();
                        for (int c = 1; c <= (int)table.Columns.Count; c++)
                        {
                            string value = "";
                            try { value = ((string)table.Cell(r, c).Shape.TextFrame2.TextRange.Text).Trim(); } catch { }
                            row.Add(value.Replace('\r', ' ').Replace('\n', ' ').Replace("|", "\\|"));
                        }
                        rows.Add(row);
                    }
                    parts.Add(new MarkdownPart(name, top, left, Table: rows));
                    continue;
                }
            }
            catch { warnings.Add($"슬라이드 {slideIndex:00}: 표 '{name}'을 읽지 못했습니다."); }

            bool hasText = false;
            try { hasText = (int)shape.HasTextFrame == MsoTrue; } catch { }
            if (hasText)
            {
                var paragraphs = ReadMarkdownParagraphs(shape);
                if (paragraphs.Count > 0)
                    parts.Add(new MarkdownPart(name, top, left, paragraphs));
                continue;
            }

            // 그림·차트·SmartArt는 슬라이드 PNG가 보존한다. Markdown에는 명시적으로 경고를 남긴다.
            bool isVisual = shapeType is 11 or 13;
            try { isVisual |= (int)shape.HasChart == MsoTrue; } catch { }
            if (isVisual)
                warnings.Add($"슬라이드 {slideIndex:00}: '{name}'의 시각 요소는 슬라이드 미리보기 PNG에 보존했습니다.");
        }
    }

    private static List<MarkdownParagraph> ReadMarkdownParagraphs(dynamic shape)
    {
        var result = new List<MarkdownParagraph>();
        try
        {
            dynamic tf = shape.TextFrame2;
            int count = (int)tf.TextRange.Paragraphs.Count;
            for (int index = 1; index <= count; index++)
            {
                dynamic paragraph = tf.TextRange.Paragraphs(index, 1);
                string text = ((string)paragraph.Text).Trim('\r', '\n', '\v', ' ');
                if (string.IsNullOrWhiteSpace(text)) continue;
                int indent = 0;
                bool bullet = false;
                try { indent = Math.Max(0, (int)paragraph.ParagraphFormat.IndentLevel - 1); } catch { }
                try { bullet = (int)paragraph.ParagraphFormat.Bullet.Visible == MsoTrue; } catch { }
                result.Add(new MarkdownParagraph(text, indent, bullet));
            }
        }
        catch { }
        return result;
    }

    private static string FindMarkdownTitle(IReadOnlyList<MarkdownPart> parts, int slideIndex)
    {
        var candidate = parts
            .Where(p => p.Paragraphs is { Count: > 0 })
            .OrderBy(p => p.Top).ThenBy(p => p.Left)
            .SelectMany(p => p.Paragraphs!)
            .Select(p => p.Text)
            .FirstOrDefault(text => text.Length > 2 && !Regex.IsMatch(text, @"^(No\.\s*\d+|\d+)$", RegexOptions.IgnoreCase));
        return candidate ?? $"슬라이드 {slideIndex}";
    }

    private static void AppendMarkdownParts(StringBuilder markdown, IReadOnlyList<MarkdownPart> parts, string title)
    {
        string? pendingNumber = null;
        var pendingHeading = new List<string>();
        bool skippedTitle = false;

        void FlushHeading()
        {
            if (pendingNumber is null) return;
            var heading = string.Join(" ", pendingHeading).Trim();
            markdown.AppendLine(string.IsNullOrWhiteSpace(heading)
                ? $"### {EscapeMarkdown(pendingNumber)}"
                : $"### {EscapeMarkdown(pendingNumber)}. {EscapeMarkdown(heading)}");
            markdown.AppendLine();
            pendingNumber = null;
            pendingHeading.Clear();
        }

        foreach (var part in parts)
        {
            if (part.Table is { Count: > 0 } table)
            {
                FlushHeading();
                int columns = table.Max(row => row.Count);
                var normalized = table.Select(row => row.Concat(Enumerable.Repeat("", columns - row.Count)).ToList()).ToList();
                markdown.AppendLine("| " + string.Join(" | ", normalized[0]) + " |");
                markdown.AppendLine("| " + string.Join(" | ", Enumerable.Repeat("---", columns)) + " |");
                foreach (var row in normalized.Skip(1)) markdown.AppendLine("| " + string.Join(" | ", row) + " |");
                markdown.AppendLine();
                continue;
            }
            if (part.Paragraphs is null) continue;
            foreach (var paragraph in part.Paragraphs)
            {
                var text = paragraph.Text;
                if (!skippedTitle && string.Equals(text, title, StringComparison.Ordinal))
                {
                    skippedTitle = true;
                    continue;
                }
                if (Regex.IsMatch(text, @"^No\.\s*\d+", RegexOptions.IgnoreCase))
                {
                    FlushHeading();
                    pendingNumber = text;
                    continue;
                }
                if (pendingNumber is not null && !IsCardSectionLabel(text))
                {
                    pendingHeading.Add(text);
                    continue;
                }
                FlushHeading();
                // 카드 순번은 No. 표기와 중복되므로 본문에서는 생략한다.
                if (Regex.IsMatch(text, @"^\d{1,2}$")) continue;
                string indentation = new(' ', paragraph.Indent * 2);
                markdown.AppendLine($"{indentation}- {EscapeMarkdown(text)}");
            }
            markdown.AppendLine();
        }
        FlushHeading();
    }

    private static bool IsCardSectionLabel(string text)
        => text.Trim() is "문제점" or "도출 성과물" or "추진 내용" or "핵심 내용" or "기대효과" or "기대 효과";

    private static string EscapeMarkdown(string value)
        => value.Replace("\\", "\\\\").Replace("*", "\\*").Replace("_", "\\_").Replace("[", "\\[").Replace("]", "\\]");

    // ── AutoFit (상세 결과 반환) ──────────────────────────────────

    public Task<AutoFitResult> AutoFitAsync(string presentationPath,
        IProgress<string>? progress = null)
        => Task.Run(() => RunOnSta(() =>
        {
            var rules  = PptRules.Current;
            dynamic pptApp = GetOrCreateInstance();
            dynamic prs = pptApp.Presentations.Open(
                presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                var details = new List<AutoFitDetail>();
                foreach (dynamic slide in prs.Slides)
                {
                    int slideIdx = (int)slide.SlideIndex;
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((int)shape.HasTextFrame != MsoTrue) continue;
                        if (!IsOverflowing(shape, rules)) continue;

                        string shapeName = (string)shape.Name;
                        float  origSize  = 0f;
                        try { origSize = GetFontSize(shape); } catch { }

                        progress?.Report($"슬라이드 {slideIdx} · {shapeName}");
                        bool trimmed = ApplyAutoFit(shape, rules);
                        float finalSize = 0f;
                        try { finalSize = GetFontSize(shape); } catch { }

                        details.Add(new AutoFitDetail(slideIdx, shapeName,
                            (float)Math.Round(origSize, 1),
                            (float)Math.Round(finalSize, 1),
                            trimmed));
                    }
                }
                SaveUnlessKept(prs);
                return new AutoFitResult(details.Count, details);
            }
            finally { SafeClose(prs); }
        }));

    public Task FixShapeAsync(string presentationPath, int slideIndex, string shapeName,
        IProgress<string>? progress = null)
        => Task.Run(() => RunOnSta(() =>
        {
            var rules  = PptRules.Current;
            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs   = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs)
                prs = pptApp.Presentations.Open(presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                dynamic slide = prs!.Slides[slideIndex];
                foreach (dynamic shape in slide.Shapes)
                {
                    if ((string)shape.Name != shapeName) continue;
                    if (!IsOverflowing(shape, rules)) break;
                    progress?.Report($"슬라이드 {slideIndex} · {shapeName} 수정 중...");
                    ApplyAutoFit(shape, rules);
                    break;
                }
                SaveUnlessKept(prs);
            }
            finally { if (ownedByUs) SafeClose(prs); }
        }));

    // ── Navigate ──────────────────────────────────────────────────

    public Task NavigateToShapeAsync(string presentationPath, int slideIndex, string shapeName)
        => Task.Run(async () =>
        {
            bool success = RunOnSta(() => TryNavigate(presentationPath, slideIndex, shapeName));
            if (success) return;
            OpenPptxInPowerPoint(presentationPath);
            await Task.Delay(2500);
            RunOnSta(() => TryNavigate(presentationPath, slideIndex, shapeName));
        });

    public void NavigateToSlide(string presentationPath, int slideIndex)
        => RunOnSta(() =>
        {
            dynamic? pptApp = TryGetActiveInstance();
            if (pptApp is null) return;
            foreach (dynamic win in pptApp.Windows)
            {
                string winPath;
                try { winPath = (string)win.Presentation.FullName; } catch { continue; }
                if (!string.Equals(winPath, presentationPath, StringComparison.OrdinalIgnoreCase)) continue;
                pptApp.Activate();
                win.Activate();
                win.View.GotoSlide(slideIndex);
                return;
            }
        });

    public void OpenPptxInPowerPoint(string filePath)
    {
        // 세션이 창 없이 열어 둔 문서면 창만 띄운다. 파일을 다시 열면 읽기 전용 사본이 생긴다.
        bool shown = RunOnSta(() =>
        {
            if (TryGetActiveInstance() is not { } app || FindOpenPresentation(app, filePath) is not { } prs) return false;
            if ((int)prs.Windows.Count == 0) prs.NewWindow();
            else prs.Windows[1].Activate();
            return true;
        });
        if (!shown) Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
    }

    /// <summary>
    /// 슬라이드 번호와 제목만 빠르게 읽는다. AnalyzePresentation은 모든 도형을 측정해
    /// 31MB·38장 착수보고 PPT에서 45초 제한을 넘겼다(2026-10-06 실측).
    /// 제목 개체가 없으면 위쪽(상단 80pt 안)에서 가장 위에 있는 텍스트를 제목으로 본다.
    /// </summary>
    public IReadOnlyList<(int Index, string Title)> ListSlides(string presentationPath)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? existing = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = existing is null;
            dynamic prs = existing ?? pptApp.Presentations.Open(presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                var result = new List<(int, string)>();
                foreach (dynamic slide in prs.Slides)
                {
                    string title = "";
                    try
                    {
                        if ((int)slide.Shapes.HasTitle == MsoTrue)
                            title = (string)slide.Shapes.Title.TextFrame2.TextRange.Text;
                    }
                    catch { }
                    if (string.IsNullOrWhiteSpace(title))
                    {
                        float best = 80f;
                        foreach (dynamic shape in slide.Shapes)
                        {
                            try
                            {
                                if ((int)shape.HasTextFrame != MsoTrue || (float)shape.Top >= best) continue;
                                string text = (string)shape.TextFrame2.TextRange.Text;
                                if (string.IsNullOrWhiteSpace(text)) continue;
                                (best, title) = ((float)shape.Top, text);
                            }
                            catch { }
                        }
                    }
                    title = title.Replace('\r', ' ').Replace('\v', ' ').Trim();
                    result.Add(((int)slide.SlideIndex, title.Length <= 80 ? title : title[..80]));
                }
                return (IReadOnlyList<(int, string)>)result;
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    /// <summary>
    /// 문서를 창 없이 열어 두고, 돌려받은 핸들을 해제할 때 닫는다. 이미 열려 있던 문서는 건드리지 않는다.
    /// 다른 API는 열린 문서를 찾아 재사용하므로, 큰 파일에서 호출마다 열고 닫는 비용이 사라진다.
    /// </summary>
    public IDisposable KeepOpen(string presentationPath, bool readOnly)
    {
        bool opened = RunOnSta(() =>
        {
            dynamic app = GetOrCreateInstance();
            if (FindOpenPresentation(app, presentationPath) is not null) return false;
            app.Presentations.Open(presentationPath, readOnly ? MsoTrue : MsoFalse, MsoFalse, MsoFalse);
            return true;
        });
        if (opened && !readOnly) lock (KeptOpen) KeptOpen.Add(presentationPath);
        return new OpenHandle(opened ? presentationPath : null);
    }

    /// <summary>
    /// 세션이 열어 둔 쓰기용 문서. 편집마다 저장하지 않고 핸들을 놓을 때 한 번 저장한다.
    /// 27MB 착수보고 PPT를 카드 하나 채울 때마다 저장하니 한 번이 45초 COM 제한을 넘겼다(2026-10-06 실행 12).
    /// </summary>
    private static readonly HashSet<string> KeptOpen = new(StringComparer.OrdinalIgnoreCase);

    private static void SaveUnlessKept(dynamic prs)
    {
        string path = "";
        try { path = (string)prs.FullName; } catch { }
        lock (KeptOpen) if (KeptOpen.Contains(path)) return;
        prs.Save();
    }

    private sealed class OpenHandle(string? path) : IDisposable
    {
        private string? _path = path;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _path, null) is not { } path) return;
            bool deferred;
            lock (KeptOpen) deferred = KeptOpen.Remove(path);
            try
            {
                RunOnSta(() =>
                {
                    if (TryGetActiveInstance() is { } app && FindOpenPresentation(app, path) is { } prs)
                    {
                        if (deferred) prs.Save();   // 미뤄 둔 저장
                        // 사용자가 "결과 열어 보기"로 창을 띄웠다면 닫지 않는다.
                        if ((int)prs.Windows.Count == 0) SafeClose(prs);
                    }
                });
            }
            catch { }
        }
    }

    // ── 텍스트 치환 ───────────────────────────────────────────────

    public int ReplaceText(string presentationPath, string find, string replace,
        IReadOnlyList<(int slideIndex, string shapeName)>? targets = null)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs   = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs)
                prs = pptApp.Presentations.Open(presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                int count = 0;
                foreach (dynamic slide in prs!.Slides)
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((int)shape.HasTextFrame != MsoTrue) continue;
                        int    si = (int)slide.SlideIndex;
                        string sn = (string)shape.Name;
                        if (targets is not null &&
                            !targets.Any(t => t.slideIndex == si && t.shapeName == sn)) continue;

                        string text = (string)shape.TextFrame2.TextRange.Text;
                        if (!text.Contains(find, StringComparison.OrdinalIgnoreCase)) continue;

                        foreach (dynamic para in shape.TextFrame2.TextRange.Paragraphs)
                        {
                            try
                            {
                                int runCount = (int)para.Runs.Count;
                                for (int ri = 1; ri <= runCount; ri++)
                                {
                                    dynamic run = para.Runs[ri];
                                    string rt = (string)run.Text;
                                    if (rt.Contains(find, StringComparison.OrdinalIgnoreCase))
                                    {
                                        run.Text = rt.Replace(find, replace, StringComparison.OrdinalIgnoreCase);
                                        count++;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                SaveUnlessKept(prs);
                return count;
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    // ── 규칙 일괄 치환 ────────────────────────────────────────────

    public IReadOnlyList<(int SlideIndex, string ShapeName, string OldText, string NewText)>
        ApplyRules(string presentationPath, IReadOnlyList<ReplaceRule> rules)
        => RunOnSta(() =>
        {
            if (rules.Count == 0) return Array.Empty<(int, string, string, string)>();

            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs   = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs)
                prs = pptApp.Presentations.Open(presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                var changed = new List<(int, string, string, string)>();
                foreach (dynamic slide in prs!.Slides)
                {
                    int slideIdx = (int)slide.SlideIndex;
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((int)shape.HasTextFrame != MsoTrue) continue;
                        string shapeName = (string)shape.Name;
                        string oldText   = (string)shape.TextFrame2.TextRange.Text;
                        string newText   = oldText;

                        foreach (var rule in rules)
                        {
                            if (string.IsNullOrEmpty(rule.Find)) continue;
                            var cmp = rule.CaseSensitive
                                ? StringComparison.Ordinal
                                : StringComparison.OrdinalIgnoreCase;
                            if (!newText.Contains(rule.Find, cmp)) continue;

                            // run 단위 치환 (서식 보존)
                            foreach (dynamic para in shape.TextFrame2.TextRange.Paragraphs)
                            {
                                try
                                {
                                    int runCount = (int)para.Runs.Count;
                                    for (int ri = 1; ri <= runCount; ri++)
                                    {
                                        dynamic run = para.Runs[ri];
                                        string rt = (string)run.Text;
                                        if (rt.Contains(rule.Find, cmp))
                                            run.Text = rt.Replace(rule.Find, rule.Replace, cmp);
                                    }
                                }
                                catch { }
                            }
                            newText = newText.Replace(rule.Find, rule.Replace, cmp);
                        }

                        if (newText != oldText)
                            changed.Add((slideIdx, shapeName, oldText, newText));
                    }
                }
                if (changed.Count > 0) SaveUnlessKept(prs);
                return (IReadOnlyList<(int, string, string, string)>)changed;
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    // ── 글꼴 정규화 ───────────────────────────────────────────────

    public int NormalizeFonts(string presentationPath)
        => RunOnSta(() =>
        {
            var rules  = PptRules.Current;
            dynamic pptApp = GetOrCreateInstance();
            dynamic prs = pptApp.Presentations.Open(
                presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                int count = 0;
                foreach (dynamic slide in prs.Slides)
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((int)shape.HasTextFrame != MsoTrue) continue;
                        foreach (dynamic para in shape.TextFrame2.TextRange.Paragraphs)
                        {
                            try
                            {
                                int runCount = (int)para.Runs.Count;
                                for (int ri = 1; ri <= runCount; ri++)
                                {
                                    dynamic run      = para.Runs[ri];
                                    string  fontName = (string)run.Font.Name;
                                    if (!rules.AllowedFontsSet.Contains(fontName))
                                    {
                                        run.Font.Name = rules.FallbackMapCI
                                            .GetValueOrDefault(fontName, rules.DefaultFontFallback);
                                        count++;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                SaveUnlessKept(prs);
                return count;
            }
            finally { SafeClose(prs); }
        });

    // ── 이미지 조회 ─────────────────────────────────────────────

    public record ImageShapeInfo(string Name, int ZOrder, float Width, float Height, float Left, float Top);

    public IReadOnlyList<ImageShapeInfo> GetImageShapes(string presentationPath, int slideIndex)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? existing = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = existing is null;
            dynamic prs = existing
                ?? pptApp.Presentations.Open(presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                dynamic slide = prs.Slides[slideIndex];
                var result = new List<ImageShapeInfo>();
                foreach (dynamic shape in slide.Shapes)
                {
                    int sType = (int)shape.Type;
                    if (sType == 13 || sType == 11) // msoPicture, msoLinkedPicture
                    {
                        result.Add(new ImageShapeInfo(
                            (string)shape.Name,
                            (int)shape.ZOrderPosition,
                            (float)Math.Round((float)shape.Width, 1),
                            (float)Math.Round((float)shape.Height, 1),
                            (float)Math.Round((float)shape.Left, 1),
                            (float)Math.Round((float)shape.Top, 1)));
                    }
                }
                return (IReadOnlyList<ImageShapeInfo>)result;
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    // ── 이미지 교체 ─────────────────────────────────────────────

    /// <summary>위치(left, top)로 이미지 shape을 찾아 교체. zOrder 변동에 안전.</summary>
    public void ReplaceImage(string presentationPath, int slideIndex,
        float targetLeft, float targetTop, string newImagePath)
        => RunOnSta(() =>
        {
            newImagePath = Path.GetFullPath(newImagePath);
            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs)
                prs = pptApp.Presentations.Open(presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                dynamic slide = prs!.Slides[slideIndex];

                // 위치 기반으로 이미지 shape 찾기 (±2pt 허용)
                dynamic? target = null;
                foreach (dynamic shape in slide.Shapes)
                {
                    int sType = (int)shape.Type;
                    if (sType != 13 && sType != 11) continue;
                    float l = (float)shape.Left, t = (float)shape.Top;
                    if (Math.Abs(l - targetLeft) < 2f && Math.Abs(t - targetTop) < 2f)
                    {
                        target = shape;
                        break;
                    }
                }
                if (target is null) return;

                float left   = (float)target.Left;
                float top    = (float)target.Top;
                float width  = (float)target.Width;
                float height = (float)target.Height;

                target.Delete();

                dynamic newShape = slide.Shapes.AddPicture(
                    newImagePath, MsoFalse, MsoTrue, left, top, width, height);

                SaveUnlessKept(prs);
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    // ── Shape 텍스트 전체 교체 ────────────────────────────────────

    public void SetShapeText(string presentationPath, int slideIndex, string shapeName, string newText)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs   = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs)
                prs = pptApp.Presentations.Open(presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                dynamic slide = prs!.Slides[slideIndex];

                // 테이블 셀 이름 파싱: "Table 1 [R2C3]"
                var cellMatch = Regex.Match(shapeName, @"^(.+)\s+\[R(\d+)C(\d+)\]$");
                if (cellMatch.Success)
                {
                    string tableName = cellMatch.Groups[1].Value;
                    int    row       = int.Parse(cellMatch.Groups[2].Value);
                    int    col       = int.Parse(cellMatch.Groups[3].Value);
                    foreach (dynamic shape in MatchingShape(slide, shapeName))
                    {
                        if ((int)shape.HasTable != MsoTrue) break;
                        shape.Table.Cell(row, col).Shape.TextFrame2.TextRange.Text = newText;
                        break;
                    }
                }
                else
                {
                    foreach (dynamic shape in MatchingShape(slide, shapeName))
                    {
                        if ((int)shape.HasTextFrame != MsoTrue) break;

                        var    tf        = shape.TextFrame2;
                        var    tr        = tf.TextRange;
                        string origText  = (string)tr.Text;
                        var    newParas  = newText.Split('\r');
                        int    pCount    = (int)tr.Paragraphs.Count;

                        // ★ \v 포함 or 문단 수 불일치 → 전체 교체 fallback
                        if (origText.Contains('\v') || newParas.Length != pCount)
                        {
                            tr.Text = newText;
                            break;
                        }

                        for (int i = 0; i < pCount; i++)
                        {
                            dynamic para = tr.Paragraphs(i + 1, 1);

                            // 변경 없으면 COM 호출 생략
                            if ((string)para.Text == newParas[i]) continue;

                            // ── 백업 ──────────────────────────────────────────
                            var bulletVisible = para.ParagraphFormat.Bullet.Visible;
                            var bulletChar    = para.ParagraphFormat.Bullet.Character;
                            var indentLevel   = para.ParagraphFormat.IndentLevel;
                            var fontName      = para.Font.Name;
                            var fontSize      = para.Font.Size;

                            // ── 텍스트 교체 (순수 텍스트, \r 없음) ───────────
                            para.Text = newParas[i];

                            // ── 복원: Bullet 먼저, 다음 Indent (순서 필수) ───
                            para.ParagraphFormat.Bullet.Visible = bulletVisible;
                            if ((int)bulletVisible == MsoTrue)
                            {
                                try { para.ParagraphFormat.Bullet.Character = bulletChar; } catch { }
                            }
                            para.ParagraphFormat.IndentLevel = indentLevel;

                            // ── Font 최소 복원 (혼합 서식 초기화 방어) ────────
                            try { para.Font.Name = fontName; } catch { }
                            try { para.Font.Size = fontSize; } catch { }
                        }
                        break;
                    }
                }
                SaveUnlessKept(prs);
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    // ── 문단 단위 치환 ───────────────────────────────────────────

    /// <summary>
    /// 도형 이름으로 TextFrame2를 찾는다. 테이블 셀 이름("Table 1 [R2C3]")도 처리.
    /// SetShapeText·GetTextLineBreaks와 같은 규약을 쓴다.
    /// </summary>
    private static dynamic? ResolveTextFrame(dynamic slide, string shapeName)
    {
        var cellMatch = Regex.Match(shapeName, @"^(.+)\s+\[R(\d+)C(\d+)\]$");
        if (cellMatch.Success)
        {
            string tableName = cellMatch.Groups[1].Value;
            int    row       = int.Parse(cellMatch.Groups[2].Value);
            int    col       = int.Parse(cellMatch.Groups[3].Value);
            foreach (dynamic shape in MatchingShape(slide, shapeName))
            {
                if ((int)shape.HasTable != MsoTrue) break;
                return shape.Table.Cell(row, col).Shape.TextFrame2;
            }
            return null;
        }

        foreach (dynamic shape in MatchingShape(slide, shapeName))
        {
            if ((int)shape.HasTextFrame != MsoTrue) break;
            return shape.TextFrame2;
        }
        return null;
    }

    /// <summary>문단의 현재 글자색을 읽는다. 테마색이 걸려 있으면 함께 보관한다.</summary>
    private static ParagraphColorSnapshot ReadColor(
        dynamic para, int slideIndex, string shapeName, int paragraphIndex)
    {
        int rgb = 0, theme = 0;
        try { rgb   = (int)para.Font.Fill.ForeColor.RGB; }              catch { }
        try { theme = (int)para.Font.Fill.ForeColor.ObjectThemeColor; } catch { }
        return new ParagraphColorSnapshot(slideIndex, shapeName, paragraphIndex, rgb, theme);
    }

    public ParagraphEditResult SetParagraphTexts(
        string presentationPath,
        IReadOnlyList<ParagraphEdit> edits,
        bool markRed)
        => RunOnSta(() =>
        {
            var snapshots = new List<ParagraphColorSnapshot>();
            var failures  = new List<string>();

            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs   = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs)
                prs = pptApp.Presentations.Open(presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                // 도형 단위로 묶어 Presentations.Open/Save를 한 번만 타게 한다.
                foreach (var group in edits.GroupBy(e => (e.SlideIndex, e.ShapeName)))
                {
                    dynamic slide = prs!.Slides[group.Key.SlideIndex];
                    dynamic? tf   = ResolveTextFrame(slide, group.Key.ShapeName);
                    if (tf is null)
                    {
                        failures.Add($"{group.Key.SlideIndex}쪽 [{group.Key.ShapeName}] — 도형을 찾지 못함");
                        continue;
                    }

                    dynamic tr = tf.TextRange;

                    // 적용 순서가 곧 인덱스 정합성이다.
                    //   1) 교체 — 문단 수를 바꾸지 않으므로 앞에서부터 해도 안전하다.
                    //   2) 삭제·삽입 — 뒤에서부터 한 번에 처리한다.
                    // 삭제와 삽입을 따로 돌리면 "p2 삭제 + p3 뒤 삽입"처럼 섞였을 때
                    // 뒤 작업이 밀린 인덱스를 보게 된다. 큰 인덱스부터 내려오면
                    // 아직 처리하지 않은 작업의 위치는 항상 그대로다.
                    var ordered = group
                        .Where(e => e.Kind == ParagraphEditKind.Replace)
                        .OrderBy(e => e.ParagraphIndex)
                        .Concat(group.Where(e => e.Kind != ParagraphEditKind.Replace)
                                     .OrderByDescending(e => e.ParagraphIndex));

                    foreach (var edit in ordered)
                    {
                        int paraNum = (int)tr.Paragraphs.Count;

                        // 삽입의 인덱스 0은 "맨 앞에 넣기"를 뜻한다.
                        int lowest = edit.Kind == ParagraphEditKind.Insert ? 0 : 1;
                        if (edit.ParagraphIndex < lowest || edit.ParagraphIndex > paraNum)
                        {
                            failures.Add($"{edit.SlideIndex}쪽 [{edit.ShapeName}] " +
                                         $"문단 {edit.ParagraphIndex} — 문단 수({paraNum})와 어긋남");
                            continue;
                        }
                        try
                        {
                            switch (edit.Kind)
                            {
                                case ParagraphEditKind.Delete:
                                    tr.Paragraphs(edit.ParagraphIndex, 1).Delete();
                                    break;
                                case ParagraphEditKind.Insert:
                                    InsertParagraphAfter(tr, edit, markRed, snapshots);
                                    break;
                                default:
                                    ReplaceParagraph(tr, edit, markRed, snapshots);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            failures.Add($"{edit.SlideIndex}쪽 [{edit.ShapeName}] " +
                                         $"문단 {edit.ParagraphIndex} — {ex.Message}");
                        }
                    }
                }
                SaveUnlessKept(prs!);
            }
            finally { if (ownedByUs) SafeClose(prs); }

            return new ParagraphEditResult(snapshots, failures);
        });

    /// <summary>기존 문단의 텍스트만 교체. Bullet → Indent → Font 순서 복원은 필수다.</summary>
    private static void ReplaceParagraph(
        dynamic tr, ParagraphEdit edit, bool markRed, List<ParagraphColorSnapshot> snapshots)
    {
        dynamic para = tr.Paragraphs(edit.ParagraphIndex, 1);

        // 문단 범위에는 끝의 \r이 포함되므로 비교 시 떼어낸다.
        string current = ((string)para.Text).TrimEnd('\r');
        if (current == edit.NewText && !markRed) return;

        var snap = ReadColor(para, edit.SlideIndex, edit.ShapeName, edit.ParagraphIndex);

        var bulletVisible = para.ParagraphFormat.Bullet.Visible;
        var bulletChar    = para.ParagraphFormat.Bullet.Character;
        var indentLevel   = para.ParagraphFormat.IndentLevel;
        var fontName      = para.Font.Name;
        var fontSize      = para.Font.Size;

        para.Text = edit.NewText;

        // 복원 순서를 바꾸면 불릿이 되살아나지 않는다.
        para.ParagraphFormat.Bullet.Visible = bulletVisible;
        if ((int)bulletVisible == MsoTrue)
        {
            try { para.ParagraphFormat.Bullet.Character = bulletChar; } catch { }
        }
        para.ParagraphFormat.IndentLevel = indentLevel;
        try { para.Font.Name = fontName; } catch { }
        try { para.Font.Size = fontSize; } catch { }

        if (markRed)
        {
            snapshots.Add(snap);
            // PowerPoint COM의 RGB는 BGR 순서다. 빨강은 0xFF0000이 아니라 255.
            try { tr.Paragraphs(edit.ParagraphIndex, 1).Font.Fill.ForeColor.RGB = 255; } catch { }
        }
    }

    /// <summary>
    /// 기준 문단 뒤에 새 문단을 삽입한다. InsertAfter/InsertBefore가 이웃 텍스트의 서식을
    /// 물려받으므로 정렬·글꼴·크기가 자동으로 보존된다.
    /// ParagraphIndex가 0이면 맨 앞에 넣는다 (지웠던 첫 문단을 되살릴 때).
    /// </summary>
    private static void InsertParagraphAfter(
        dynamic tr, ParagraphEdit edit, bool markRed, List<ParagraphColorSnapshot> snapshots)
    {
        dynamic inserted;
        dynamic neighbor;
        int     newIndex;

        if (edit.ParagraphIndex == 0)
        {
            neighbor = tr.Paragraphs(1, 1);
            inserted = neighbor.InsertBefore(edit.NewText + "\r");
            newIndex = 1;
        }
        else
        {
            neighbor = tr.Paragraphs(edit.ParagraphIndex, 1);
            // 마지막 문단은 끝에 \r이 없으므로 앞에 붙여야 한 문단이 더 생긴다.
            inserted = ((string)neighbor.Text).EndsWith('\r')
                ? neighbor.InsertAfter(edit.NewText + "\r")
                : neighbor.InsertAfter("\r" + edit.NewText);
            newIndex = edit.ParagraphIndex + 1;
        }

        if (!markRed) return;

        var neighborColor = ReadColor(neighbor, edit.SlideIndex, edit.ShapeName, newIndex);
        snapshots.Add(neighborColor);
        try { inserted.Font.Fill.ForeColor.RGB = 255; } catch { }
    }

    public void RestoreParagraphColors(
        string presentationPath, IReadOnlyList<ParagraphColorSnapshot> snapshots)
        => RunOnSta(() =>
        {
            if (snapshots.Count == 0) return;

            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs   = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs)
                prs = pptApp.Presentations.Open(presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                foreach (var group in snapshots.GroupBy(s => (s.SlideIndex, s.ShapeName)))
                {
                    dynamic slide = prs!.Slides[group.Key.SlideIndex];
                    dynamic? tf   = ResolveTextFrame(slide, group.Key.ShapeName);
                    if (tf is null) continue;

                    dynamic tr      = tf.TextRange;
                    int     paraNum = (int)tr.Paragraphs.Count;
                    foreach (var snap in group)
                    {
                        if (snap.ParagraphIndex < 1 || snap.ParagraphIndex > paraNum) continue;
                        try
                        {
                            dynamic fore = tr.Paragraphs(snap.ParagraphIndex, 1).Font.Fill.ForeColor;
                            // 테마색이 걸려 있던 문단은 RGB만 되돌리면 테마 연결이 끊긴다.
                            if (snap.ObjectThemeColor > 0) fore.ObjectThemeColor = snap.ObjectThemeColor;
                            else                           fore.RGB             = snap.Rgb;
                        }
                        catch { }
                    }
                }
                SaveUnlessKept(prs!);
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    public int DuplicateSlide(string presentationPath, int sourceSlideIndex, int insertAfterSlideIndex)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs) prs = pptApp.Presentations.Open(presentationPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                int slideCount = (int)prs!.Slides.Count;
                if (sourceSlideIndex < 1 || sourceSlideIndex > slideCount)
                    throw new ArgumentOutOfRangeException(nameof(sourceSlideIndex));
                if (insertAfterSlideIndex < 0 || insertAfterSlideIndex > slideCount)
                    throw new ArgumentOutOfRangeException(nameof(insertAfterSlideIndex));
                dynamic duplicated = prs.Slides[sourceSlideIndex].Duplicate();
                int targetIndex = insertAfterSlideIndex + 1;
                duplicated.MoveTo(targetIndex);
                SaveUnlessKept(prs);
                return targetIndex;
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    public int ImportSlide(string sourcePath, int sourceSlideIndex, string targetPath, int insertAfterSlideIndex)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? source = FindOpenPresentation(pptApp, sourcePath);
            dynamic? target = FindOpenPresentation(pptApp, targetPath);
            bool ownsSource = source is null;
            bool ownsTarget = target is null;
            if (ownsSource) source = pptApp.Presentations.Open(sourcePath, MsoTrue, MsoFalse, MsoFalse);
            if (ownsTarget) target = pptApp.Presentations.Open(targetPath, MsoFalse, MsoFalse, MsoFalse);
            try
            {
                if (sourceSlideIndex < 1 || sourceSlideIndex > (int)source!.Slides.Count)
                    throw new ArgumentOutOfRangeException(nameof(sourceSlideIndex));
                if (insertAfterSlideIndex < 0 || insertAfterSlideIndex > (int)target!.Slides.Count)
                    throw new ArgumentOutOfRangeException(nameof(insertAfterSlideIndex));
                source.Slides[sourceSlideIndex].Copy();
                dynamic pasted = target.Slides.Paste();
                var targetIndex = insertAfterSlideIndex + 1;
                pasted.MoveTo(targetIndex);
                SaveUnlessKept(target);
                return targetIndex;
            }
            finally
            {
                if (ownsTarget) SafeClose(target);
                if (ownsSource) SafeClose(source);
            }
        });

    // ── 도형 목록·배치·서식 ──────────────────────────────────────
    // 주소는 모두 ShapeAddress 규약("이름", "이름#2", "그룹/이름")을 따른다.

    /// <summary>슬라이드 하나의 작업을 열린 문서(없으면 직접 연 문서)에서 수행한다.</summary>
    private static T OnSlide<T>(string presentationPath, int slideIndex, bool save, Func<dynamic, T> work)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? prs   = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = prs is null;
            if (ownedByUs)
                prs = pptApp.Presentations.Open(presentationPath, save ? MsoFalse : MsoTrue, MsoFalse, MsoFalse);
            try
            {
                T result = work(prs!.Slides[slideIndex]);
                if (save) SaveUnlessKept(prs);
                return result;
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    /// <summary>그룹 안까지 펼친 도형 목록. 이름이 겹치면 주소에 "#n"을 붙여 구분한다.</summary>
    public IReadOnlyList<ShapeEntry> ListShapes(string presentationPath, int slideIndex)
        => OnSlide(presentationPath, slideIndex, save: false, slide =>
        {
            List<(dynamic Shape, IReadOnlyList<string> Path)> shapes = FlattenShapesWithPath(slide.Shapes);
            List<string> names = shapes.Select(s => (string)s.Shape.Name).ToList();
            var seen   = new Dictionary<string, int>();
            var result = new List<ShapeEntry>(shapes.Count);
            for (int i = 0; i < shapes.Count; i++)
            {
                dynamic shape = shapes[i].Shape;
                string  name  = names[i];
                seen[name] = seen.GetValueOrDefault(name) + 1;

                bool   hasText = false;
                string preview = "";
                try
                {
                    if ((int)shape.HasTextFrame == MsoTrue)
                    {
                        string text = (string)shape.TextFrame2.TextRange.Text;
                        hasText = text.Length > 0;
                        // 문단 경계를 남겨야 미리보기를 원문으로 착각하지 않는다(¶ 문단, ↵ Shift+Enter).
                        preview = text[..Math.Min(60, text.Length)].Replace("\r", " ¶ ").Replace("\v", " ↵ ");
                    }
                }
                catch { }

                int type = 0, id = 0;
                try { type = (int)shape.Type; } catch { }
                try { id = (int)shape.Id; } catch { }
                // 겹치는 이름은 "#n" 대신 Id 주소를 준다. "#n"은 다른 도형 이름을 바꾸면 밀려
                // 모델이 엉뚱한 상자를 고쳤다(2026-10-06 실측).
                int sameName = names.Count(n => n == name);
                string address = id > 0 ? ShapeAddress.FormatStable(name, sameName, id) : ShapeAddress.Format(name, seen[name], sameName);
                result.Add(new ShapeEntry(
                    address, name, string.Join("/", shapes[i].Path), type,
                    (float)shape.Left, (float)shape.Top, (float)shape.Width, (float)shape.Height,
                    hasText, preview, id));
            }
            return (IReadOnlyList<ShapeEntry>)result;
        });

    public void RenameShape(string presentationPath, int slideIndex, string address, string newName)
        => OnSlide(presentationPath, slideIndex, save: true, slide =>
        {
            RequireShape(slide, slideIndex, address).Name = newName;
            return true;
        });

    public void SetShapeGeometry(string presentationPath, int slideIndex, string address, ShapeGeometry geometry)
        => OnSlide(presentationPath, slideIndex, save: true, slide =>
        {
            dynamic shape = RequireShape(slide, slideIndex, address);
            // 좌우 반전된 도형은 너비를 바꾸면 Left가 밀리므로 크기를 먼저 바꾸고 위치를 다시 맞춘다.
            float left = (float)shape.Left, top = (float)shape.Top;
            if (geometry.Width  is float w) shape.Width  = w;
            if (geometry.Height is float h) shape.Height = h;
            shape.Left = geometry.Left ?? left;
            shape.Top  = geometry.Top  ?? top;
            return true;
        });

    public void DeleteShape(string presentationPath, int slideIndex, string address)
        => OnSlide(presentationPath, slideIndex, save: true, slide =>
        {
            RequireShape(slide, slideIndex, address).Delete();
            return true;
        });

    /// <summary>텍스트를 담은 사각형(둥근 모서리 가능)을 추가하고 최종 이름을 돌려준다.</summary>
    public string AddTextShape(string presentationPath, int slideIndex, TextShapeRequest request)
        => OnSlide(presentationPath, slideIndex, save: true, slide =>
        {
            const int msoShapeRectangle = 1, msoShapeRoundedRectangle = 5;
            int kind = request.Kind == TextShapeKind.RoundedRectangle ? msoShapeRoundedRectangle : msoShapeRectangle;
            dynamic shape = slide.Shapes.AddShape(kind, request.Left, request.Top, request.Width, request.Height);
            if (!string.IsNullOrEmpty(request.Name)) shape.Name = request.Name;
            if (request.Kind == TextShapeKind.RoundedRectangle)
                try { shape.Adjustments[1] = 0.5f; } catch { }

            if (request.FillColor is null) shape.Fill.Visible = MsoFalse;
            else shape.Fill.ForeColor.RGB = OleColor.FromHex(request.FillColor);
            shape.Line.Visible = MsoFalse;
            try { shape.Shadow.Visible = MsoFalse; } catch { }

            dynamic tf = shape.TextFrame2;
            tf.MarginLeft = request.Padding; tf.MarginRight = request.Padding;
            tf.MarginTop  = 0f;              tf.MarginBottom = 0f;
            tf.VerticalAnchor = 3; // msoAnchorMiddle
            tf.WordWrap = request.AutoWidth ? MsoFalse : MsoTrue;

            dynamic tr = tf.TextRange;
            tr.Text = request.Text;
            tr.ParagraphFormat.Alignment = 1; // msoAlignLeft
            if (!string.IsNullOrEmpty(request.FontName))
            {
                tr.Font.Name = request.FontName;
                tr.Font.NameFarEast = request.FontName;
            }
            tr.Font.Size = request.FontSize;
            tr.Font.Bold = request.Bold ? MsoTrue : MsoFalse;
            tr.Font.Fill.ForeColor.RGB = OleColor.FromHex(request.FontColor);

            if (request.AutoWidth)
            {
                // 글자 폭에 맞춰 늘린 뒤 높이는 요청값으로 되돌린다(자동 맞춤은 높이도 줄인다).
                tf.AutoSize = 1; // msoAutoSizeShapeToFitText
                tf.AutoSize = 0;
                shape.Height = request.Height;
            }
            shape.Top  = request.Top;
            shape.Left = request.RightEdge is float right ? right - (float)shape.Width : request.Left;
            return (string)shape.Name;
        });

    /// <summary>도형 텍스트에서 <paramref name="find"/>가 나오는 곳마다 글자색을 바꾼다. 바꾼 횟수를 돌려준다.</summary>
    public int SetTextColor(string presentationPath, int slideIndex, string address, string find, string color)
        => OnSlide(presentationPath, slideIndex, save: true, slide =>
        {
            if (string.IsNullOrEmpty(find)) return 0;
            dynamic tr   = RequireShape(slide, slideIndex, address).TextFrame2.TextRange;
            string  text = (string)tr.Text;
            int     ole  = OleColor.FromHex(color);
            int     count = 0;
            for (int i = text.IndexOf(find, StringComparison.Ordinal); i >= 0;
                 i = text.IndexOf(find, i + find.Length, StringComparison.Ordinal))
            {
                tr.Characters(i + 1, find.Length).Font.Fill.ForeColor.RGB = ole;
                count++;
            }
            return count;
        });

    public void SetTextFormat(string presentationPath, int slideIndex, string address, TextFormatPatch patch)
        => OnSlide(presentationPath, slideIndex, save: true, slide =>
        {
            dynamic font = RequireShape(slide, slideIndex, address).TextFrame2.TextRange.Font;
            if (patch.Size    is float size)    font.Size    = size;
            if (patch.Bold    is bool  bold)    font.Bold    = bold ? MsoTrue : MsoFalse;
            if (patch.Spacing is float spacing) font.Spacing = spacing;
            if (patch.Color   is string color)  font.Fill.ForeColor.RGB = OleColor.FromHex(color);
            if (!string.IsNullOrEmpty(patch.FontName))
            {
                font.Name = patch.FontName;
                font.NameFarEast = patch.FontName;
            }
            return true;
        });

    /// <summary>
    /// 한 도형 안에서만 치환한다. TextRange2.Replace는 바뀐 자리 첫 글자의 서식을 이어받으므로
    /// 글자 크기가 섞인 제목도 서식을 지킨다. 런 경계에 걸친 문자열도 찾는다.
    /// </summary>
    public int ReplaceTextInShape(string presentationPath, int slideIndex, string address, string find, string replace)
        => OnSlide(presentationPath, slideIndex, save: true, slide =>
        {
            if (string.IsNullOrEmpty(find)) return 0;
            dynamic tr = RequireShape(slide, slideIndex, address).TextFrame2.TextRange;
            int count = 0, after = 0;
            while (count < 1000)
            {
                dynamic? hit = tr.Replace(find, replace, after, MsoTrue, MsoFalse);
                if (hit is null) break;
                count++;
                // 바꾼 글자 뒤부터 다시 찾는다. replace에 find가 들어 있어도 끝나도록.
                after = (int)hit.Start - 1 + (int)hit.Length;
            }
            return count;
        });

    /// <summary>첫 런의 서식. 한 상자 안에서 섞인 서식은 LayoutLint가 첫 런 기준으로 비교한다.</summary>
    public TextFormatInfo GetTextFormat(string presentationPath, int slideIndex, string address)
        => OnSlide(presentationPath, slideIndex, save: false, slide =>
        {
            dynamic range = RequireShape(slide, slideIndex, address).TextFrame2.TextRange;
            dynamic font  = (int)range.Runs.Count > 0 ? range.Runs(1, 1).Font : range.Font;
            string name = "";
            try { name = (string)font.NameFarEast; } catch { }
            if (string.IsNullOrEmpty(name)) try { name = (string)font.Name; } catch { }
            float size = 0f, spacing = 0f;
            bool  bold = false;
            string color = "";
            try { size    = (float)Math.Round((float)font.Size, 1); }    catch { }
            try { spacing = (float)Math.Round((float)font.Spacing, 1); } catch { }
            try { bold    = (int)font.Bold == MsoTrue; }                  catch { }
            try { color   = OleColor.ToHex((int)font.Fill.ForeColor.RGB); } catch { }
            return new TextFormatInfo(name, size, spacing, bold, color);
        });

    // ── 단일 Shape 정보 읽기 ─────────────────────────────────────

    public PptSlideInfo? GetShapeInfo(string presentationPath, int slideIndex, string shapeName)
        => RunOnSta(() =>
        {
            var rules  = PptRules.Current;
            dynamic pptApp = GetOrCreateInstance();
            dynamic? existing = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = existing is null;
            dynamic prs = existing
                ?? pptApp.Presentations.Open(presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                dynamic slide = prs.Slides[slideIndex];

                // 테이블 셀 이름 파싱: "Table 1 [R2C3]"
                var cellMatch = Regex.Match(shapeName, @"^(.+)\s+\[R(\d+)C(\d+)\]$");
                if (cellMatch.Success)
                {
                    string tableName = cellMatch.Groups[1].Value;
                    int    row       = int.Parse(cellMatch.Groups[2].Value);
                    int    col       = int.Parse(cellMatch.Groups[3].Value);
                    foreach (dynamic shape in MatchingShape(slide, shapeName))
                    {
                        if ((int)shape.HasTable != MsoTrue) break;
                        try
                        {
                            dynamic cell  = shape.Table.Cell(row, col);
                            dynamic ctf   = cell.Shape.TextFrame2;
                            string  text  = (string)ctf.TextRange.Text;
                            if (string.IsNullOrWhiteSpace(text))
                                return new PptSlideInfo(slideIndex, shapeName, "", "(빈 셀)",
                                    0f, 0f, 0, false) { IsEmpty = true };

                            float cbH = 0f, ccH = 0f; int cln = 0; float cfs = 0f;
                            try { cbH = (float)Math.Round((float)ctf.TextRange.BoundHeight, 1); } catch { }
                            try { ccH = (float)Math.Round((float)cell.Shape.Height, 1); }         catch { }
                            try { cln = (int)ctf.TextRange.Lines.Count; }                         catch { }
                            try { cfs = (float)Math.Round((float)ctf.TextRange.Font.Size, 1); }   catch { }
                            bool cov = cbH > ccH + rules.OverflowTolerance;
                            string cprev = text[..Math.Min(60, text.Length)]
                                .Replace('\r', ' ').Replace('\n', ' ');
                            return new PptSlideInfo(slideIndex, shapeName, text, cprev,
                                cbH, ccH, cln, cov, cfs)
                            {
                                MaxLineHw = CalcMaxLineHw(text),
                                HasBullet = ReadHasBullet(ctf),
                            };
                        }
                        catch { break; }
                    }
                    return null;
                }

                // 일반 텍스트 shape
                foreach (dynamic shape in MatchingShape(slide, shapeName))
                {
                    if ((int)shape.HasTextFrame != MsoTrue) return null;
                    try
                    {
                        var    tf   = shape.TextFrame2;
                        string text = (string)tf.TextRange.Text;
                        if (string.IsNullOrWhiteSpace(text))
                            return new PptSlideInfo(slideIndex, shapeName, "", "(빈 상자)",
                                0f, (float)Math.Round((float)shape.Height, 1), 0, false) { IsEmpty = true };

                        float boundH    = (float)Math.Round((float)tf.TextRange.BoundHeight, 1);
                        float boxH      = (float)Math.Round((float)shape.Height, 1);
                        int   lineCount = 0;
                        float fontSize  = 0f;
                        try { lineCount = (int)tf.TextRange.Lines.Count; }                      catch { }
                        try { fontSize  = (float)Math.Round((float)tf.TextRange.Font.Size, 1); } catch { }
                        bool overflow   = boundH > boxH + rules.OverflowTolerance;
                        string preview  = text[..Math.Min(60, text.Length)]
                            .Replace('\r', ' ').Replace('\n', ' ');
                        return new PptSlideInfo(slideIndex, shapeName, text, preview,
                            boundH, boxH, lineCount, overflow, fontSize)
                        {
                            MaxLineHw = CalcMaxLineHw(text),
                            HasBullet = ReadHasBullet(tf),
                        };
                    }
                    catch { return null; }
                }
                return null;
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    // ── 슬라이드 레이아웃 + 썸네일 ───────────────────────────────

    public SlideLayout GetSlideLayout(string presentationPath, int slideIndex)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? existing = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = existing is null;
            dynamic prs = existing
                ?? pptApp.Presentations.Open(presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                float slideW = (float)prs.PageSetup.SlideWidth;
                float slideH = (float)prs.PageSetup.SlideHeight;
                var   shapes = new List<ShapeLayoutInfo>();
                dynamic slide = prs.Slides[slideIndex];

                foreach (dynamic shape in FlattenShapes(slide.Shapes))
                {
                    float  left   = (float)shape.Left;
                    float  top    = (float)shape.Top;
                    float  width  = (float)shape.Width;
                    float  height = (float)shape.Height;
                    string name   = (string)shape.Name;
                    int    sType  = (int)shape.Type;

                    ShapeCategory cat;
                    IReadOnlyList<TextLineLayout>? textLines = null;
                    float boundH = 0f;

                    if (sType == 13 || sType == 11)
                    {
                        cat = ShapeCategory.Image;
                    }
                    else if ((int)shape.HasTextFrame == MsoTrue)
                    {
                        string text = "";
                        try { text = (string)shape.TextFrame2.TextRange.Text; } catch { }
                        cat = string.IsNullOrWhiteSpace(text) ? ShapeCategory.Other : ShapeCategory.Text;

                        // 텍스트 shape: 줄별 폭 비율 + 정렬 수집
                        if (cat == ShapeCategory.Text)
                        {
                            try
                            {
                                var tf = shape.TextFrame2;
                                float marginL = 0f, marginR = 0f;
                                try { marginL = (float)tf.MarginLeft; } catch { }
                                try { marginR = (float)tf.MarginRight; } catch { }
                                float innerW = width - marginL - marginR;
                                if (innerW < 1f) innerW = width;

                                // 문단별 불릿 유무 수집
                                int paraCount = (int)tf.TextRange.Paragraphs.Count;
                                var paraBullets = new List<bool>(paraCount);
                                for (int p = 1; p <= paraCount; p++)
                                {
                                    bool hasBullet = false;
                                    try
                                    {
                                        int bt = (int)tf.TextRange.Paragraphs(p, 1).ParagraphFormat.Bullet.Type;
                                        hasBullet = bt >= 1; // 1=Unnumbered, 2=Numbered, 3=Picture
                                    }
                                    catch { }
                                    paraBullets.Add(hasBullet);
                                }

                                int lineCount = (int)tf.TextRange.Lines.Count;
                                var lines = new List<TextLineLayout>(lineCount);
                                int paraIdx = 0;
                                bool isFirstLineOfPara = true;
                                for (int i = 1; i <= lineCount; i++)
                                {
                                    dynamic lineRange = tf.TextRange.Lines(i, 1);
                                    float boundW = 0f;
                                    int align = 1; // ppAlignLeft
                                    float fontSize = 12f;
                                    try { boundW = (float)lineRange.BoundWidth; } catch { }
                                    try { align = (int)lineRange.ParagraphFormat.Alignment; } catch { }
                                    try { fontSize = (float)lineRange.Font.Size; } catch { }
                                    float ratio = innerW > 0 ? Math.Clamp(boundW / innerW, 0f, 1f) : 1f;

                                    // 줄넘김 유형: 0=None, 1=AutoWrap, 2=Enter, 3=ShiftEnter
                                    string lt = (string)lineRange.Text;
                                    int breakType = 0; // None (last line)
                                    if (i < lineCount && lt.Length > 0)
                                    {
                                        char last = lt[^1];
                                        breakType = last == '\r' ? 2 : last == '\v' ? 3 : 1;
                                    }

                                    bool bullet = isFirstLineOfPara && paraIdx < paraBullets.Count && paraBullets[paraIdx];
                                    lines.Add(new TextLineLayout(ratio, align, bullet, fontSize, breakType));

                                    // Enter(\r)로 끝나면 다음 문단
                                    if (breakType == 2)
                                    {
                                        paraIdx++;
                                        isFirstLineOfPara = true;
                                    }
                                    else
                                    {
                                        isFirstLineOfPara = false;
                                    }
                                }
                                textLines = lines;
                                try { boundH = (float)tf.TextRange.BoundHeight; } catch { }
                            }
                            catch { }
                        }
                    }
                    else
                    {
                        cat = ShapeCategory.Other;
                    }

                    if (cat != ShapeCategory.Image && sType == 14)
                    {
                        try { if ((int)shape.Fill.Type == 6) cat = ShapeCategory.Image; } catch { }
                    }

                    float rotation = 0f;
                    try { rotation = (float)shape.Rotation; } catch { }

                    shapes.Add(new ShapeLayoutInfo(left, top, width, height, cat, name, textLines, boundH, rotation));
                }

                return new SlideLayout(slideW, slideH, shapes);
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    public Task<string?> GetSlideThumbnailAsync(string presentationPath, int slideIndex)
        => Task.Run(() => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? existing = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = existing is null;
            dynamic prs = existing ?? pptApp.Presentations.Open(
                presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                string tempFile = Path.Combine(
                    Path.GetTempPath(),
                    $"bckppt_thumb_{Path.GetFileNameWithoutExtension(presentationPath)}_{slideIndex}_{Guid.NewGuid():N}.png");
                // Export(path, filterName, scaleWidth, scaleHeight)
                prs.Slides[slideIndex].Export(tempFile, "PNG", 960, 540);
                return (string?)tempFile;
            }
            catch { return (string?)null; }
            finally { if (ownedByUs) SafeClose(prs); }
        }));

    // ── 최근 파일 목록 ────────────────────────────────────────────

    public IReadOnlyList<RecentPptxFile> GetRecentFiles()
    {
        var result = new List<RecentPptxFile>();
        try
        {
            // PowerPoint MRU: HKCU\Software\Microsoft\Office\{ver}\PowerPoint\File MRU
            using var officeKey = Registry.CurrentUser
                .OpenSubKey(@"Software\Microsoft\Office");
            if (officeKey is null) return result;

            foreach (var verName in officeKey.GetSubKeyNames())
            {
                if (!float.TryParse(verName, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out _)) continue;
                using var mruKey = officeKey
                    .OpenSubKey($@"{verName}\PowerPoint\File MRU");
                if (mruKey is null) continue;

                foreach (var valueName in mruKey.GetValueNames())
                {
                    if (mruKey.GetValue(valueName) is not string raw) continue;
                    // 형식: "[F00000000][T…][O…]*path" 또는 단순 path
                    var path = ExtractPath(raw);
                    if (string.IsNullOrEmpty(path)) continue;
                    if (!File.Exists(path)) continue;
                    var info = new FileInfo(path);
                    result.Add(new RecentPptxFile(info.Name, info.FullName,
                        info.LastWriteTime, info.Length));
                }
                break;
            }
        }
        catch { }

        return result
            .GroupBy(f => f.FullPath, StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderByDescending(f => f.ModifiedAt)
            .Take(20)
            .ToList();
    }

    private static string ExtractPath(string raw)
    {
        int star = raw.LastIndexOf('*');
        return star >= 0 ? raw[(star + 1)..] : raw;
    }

    // ── 텍스트 줄넘김 분석 ──────────────────────────────────────────

    public IReadOnlyList<TextLineInfo> GetTextLineBreaks(string presentationPath, int slideIndex, string shapeName)
        => RunOnSta(() =>
        {
            dynamic pptApp = GetOrCreateInstance();
            dynamic? existing = FindOpenPresentation(pptApp, presentationPath);
            bool ownedByUs = existing is null;
            dynamic prs = existing
                ?? pptApp.Presentations.Open(presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                dynamic slide = prs.Slides[slideIndex];
                dynamic? tf = null;

                // 테이블 셀 파싱
                var cellMatch = Regex.Match(shapeName, @"^(.+)\s+\[R(\d+)C(\d+)\]$");
                if (cellMatch.Success)
                {
                    string tableName = cellMatch.Groups[1].Value;
                    int row = int.Parse(cellMatch.Groups[2].Value);
                    int col = int.Parse(cellMatch.Groups[3].Value);
                    foreach (dynamic shape in MatchingShape(slide, shapeName))
                    {
                        if ((int)shape.HasTable != MsoTrue) break;
                        tf = shape.Table.Cell(row, col).Shape.TextFrame2;
                        break;
                    }
                }
                else
                {
                    foreach (dynamic shape in MatchingShape(slide, shapeName))
                    {
                        if ((int)shape.HasTextFrame != MsoTrue) break;
                        tf = shape.TextFrame2;
                        break;
                    }
                }

                if (tf is null) return Array.Empty<TextLineInfo>();

                int lineCount = (int)tf.TextRange.Lines.Count;
                if (lineCount == 0) return Array.Empty<TextLineInfo>();

                // ── 문단별 불릿/들여쓰기/fontSize 수집 ──
                var paraBullets = new List<(string bullet, int indent, float fontSize)>();
                int paraCount = (int)tf.TextRange.Paragraphs.Count;
                for (int p = 1; p <= paraCount; p++)
                {
                    string bullet = "";
                    int indent = 1;
                    float fontSize = 12f;
                    try
                    {
                        dynamic para = tf.TextRange.Paragraphs(p, 1);
                        try { indent = (int)para.ParagraphFormat.IndentLevel; } catch { }
                        try
                        {
                            float fs = (float)para.Font.Size;
                            if (fs > 0) fontSize = fs;
                        }
                        catch { }
                        int bulletType = (int)para.ParagraphFormat.Bullet.Type;
                        // ppBulletUnnumbered = 1
                        if (bulletType == 1)
                        {
                            try
                            {
                                int charCode = (int)para.ParagraphFormat.Bullet.Character;
                                bullet = charCode > 0 ? ((char)charCode).ToString() : "\u2022";
                            }
                            catch { bullet = "\u2022"; }
                        }
                        // ppBulletNumbered = 2
                        else if (bulletType == 2)
                        {
                            try
                            {
                                int num = (int)para.ParagraphFormat.Bullet.Number;
                                bullet = $"{num}.";
                            }
                            catch { bullet = "#."; }
                        }
                        // ppBulletPicture = 3
                        else if (bulletType == 3)
                        {
                            bullet = "\u25A0"; // ■ picture bullet placeholder
                        }
                    }
                    catch { }
                    paraBullets.Add((bullet, indent, fontSize));
                }

                // ── 시각적 줄 순회 + 문단 매핑 ──
                int paraIdx = 0;
                bool isFirstLineOfPara = true;
                var result = new List<TextLineInfo>(lineCount);

                for (int i = 1; i <= lineCount; i++)
                {
                    dynamic lineRange = tf.TextRange.Lines(i, 1);
                    string lineText = (string)lineRange.Text;

                    LineBreakType breakType = LineBreakType.None;
                    string displayText = lineText;

                    if (i < lineCount && lineText.Length > 0)
                    {
                        char last = lineText[^1];
                        if (last == '\r')
                        {
                            breakType = LineBreakType.Enter;
                            displayText = lineText[..^1];
                        }
                        else if (last == '\v')
                        {
                            breakType = LineBreakType.ShiftEnter;
                            displayText = lineText[..^1];
                        }
                        else
                        {
                            breakType = LineBreakType.AutoWrap;
                        }
                    }
                    else
                    {
                        // 마지막 줄: trailing \r 제거
                        displayText = lineText.TrimEnd('\r');
                    }

                    // 문단 첫 줄이면 불릿 정보 부여
                    string bulletText = "";
                    int indentLevel = 1;
                    float lineFs = 12f;
                    if (paraIdx < paraBullets.Count)
                    {
                        indentLevel = paraBullets[paraIdx].indent;
                        lineFs      = paraBullets[paraIdx].fontSize;
                        if (isFirstLineOfPara)
                            bulletText = paraBullets[paraIdx].bullet;
                    }

                    result.Add(new TextLineInfo(i, displayText, breakType, bulletText, indentLevel, lineFs));

                    // Enter → 다음 문단으로 이동
                    if (breakType == LineBreakType.Enter)
                    {
                        paraIdx++;
                        isFirstLineOfPara = true;
                    }
                    else
                    {
                        isFirstLineOfPara = false;
                    }
                }
                return (IReadOnlyList<TextLineInfo>)result;
            }
            finally { if (ownedByUs) SafeClose(prs); }
        });

    // ── AutoFit 내부 로직 ─────────────────────────────────────────

    private static bool IsOverflowing(dynamic shape, PptRules rules)
        => (float)shape.TextFrame2.TextRange.BoundHeight
           > (float)shape.Height + rules.OverflowTolerance;

    private static float GetFontSize(dynamic shape)
        => (float)shape.TextFrame2.TextRange.Font.Size;

    /// <returns>true if text was forcibly trimmed</returns>
    private static bool ApplyAutoFit(dynamic shape, PptRules rules)
    {
        while (IsOverflowing(shape, rules) && GetFontSize(shape) > rules.MinFontSize)
            shape.TextFrame2.TextRange.Font.Size = GetFontSize(shape) - 1f;

        if (IsOverflowing(shape, rules)) AdjustSpacing(shape, rules);
        if (IsOverflowing(shape, rules)) return ForceTrim(shape, rules);
        return false;
    }

    private static void AdjustSpacing(dynamic shape, PptRules rules)
    {
        foreach (dynamic para in shape.TextFrame2.TextRange.Paragraphs)
        {
            try
            {
                float spacing = (float)para.Font.Spacing;
                para.Font.Spacing = Math.Max(spacing - rules.SpacingStep, rules.SpacingMinimum);
                float spaceAfter = (float)para.ParagraphFormat.SpaceAfter;
                para.ParagraphFormat.SpaceAfter = Math.Max(spaceAfter - rules.SpaceAfterStep, rules.SpaceAfterMinimum);
            }
            catch { }
        }
    }

    private static bool ForceTrim(dynamic shape, PptRules rules)
    {
        dynamic paras = shape.TextFrame2.TextRange.Paragraphs;
        while (IsOverflowing(shape, rules) && (int)paras.Count > 1)
            paras[(int)paras.Count].Delete();

        if (IsOverflowing(shape, rules) && (int)paras.Count >= 1)
        {
            dynamic lastPara = paras[(int)paras.Count];
            string text = ((string)lastPara.Text).TrimEnd('\r', '\n', ' ');
            while (IsOverflowing(shape, rules) && text.Length > 5)
            {
                text = text[..^2] + "\u2026";
                if ((int)lastPara.Runs.Count > 0)
                    lastPara.Runs[1].Text = text;
            }
            return true;
        }
        return false;
    }

    // ── Navigate 헬퍼 ─────────────────────────────────────────────

    private static bool TryNavigate(string path, int slideIndex, string shapeName)
    {
        dynamic? pptApp = TryGetActiveInstance();
        if (pptApp is null) return false;
        foreach (dynamic win in pptApp.Windows)
        {
            string winPath;
            try { winPath = (string)win.Presentation.FullName; } catch { continue; }
            if (!string.Equals(winPath, path, StringComparison.OrdinalIgnoreCase)) continue;
            pptApp.Activate();
            win.Activate();
            win.View.GotoSlide(slideIndex);

            // 표 셀 여부 확인: "TableName [R2C3]" 형식
            var cellMatch = Regex.Match(shapeName, @"^(.+)\s+\[R(\d+)C(\d+)\]$");
            if (cellMatch.Success)
            {
                string tableName = cellMatch.Groups[1].Value;
                int    row       = int.Parse(cellMatch.Groups[2].Value);
                int    col       = int.Parse(cellMatch.Groups[3].Value);
                foreach (dynamic shape in win.Presentation.Slides[slideIndex].Shapes)
                {
                    if ((string)shape.Name != tableName) continue;
                    if ((int)shape.HasTable != MsoTrue) break;
                    try
                    {
                        shape.Select(MsoTrue);
                        shape.Table.Cell(row, col).Shape.TextFrame.TextRange.Select();
                    }
                    catch { shape.Select(MsoTrue); }
                    break;
                }
            }
            else
            {
                foreach (dynamic shape in win.Presentation.Slides[slideIndex].Shapes)
                {
                    if ((string)shape.Name == shapeName) { shape.Select(MsoTrue); break; }
                }
            }
            return true;
        }
        return false;
    }

    private static dynamic? FindOpenPresentation(dynamic pptApp, string path)
    {
        foreach (dynamic prs in pptApp.Presentations)
        {
            try
            {
                if (string.Equals((string)prs.FullName, path, StringComparison.OrdinalIgnoreCase))
                    return prs;
            }
            catch { }
        }
        return null;
    }

    // ── COM 연결 ──────────────────────────────────────────────────

    private static dynamic? TryGetActiveInstance()
    {
        try { return (dynamic)OleGetActiveObject("PowerPoint.Application"); }
        catch (COMException) { return null; }
    }

    private static dynamic GetOrCreateInstance()
    {
        try { return (dynamic)OleGetActiveObject("PowerPoint.Application"); }
        catch (COMException)
        {
            var pptType = Type.GetTypeFromProgID("PowerPoint.Application")
                ?? throw new InvalidOperationException("PowerPoint가 설치되어 있지 않습니다.");
            dynamic app = Activator.CreateInstance(pptType)!;
            app.Visible = MsoTrue;
            return app;
        }
    }
}
