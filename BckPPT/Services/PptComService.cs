using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using CellBox2.BckPPT.Models;
using Microsoft.Win32;

namespace CellBox2.BckPPT.Services;

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
                prs.Save();
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
                prs.Save();
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
        => Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });

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
                prs.Save();
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
                if (changed.Count > 0) prs.Save();
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
                prs.Save();
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

                prs.Save();
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
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((string)shape.Name != tableName) continue;
                        if ((int)shape.HasTable != MsoTrue) break;
                        shape.Table.Cell(row, col).Shape.TextFrame2.TextRange.Text = newText;
                        break;
                    }
                }
                else
                {
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((string)shape.Name != shapeName) continue;
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
                prs.Save();
            }
            finally { if (ownedByUs) SafeClose(prs); }
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
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((string)shape.Name != tableName) continue;
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
                foreach (dynamic shape in slide.Shapes)
                {
                    if ((string)shape.Name != shapeName) continue;
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
            dynamic prs = pptApp.Presentations.Open(
                presentationPath, MsoTrue, MsoFalse, MsoFalse);
            try
            {
                string tempFile = Path.Combine(
                    Path.GetTempPath(),
                    $"bckppt_thumb_{Path.GetFileNameWithoutExtension(presentationPath)}_{slideIndex}.png");
                // Export(path, filterName, scaleWidth, scaleHeight)
                prs.Slides[slideIndex].Export(tempFile, "PNG", 960, 540);
                return (string?)tempFile;
            }
            catch { return (string?)null; }
            finally { SafeClose(prs); }
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
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((string)shape.Name != tableName) continue;
                        if ((int)shape.HasTable != MsoTrue) break;
                        tf = shape.Table.Cell(row, col).Shape.TextFrame2;
                        break;
                    }
                }
                else
                {
                    foreach (dynamic shape in slide.Shapes)
                    {
                        if ((string)shape.Name != shapeName) continue;
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
