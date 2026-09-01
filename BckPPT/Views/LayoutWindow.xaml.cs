using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Views;

public partial class LayoutWindow : Window
{
    private readonly string _presentationPath;
    private readonly IReadOnlyList<SlideListItem> _slides; // IsAll 제외
    private int _currentIndex; // _slides 내 인덱스
    private bool _isLoading;

    public LayoutWindow(string presentationPath,
        IReadOnlyList<SlideListItem> slides, int initialSlideIndex)
    {
        InitializeComponent();
        _presentationPath = presentationPath;
        _slides = slides.Where(s => !s.IsAll).ToList();
        _currentIndex = _slides.ToList().FindIndex(s => s.SlideIndex == initialSlideIndex);
        if (_currentIndex < 0) _currentIndex = 0;

        SourceInitialized += (_, _) =>
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            int value = 1;
            DwmSetWindowAttribute(hwnd, 20, ref value, sizeof(int));
        };

        Loaded += async (_, _) => await LoadCurrentSlideAsync();
    }

    // ── 슬라이드 로드 ────────────────────────────────────────────

    private async Task LoadCurrentSlideAsync()
    {
        if (_isLoading || _slides.Count == 0) return;
        _isLoading = true;

        var slide = _slides[_currentIndex];
        SlideTitle.Text = $"{slide.SlideIndex}. {slide.Title}";
        SlideCounter.Text = $"{_currentIndex + 1} / {_slides.Count}";
        Title = $"레이아웃 — {slide.DisplayText}";

        try
        {
            var svc = new PptComService();
            var layout = await Task.Run(() =>
                svc.GetSlideLayout(_presentationPath, slide.SlideIndex));
            DrawLayout(layout);
        }
        catch (Exception ex)
        {
            LayoutCanvas.Children.Clear();
            SlideTitle.Text = $"오류: {ex.Message}";
        }
        finally { _isLoading = false; }
    }

    // ── 마우스 휠 → 슬라이드 이동 ────────────────────────────────

    private async void Window_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (_isLoading || _slides.Count == 0) return;

        int next = _currentIndex + (e.Delta < 0 ? 1 : -1);
        if (next < 0 || next >= _slides.Count) return;

        _currentIndex = next;
        e.Handled = true;
        await LoadCurrentSlideAsync();
    }

    // ── PPT 슬라이드 이동 ──────────────────────────────────────────

    private void GoToSlide_Click(object sender, RoutedEventArgs e)
    {
        if (_slides.Count == 0) return;
        var slide = _slides[_currentIndex];
        try { new PptComService().NavigateToSlide(_presentationPath, slide.SlideIndex); }
        catch { }
    }

    // ── 레이아웃 렌더링 ──────────────────────────────────────────

    private void DrawLayout(SlideLayout layout)
    {
        LayoutCanvas.Children.Clear();
        LayoutCanvas.Width  = layout.SlideWidth;
        LayoutCanvas.Height = layout.SlideHeight;

        LayoutCanvas.Background = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));

        var checkerBrush = CreateCheckerBrush();
        var textBrush    = new SolidColorBrush(Color.FromRgb(0x44, 0x44, 0x44));
        var otherBrush   = new SolidColorBrush(Color.FromRgb(0x5A, 0x5A, 0x5A));
        var strokeBrush  = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66));

        foreach (var shape in layout.Shapes)
        {
            var fill = shape.Category switch
            {
                ShapeCategory.Image => (Brush)checkerBrush,
                ShapeCategory.Text  => textBrush,
                _                   => otherBrush,
            };

            string rotationLabel = Math.Abs(shape.Rotation) > 0.5f ? $"  ↻{shape.Rotation:F0}°" : "";
            var rect = new Rectangle
            {
                Width           = Math.Max(shape.Width,  1),
                Height          = Math.Max(shape.Height, 1),
                Fill            = fill,
                Stroke          = strokeBrush,
                StrokeThickness = 0.8,
                RadiusX         = 2,
                RadiusY         = 2,
                ToolTip         = $"{shape.Name}\n{shape.Category}  {shape.Width:F0} × {shape.Height:F0}{rotationLabel}",
            };

            if (Math.Abs(shape.Rotation) > 0.5f)
            {
                rect.RenderTransformOrigin = new System.Windows.Point(0.5, 0.5);
                rect.RenderTransform = new RotateTransform(shape.Rotation);
            }

            Canvas.SetLeft(rect, shape.Left);
            Canvas.SetTop(rect,  shape.Top);
            LayoutCanvas.Children.Add(rect);

            if (shape.Category == ShapeCategory.Text && shape.Height > 8 && Math.Abs(shape.Rotation) < 0.5f)
                AddTextLines(shape);
        }
    }

    // BreakType: 0=None(마지막줄), 1=AutoWrap, 2=Enter, 3=ShiftEnter

    private void AddTextLines(ShapeLayoutInfo shape)
    {
        var lineBrush   = new SolidColorBrush(Color.FromRgb(0xBB, 0xBB, 0xBB));
        var bulletBrush = new SolidColorBrush(Color.FromRgb(0xE0, 0xA0, 0x60));
        float margin    = 8f;
        float barH_line = 1.5f;   // 가로선 두께
        float thickBar  = 2.0f;   // 굵은 세로선 두께 (Enter, 문장 시작/끝)
        float thinBar   = 1.0f;   // 가는 세로선 두께 (Shift+Enter)
        float bulletSize = 4f;
        float bulletGap  = 6f;
        float maxWidth  = shape.Width - margin * 2;
        float fontScale = 0.55f;  // pt → layout 세로선 높이 스케일

        if (shape.TextLines is not { Count: > 0 } lines)
        {
            // 줄 데이터 없으면 폴백
            float lineSpacing = 14f;
            float fy = shape.Top + margin + 6;
            float maxY = shape.Top + shape.Height - margin;
            while (fy < maxY)
            {
                float ratio = 0.5f + 0.4f * ((fy.GetHashCode() & 0xFF) / 255f);
                var fallback = new Rectangle
                {
                    Width = Math.Max(maxWidth * ratio, 10), Height = barH_line,
                    Fill = lineBrush, RadiusX = 0.75, RadiusY = 0.75,
                };
                Canvas.SetLeft(fallback, shape.Left + margin);
                Canvas.SetTop(fallback, fy);
                LayoutCanvas.Children.Add(fallback);
                fy += lineSpacing;
            }
            return;
        }

        bool anyBullet = lines.Any(l => l.HasBullet);
        float bulletArea = anyBullet ? bulletSize + bulletGap : 0f;
        float textMaxW = maxWidth - bulletArea;

        // BoundHeight 기반 줄 간격
        float textH = shape.BoundHeight > shape.Height ? shape.BoundHeight : shape.Height;
        float spacing = lines.Count > 1
            ? Math.Clamp((textH - margin * 2) / lines.Count, 8f, 18f)
            : 14f;
        float y = shape.Top + margin + spacing * 0.4f;

        for (int i = 0; i < lines.Count; i++)
        {
            var tl = lines[i];
            float vBarH = Math.Max(tl.FontSize * fontScale, 5f);

            float textW = Math.Max(textMaxW * tl.WidthRatio, 6);
            float baseX = shape.Left + margin + bulletArea;

            float x = tl.Alignment switch
            {
                2 => baseX + (textMaxW - textW) / 2f,
                3 => baseX + (textMaxW - textW),
                _ => baseX,
            };

            float barCenterY = y + barH_line / 2f;

            // ── 불릿 점 ──
            if (tl.HasBullet)
            {
                var dot = new Ellipse { Width = bulletSize, Height = bulletSize, Fill = bulletBrush };
                Canvas.SetLeft(dot, shape.Left + margin);
                Canvas.SetTop(dot, barCenterY - bulletSize / 2f);
                LayoutCanvas.Children.Add(dot);
            }

            // ── 왼쪽 세로선 ──
            // 첫 줄 또는 이전 줄이 Enter/ShiftEnter → 세로선 표시
            int prevBreak = i == 0 ? 2 : lines[i - 1].BreakType; // 첫 줄은 문장 시작 = thick
            if (prevBreak >= 2) // Enter(2) 또는 ShiftEnter(3) 또는 첫 줄
            {
                float bw = prevBreak == 3 ? thinBar : thickBar;
                var leftBar = new Rectangle
                {
                    Width = bw, Height = vBarH, Fill = lineBrush,
                };
                Canvas.SetLeft(leftBar, x - bw);
                Canvas.SetTop(leftBar, barCenterY - vBarH / 2f);
                LayoutCanvas.Children.Add(leftBar);
            }

            // ── 가로선 ──
            var hBar = new Rectangle
            {
                Width = textW, Height = barH_line, Fill = lineBrush,
            };
            Canvas.SetLeft(hBar, x);
            Canvas.SetTop(hBar, y);
            LayoutCanvas.Children.Add(hBar);

            // ── 오른쪽 세로선 ──
            // Enter(2), ShiftEnter(3), 또는 마지막 줄(0) → 세로선 표시
            int curBreak = tl.BreakType;
            if (curBreak == 0 || curBreak >= 2)
            {
                float bw = curBreak == 3 ? thinBar : thickBar;
                var rightBar = new Rectangle
                {
                    Width = bw, Height = vBarH, Fill = lineBrush,
                };
                Canvas.SetLeft(rightBar, x + textW);
                Canvas.SetTop(rightBar, barCenterY - vBarH / 2f);
                LayoutCanvas.Children.Add(rightBar);
            }

            y += spacing;
        }
    }

    // ── 체커보드 브러시 ──────────────────────────────────────────

    private static DrawingBrush CreateCheckerBrush()
    {
        var dark  = new SolidColorBrush(Color.FromRgb(0x3A, 0x3A, 0x3A));
        var light = new SolidColorBrush(Color.FromRgb(0x4A, 0x4A, 0x4A));
        var group = new DrawingGroup();
        group.Children.Add(new GeometryDrawing(dark,  null, new RectangleGeometry(new Rect(0,  0,  20, 20))));
        group.Children.Add(new GeometryDrawing(light, null, new RectangleGeometry(new Rect(0,  0,  10, 10))));
        group.Children.Add(new GeometryDrawing(light, null, new RectangleGeometry(new Rect(10, 10, 10, 10))));
        return new DrawingBrush(group)
        {
            TileMode      = TileMode.Tile,
            Viewport      = new Rect(0, 0, 20, 20),
            ViewportUnits = BrushMappingMode.Absolute,
        };
    }

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
}
