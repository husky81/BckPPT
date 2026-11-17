using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using BckPPT.Models;

namespace BckPPT.Mcp;

/// <summary>
/// 그림 원본 전체를 줄여 그리고 0.1 간격 격자·눈금(원본 기준 0~1)을 얹는다. 지금 자른 영역(빨간 실선, 밖은 어둡게)과
/// 바꿔 볼 영역(노란 점선)을 그려 crop·focus 값을 눈으로 정하게 한다.
/// </summary>
public static class PicturePreview
{
    public const int LongSide = 900, MarginLeft = 34, MarginTop = 22, MarginRight = 12, MarginBottom = 12;

    public static (byte[] Png, int PixelWidth, int PixelHeight) Render(byte[] image, CropFractions? current, CropFractions? proposed)
    {
        Image source;
        try { source = Image.FromStream(new MemoryStream(image)); }
        catch (ArgumentException) { throw new InvalidOperationException("이 그림 형식은 미리보기를 그릴 수 없습니다(png·jpg·gif·bmp·tif·emf·wmf만)."); }
        using (source)
        {
            int pw = source.Width, ph = source.Height;
            float scale = (float)LongSide / Math.Max(pw, ph);
            int w = Math.Max(1, (int)Math.Round(pw * scale)), h = Math.Max(1, (int)Math.Round(ph * scale));
            using var canvas = new Bitmap(MarginLeft + w + MarginRight, MarginTop + h + MarginBottom, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(canvas))
            {
                g.Clear(Color.FromArgb(245, 245, 245));
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                var area = new RectangleF(MarginLeft, MarginTop, w, h);
                g.DrawImage(source, area);

                RectangleF Rect(CropFractions c) => new(
                    area.Left + c.Left * w, area.Top + c.Top * h, (1 - c.Left - c.Right) * w, (1 - c.Top - c.Bottom) * h);

                if (current is { IsNone: false } cur)
                {
                    // 잘려 나간 부분을 어둡게.
                    var keep = Rect(cur);
                    using var region = new Region(area);
                    region.Exclude(keep);
                    using var dim = new SolidBrush(Color.FromArgb(120, 0, 0, 0));
                    g.FillRegion(dim, region);
                }

                // 격자: 0.1마다 얇게, 0.5는 굵게. 밝은 그림·어두운 그림 모두에서 보이게 검은 테 위에 흰 선.
                using var labelFont = new Font("Malgun Gothic", 8.5f);
                using var shade = new Pen(Color.FromArgb(90, 0, 0, 0), 3f);
                using var thin = new Pen(Color.FromArgb(150, 255, 255, 255), 1f);
                using var half = new Pen(Color.FromArgb(220, 255, 255, 255), 1.6f);
                for (int i = 0; i <= 10; i++)
                {
                    float x = area.Left + w * i / 10f, y = area.Top + h * i / 10f;
                    var pen = i == 5 ? half : thin;
                    if (i is > 0 and < 10)
                    {
                        g.DrawLine(shade, x, area.Top, x, area.Bottom); g.DrawLine(pen, x, area.Top, x, area.Bottom);
                        g.DrawLine(shade, area.Left, y, area.Right, y); g.DrawLine(pen, area.Left, y, area.Right, y);
                    }
                    string label = i == 0 ? "0" : i == 10 ? "1" : $".{i}";
                    var size = g.MeasureString(label, labelFont);
                    g.DrawString(label, labelFont, Brushes.Black, x - size.Width / 2, area.Top - size.Height - 2);
                    g.DrawString(label, labelFont, Brushes.Black, area.Left - size.Width - 3, y - size.Height / 2);
                }

                if (current is { IsNone: false } c1)
                {
                    using var red = new Pen(Color.FromArgb(230, 30, 30), 3f);
                    var r = Rect(c1);
                    g.DrawRectangle(red, r.X, r.Y, r.Width, r.Height);
                }
                if (proposed is { } c2)
                {
                    using var yellow = new Pen(Color.FromArgb(255, 200, 0), 3f) { DashStyle = DashStyle.Dash };
                    var r = Rect(c2);
                    g.DrawRectangle(yellow, r.X, r.Y, r.Width, r.Height);
                }
                using var frame = new Pen(Color.FromArgb(120, 120, 120));
                g.DrawRectangle(frame, area.X, area.Y, area.Width, area.Height);
            }
            using var ms = new MemoryStream();
            canvas.Save(ms, ImageFormat.Png);
            return (ms.ToArray(), pw, ph);
        }
    }

    /// <summary>자른 영역을 focus·zoom으로 나타낸다(그 영역을 칸 비율 boxRatio로 다시 잡을 때의 값).</summary>
    public static (FocusPoint Focus, float Zoom) AsFocus(CropFractions crop, float sourceWidth, float sourceHeight, float boxRatio)
    {
        var full = PictureFitter.FocusCrop(sourceWidth, sourceHeight, null, boxRatio, null, 1);
        float zoom = (1 - full.Left - full.Right) / Math.Max(0.0001f, 1 - crop.Left - crop.Right);
        return (new FocusPoint((crop.Left + 1 - crop.Right) / 2, (crop.Top + 1 - crop.Bottom) / 2), zoom);
    }

    /// <summary>「x 0.10~0.90, y 0.05~0.95」</summary>
    public static string Range(CropFractions c) =>
        $"x {c.Left:0.00}~{1 - c.Right:0.00}, y {c.Top:0.00}~{1 - c.Bottom:0.00}";
}
