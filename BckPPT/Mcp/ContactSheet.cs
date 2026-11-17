using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace BckPPT.Mcp;

/// <summary>여러 쪽 PNG를 한 장(축소판)에 모은다. 칸마다 왼쪽 위에 쪽 번호를 붙인다.</summary>
public static class ContactSheet
{
    public const int CellWidth = 320, Gap = 8, LabelHeight = 20;

    public static byte[] Compose(IReadOnlyList<(int Slide, string PngPath)> pages, int columns)
    {
        if (pages.Count == 0) throw new ArgumentException("모을 쪽이 없습니다.");
        var images = pages.Select(p => (p.Slide, Image: Image.FromFile(p.PngPath))).ToList();
        try
        {
            var first = images[0].Image;
            int cellHeight = (int)Math.Round(CellWidth * (double)first.Height / first.Width);
            columns = Math.Min(columns, images.Count);
            int rows = (images.Count + columns - 1) / columns;
            int width = Gap + columns * (CellWidth + Gap);
            int height = Gap + rows * (LabelHeight + cellHeight + Gap);
            using var sheet = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(sheet))
            {
                g.Clear(Color.FromArgb(232, 232, 232));
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                using var font = new Font("Malgun Gothic", 10f, FontStyle.Bold);
                using var border = new Pen(Color.FromArgb(160, 160, 160));
                for (int i = 0; i < images.Count; i++)
                {
                    int x = Gap + i % columns * (CellWidth + Gap);
                    int y = Gap + i / columns * (LabelHeight + cellHeight + Gap);
                    g.DrawString($"{images[i].Slide}쪽", font, Brushes.Black, x, y + 2);
                    var rect = new Rectangle(x, y + LabelHeight, CellWidth, cellHeight);
                    g.DrawImage(images[i].Image, rect);
                    g.DrawRectangle(border, rect);
                }
            }
            using var ms = new MemoryStream();
            sheet.Save(ms, ImageFormat.Png);
            return ms.ToArray();
        }
        finally { foreach (var (_, image) in images) image.Dispose(); }
    }
}
