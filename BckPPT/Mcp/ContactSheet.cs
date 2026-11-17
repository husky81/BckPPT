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

    /// <summary>
    /// 후보 그림 여러 장을 한 장에 모은다. 칸마다 위에 이름표(번호·파일 이름·픽셀) 두 줄, 그림은 칸 안에 비율 유지로.
    /// 읽지 못한 그림은 칸에 「읽지 못함」을 쓴다.
    /// </summary>
    public static byte[] ComposeImages(IReadOnlyList<(string Title, string Detail, byte[]? Bytes)> items, int columns, int cellWidth = 300, int cellHeight = 225)
    {
        if (items.Count == 0) throw new ArgumentException("모을 그림이 없습니다.");
        const int label = 34;
        columns = Math.Clamp(Math.Min(columns, items.Count), 1, 8);
        int rows = (items.Count + columns - 1) / columns;
        using var sheet = new Bitmap(Gap + columns * (cellWidth + Gap), Gap + rows * (label + cellHeight + Gap), PixelFormat.Format24bppRgb);
        using (var g = Graphics.FromImage(sheet))
        {
            g.Clear(Color.FromArgb(232, 232, 232));
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            using var bold = new Font("Malgun Gothic", 9.5f, FontStyle.Bold);
            using var small = new Font("Malgun Gothic", 8.5f);
            using var border = new Pen(Color.FromArgb(160, 160, 160));
            using var cellBack = new SolidBrush(Color.White);
            var format = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap };
            for (int i = 0; i < items.Count; i++)
            {
                int x = Gap + i % columns * (cellWidth + Gap), y = Gap + i / columns * (label + cellHeight + Gap);
                g.DrawString(items[i].Title, bold, Brushes.Black, new RectangleF(x, y, cellWidth, 17), format);
                g.DrawString(items[i].Detail, small, Brushes.DimGray, new RectangleF(x, y + 16, cellWidth, 16), format);
                var cell = new Rectangle(x, y + label, cellWidth, cellHeight);
                g.FillRectangle(cellBack, cell);
                Image? image = null;
                try { if (items[i].Bytes is { } b) image = Image.FromStream(new MemoryStream(b)); } catch (ArgumentException) { }
                if (image is null) g.DrawString("읽지 못함", small, Brushes.Firebrick, x + 8, y + label + 8);
                else
                    using (image)
                    {
                        float scale = Math.Min((float)cellWidth / image.Width, (float)cellHeight / image.Height);
                        float w = image.Width * scale, h = image.Height * scale;
                        g.DrawImage(image, x + (cellWidth - w) / 2, y + label + (cellHeight - h) / 2, w, h);
                    }
                g.DrawRectangle(border, cell);
            }
        }
        using var ms = new MemoryStream();
        sheet.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }
}
