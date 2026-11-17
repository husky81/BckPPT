namespace BckPPT.Models;

/// <summary>그림을 칸에 맞추는 방법. cover: 칸을 채우고 넘는 쪽을 가운데 기준으로 자른다. contain: 비율 유지, 칸 안 가운데. stretch: 칸 크기로 늘린다.</summary>
public enum PictureFit { Cover, Contain, Stretch }

/// <summary>원본 그림 기준으로 잘라 낼 비율(0~1). 음수면 그림 밖 여백(PowerPoint가 허용한다).</summary>
public record CropFractions(float Left = 0, float Top = 0, float Right = 0, float Bottom = 0)
{
    public static readonly CropFractions None = new();
    public bool IsNone => Math.Abs(Left) < 1e-4 && Math.Abs(Top) < 1e-4 && Math.Abs(Right) < 1e-4 && Math.Abs(Bottom) < 1e-4;

    /// <summary>「L10% T0% R10% B0%」</summary>
    public override string ToString() => $"L{Pct(Left)}% T{Pct(Top)}% R{Pct(Right)}% B{Pct(Bottom)}%";

    // 계산 끝자리(-0.00001)가 「-0%」로 보이지 않게 0.05% 미만은 0으로 쓴다.
    private static string Pct(float v) => Math.Abs(v) < 0.0005f ? "0" : (v * 100).ToString("0.#");
}

public readonly record struct BoxF(float Left, float Top, float Width, float Height);

/// <summary>원본 기준 한 점(0~1). 0,0이 왼쪽 위.</summary>
public readonly record struct FocusPoint(float X, float Y);

/// <summary>넣은 그림을 앞뒤 어디에 둘지. None이면 새 도형이 생기는 자리(맨 앞, 바꾸기면 원래 자리).</summary>
public enum PictureZ { None, Front, Back }

/// <summary>그림 테두리. Color는 "#RRGGBB".</summary>
public record PictureBorder(string Color, float Width);

/// <summary>
/// 그림 넣기 요청 한 장. 칸(Left·Top·Width·Height, pt)에 Fit으로 맞추고, Crop은 먼저 원본에서 잘라 둘 부분.
/// Focus·Zoom을 주면 cover로 맞추되 Crop 뒤 남은 부분에서 그 점이 가운데 오게 자른다.
/// </summary>
public record PictureRequest(
    string ImagePath,
    float Left, float Top, float Width, float Height,
    PictureFit Fit = PictureFit.Cover,
    CropFractions? Crop = null,
    string? Name = null)
{
    public FocusPoint? Focus { get; init; }
    public float? Zoom { get; init; }
    public PictureZ Z { get; init; }
    public PictureBorder? Border { get; init; }
    public bool Shadow { get; init; }
}

/// <summary>이미 넣은 그림의 자르기 바꾸기. Crop·Focus/Zoom·Reset 중 하나. KeepFrame이면 도형 자리·크기를 그대로 둔다.</summary>
public record PictureCropRequest(CropFractions? Crop = null, FocusPoint? Focus = null, float? Zoom = null, bool Reset = false, bool KeepFrame = true);

/// <summary>미리보기용 그림 원본. Bytes는 원본 파일 그대로(또는 PowerPoint에서 뽑은 PNG), Origin은 어디서 왔는지.</summary>
public record PictureSource(byte[] Bytes, string Origin, string? SourcePath, PictureInfo Info, BoxF Frame);

/// <summary>넣은 그림. 주소는 list_shapes와 같은 규약, 크기는 실제 도형 크기(pt), Crop은 원본 기준 최종 자른 비율.</summary>
public record PlacedPicture(string Name, string Address, float Left, float Top, float Width, float Height, CropFractions Crop,
    int? PixelWidth = null, int? PixelHeight = null);

/// <summary>목록에 붙는 그림 정보. FullWidth/Height는 자르기 전 그림 전체가 지금 배율로 차지하는 크기(pt).
/// SourcePath는 BckPPT가 넣을 때 태그에 적어 둔 원본 파일(또는 대체 텍스트에 남은 파일 이름).</summary>
public record PictureInfo(CropFractions Crop, float FullWidth, float FullHeight, int? PixelWidth = null, int? PixelHeight = null, string? SourcePath = null);

/// <summary>그림 맞추기 계산. PowerPoint 없이 시험할 수 있게 COM과 떼어 둔다.</summary>
public static class PictureFitter
{
    /// <summary>
    /// 원본 크기(sourceWidth×sourceHeight, 비율만 쓴다)의 그림을 box에 fit으로 맞출 때 도형 자리와 원본 기준 자를 비율.
    /// userCrop을 먼저 적용한 뒤 그 남은 부분을 맞춘다. cover는 남은 부분의 가운데를 남기고 더 자른다.
    /// </summary>
    public static (BoxF Shape, CropFractions Crop) Fit(float sourceWidth, float sourceHeight, CropFractions? userCrop, BoxF box, PictureFit fit)
    {
        if (sourceWidth <= 0 || sourceHeight <= 0) throw new ArgumentException("그림 크기가 0입니다.");
        if (box.Width <= 0 || box.Height <= 0) throw new ArgumentException("칸 크기는 0보다 커야 합니다.");
        var c = userCrop ?? CropFractions.None;
        float keepW = 1 - c.Left - c.Right, keepH = 1 - c.Top - c.Bottom;
        if (keepW <= 0.01f || keepH <= 0.01f) throw new ArgumentException("crop이 그림을 거의 다 잘라 냅니다(left+right, top+bottom은 1보다 작아야 합니다).");

        float srcW = sourceWidth * keepW, srcH = sourceHeight * keepH;
        float srcRatio = srcW / srcH, boxRatio = box.Width / box.Height;
        switch (fit)
        {
            case PictureFit.Stretch:
                return (box, c);
            case PictureFit.Contain:
            {
                float w = box.Width, h = box.Height;
                if (srcRatio > boxRatio) h = box.Width / srcRatio; else w = box.Height * srcRatio;
                return (new BoxF(box.Left + (box.Width - w) / 2, box.Top + (box.Height - h) / 2, w, h), c);
            }
            default:
            {
                if (srcRatio > boxRatio)
                {
                    // 남은 부분이 칸보다 납작하다: 좌우를 더 자른다(원본 비율로 환산).
                    float extra = keepW * (1 - boxRatio / srcRatio) / 2;
                    return (box, c with { Left = c.Left + extra, Right = c.Right + extra });
                }
                else
                {
                    float extra = keepH * (1 - srcRatio / boxRatio) / 2;
                    return (box, c with { Top = c.Top + extra, Bottom = c.Bottom + extra });
                }
            }
        }
    }

    /// <summary>
    /// focus·zoom 자르기. userCrop 뒤 남은 부분 안에서 칸 비율의 창을 잡는다. zoom 1이면 남은 부분을 겨우 채우는 가장 큰 창,
    /// 2면 그 절반 크기. 창 가운데를 focus(원본 기준, 비우면 남은 부분 가운데)에 두고, 가장자리를 넘으면 안쪽으로 민다.
    /// </summary>
    public static CropFractions FocusCrop(float sourceWidth, float sourceHeight, CropFractions? userCrop, float boxRatio, FocusPoint? focus, float zoom)
    {
        if (sourceWidth <= 0 || sourceHeight <= 0) throw new ArgumentException("그림 크기가 0입니다.");
        if (boxRatio <= 0) throw new ArgumentException("칸 크기는 0보다 커야 합니다.");
        if (zoom < 1) throw new ArgumentException("zoom은 1 이상입니다(1 = 칸을 겨우 채움).");
        var c = userCrop ?? CropFractions.None;
        float keepW = 1 - c.Left - c.Right, keepH = 1 - c.Top - c.Bottom;
        if (keepW <= 0.01f || keepH <= 0.01f) throw new ArgumentException("crop이 그림을 거의 다 잘라 냅니다.");
        // 남은 부분 크기(원본 비율 단위)에서 칸 비율의 가장 큰 창.
        float restW = sourceWidth * keepW, restH = sourceHeight * keepH;
        float winW = restW, winH = restW / boxRatio;
        if (winH > restH) { winH = restH; winW = restH * boxRatio; }
        float fw = winW / zoom / sourceWidth, fh = winH / zoom / sourceHeight;   // 원본 기준 비율
        float minX = c.Left, maxX = 1 - c.Right, minY = c.Top, maxY = 1 - c.Bottom;
        var f = focus ?? new FocusPoint((minX + maxX) / 2, (minY + maxY) / 2);
        float cx = Math.Clamp(f.X, minX + fw / 2, maxX - fw / 2), cy = Math.Clamp(f.Y, minY + fh / 2, maxY - fh / 2);
        return new CropFractions(cx - fw / 2, cy - fh / 2, 1 - (cx + fw / 2), 1 - (cy + fh / 2));
    }

    /// <summary>넣기 요청 한 장의 도형 자리와 자를 비율. Focus·Zoom이 있으면 cover + FocusCrop.</summary>
    public static (BoxF Shape, CropFractions Crop) Place(PictureRequest r, float sourceWidth, float sourceHeight)
    {
        var box = new BoxF(r.Left, r.Top, r.Width, r.Height);
        if (r.Focus is null && r.Zoom is null) return Fit(sourceWidth, sourceHeight, r.Crop, box, r.Fit);
        if (r.Fit != PictureFit.Cover) throw new ArgumentException("focus·zoom은 fit=cover일 때만 씁니다.");
        if (box.Width <= 0 || box.Height <= 0) throw new ArgumentException("칸 크기는 0보다 커야 합니다.");
        return (box, FocusCrop(sourceWidth, sourceHeight, r.Crop, box.Width / box.Height, r.Focus, r.Zoom ?? 1));
    }

    /// <summary>
    /// 넣은 그림의 자르기를 바꿀 때 새 도형 자리와 자를 비율. frame은 지금 도형 자리, fullWidth/Height는 지금 배율의 전체 그림 크기(pt),
    /// current는 지금 자른 비율. KeepFrame이면 frame을 그대로 두고(crop은 cover로 맞춤, reset은 frame 안에 contain),
    /// 아니면 지금 배율·그림 자리를 지키고 도형이 새 자른 영역을 따라간다(PowerPoint 자르기 도구와 같다).
    /// </summary>
    public static (BoxF Shape, CropFractions Crop) Recrop(BoxF frame, float fullWidth, float fullHeight, CropFractions current, PictureCropRequest req)
    {
        if (fullWidth <= 0 || fullHeight <= 0) throw new ArgumentException("그림 크기를 읽지 못했습니다.");
        if (req.KeepFrame)
        {
            if (req.Reset) return Fit(fullWidth, fullHeight, null, frame, PictureFit.Contain);
            if (req.Focus is not null || req.Zoom is not null)
                return (frame, FocusCrop(fullWidth, fullHeight, req.Crop, frame.Width / frame.Height, req.Focus, req.Zoom ?? 1));
            return Fit(fullWidth, fullHeight, req.Crop ?? CropFractions.None, frame, PictureFit.Cover);
        }
        // 그림 전체의 왼쪽 위(pt): 지금 자른 비율만큼 도형 왼쪽 위에서 물러난 자리.
        float picLeft = frame.Left - current.Left * fullWidth, picTop = frame.Top - current.Top * fullHeight;
        CropFractions crop = req.Reset ? CropFractions.None
            : req.Focus is not null || req.Zoom is not null ? FocusCrop(fullWidth, fullHeight, req.Crop, frame.Width / frame.Height, req.Focus, req.Zoom ?? 1)
            : req.Crop ?? CropFractions.None;
        var shape = new BoxF(picLeft + crop.Left * fullWidth, picTop + crop.Top * fullHeight,
            fullWidth * (1 - crop.Left - crop.Right), fullHeight * (1 - crop.Top - crop.Bottom));
        return (shape, crop);
    }

    /// <summary>칸을 rows×columns 격자로 나눈다(왼쪽 위부터 줄 단위). 사이 간격 gap.</summary>
    public static IReadOnlyList<BoxF> Grid(BoxF box, int rows, int columns, float gap)
    {
        if (rows < 1 || columns < 1) throw new ArgumentException("rows·columns는 1 이상입니다.");
        float w = (box.Width - gap * (columns - 1)) / columns, h = (box.Height - gap * (rows - 1)) / rows;
        if (w <= 1 || h <= 1) throw new ArgumentException("칸이 너무 좁아 나눌 수 없습니다(gap을 줄이세요).");
        var cells = new List<BoxF>(rows * columns);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < columns; c++)
                cells.Add(new BoxF(box.Left + c * (w + gap), box.Top + r * (h + gap), w, h));
        return cells;
    }

    /// <summary>칸을 가로로 n등분한다(사이 간격 gap). 세로가 더 길면 세로로 나눈다.</summary>
    public static IReadOnlyList<BoxF> Split(BoxF box, int count, float gap, bool? horizontal = null)
    {
        if (count < 1) throw new ArgumentException("그림이 한 장 이상이어야 합니다.");
        bool across = horizontal ?? box.Width >= box.Height;
        float total = across ? box.Width : box.Height;
        float each = (total - gap * (count - 1)) / count;
        if (each <= 1) throw new ArgumentException("칸이 너무 좁아 나눌 수 없습니다(gap을 줄이세요).");
        return Enumerable.Range(0, count)
            .Select(i => across
                ? new BoxF(box.Left + i * (each + gap), box.Top, each, box.Height)
                : new BoxF(box.Left, box.Top + i * (each + gap), box.Width, each))
            .ToList();
    }

    /// <summary>PowerPoint Crop 개체 값(전체 그림 크기·그림 가운데의 도형 가운데 대비 위치·도형 크기)에서 원본 기준 자른 비율을 구한다.</summary>
    public static CropFractions FromCropObject(float pictureWidth, float pictureHeight, float offsetX, float offsetY, float shapeWidth, float shapeHeight)
    {
        if (pictureWidth <= 0 || pictureHeight <= 0) return CropFractions.None;
        float sumX = 1 - shapeWidth / pictureWidth, diffX = 2 * offsetX / pictureWidth;
        float sumY = 1 - shapeHeight / pictureHeight, diffY = 2 * offsetY / pictureHeight;
        return new CropFractions((sumX - diffX) / 2, (sumY - diffY) / 2, (sumX + diffX) / 2, (sumY + diffY) / 2);
    }

    /// <summary>자른 비율과 도형 자리로 Crop 개체에 넣을 값(전체 그림 크기, 그림 가운데의 도형 가운데 대비 위치).</summary>
    public static (float PictureWidth, float PictureHeight, float OffsetX, float OffsetY) ToCropObject(BoxF shape, CropFractions crop)
    {
        float pw = shape.Width / (1 - crop.Left - crop.Right), ph = shape.Height / (1 - crop.Top - crop.Bottom);
        return (pw, ph, pw * (crop.Right - crop.Left) / 2, ph * (crop.Bottom - crop.Top) / 2);
    }
}
