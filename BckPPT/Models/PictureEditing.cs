namespace BckPPT.Models;

/// <summary>그림을 칸에 맞추는 방법. cover: 칸을 채우고 넘는 쪽을 가운데 기준으로 자른다. contain: 비율 유지, 칸 안 가운데. stretch: 칸 크기로 늘린다.</summary>
public enum PictureFit { Cover, Contain, Stretch }

/// <summary>원본 그림 기준으로 잘라 낼 비율(0~1). 음수면 그림 밖 여백(PowerPoint가 허용한다).</summary>
public record CropFractions(float Left = 0, float Top = 0, float Right = 0, float Bottom = 0)
{
    public static readonly CropFractions None = new();
    public bool IsNone => Math.Abs(Left) < 1e-4 && Math.Abs(Top) < 1e-4 && Math.Abs(Right) < 1e-4 && Math.Abs(Bottom) < 1e-4;

    /// <summary>「L10% T0% R10% B0%」</summary>
    public override string ToString() => $"L{Left * 100:0.#}% T{Top * 100:0.#}% R{Right * 100:0.#}% B{Bottom * 100:0.#}%";
}

public readonly record struct BoxF(float Left, float Top, float Width, float Height);

/// <summary>그림 넣기 요청 한 장. 칸(Left·Top·Width·Height, pt)에 Fit으로 맞추고, Crop은 먼저 원본에서 잘라 둘 부분.</summary>
public record PictureRequest(
    string ImagePath,
    float Left, float Top, float Width, float Height,
    PictureFit Fit = PictureFit.Cover,
    CropFractions? Crop = null,
    string? Name = null);

/// <summary>넣은 그림. 주소는 list_shapes와 같은 규약, 크기는 실제 도형 크기(pt), Crop은 원본 기준 최종 자른 비율.</summary>
public record PlacedPicture(string Name, string Address, float Left, float Top, float Width, float Height, CropFractions Crop,
    int? PixelWidth = null, int? PixelHeight = null);

/// <summary>목록에 붙는 그림 정보. FullWidth/Height는 자르기 전 그림 전체가 지금 배율로 차지하는 크기(pt).</summary>
public record PictureInfo(CropFractions Crop, float FullWidth, float FullHeight, int? PixelWidth = null, int? PixelHeight = null);

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
