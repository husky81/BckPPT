namespace CellBox2.BckPPT.Models;

public enum ShapeCategory { Text, Image, Other }

/// <summary>텍스트 shape의 시각적 줄 하나에 대한 레이아웃 정보.</summary>
/// <param name="WidthRatio">줄 텍스트 폭 / shape 내부 폭 (0.0 ~ 1.0)</param>
/// <param name="Alignment">ppAlignLeft=1, Center=2, Right=3, Justify=4</param>
/// <param name="HasBullet">이 줄이 불릿 문단의 첫 줄인지</param>
/// <param name="FontSize">이 줄의 글꼴 크기 (pt)</param>
/// <param name="BreakType">이 줄 끝의 줄넘김 유형 (0=None, 1=AutoWrap, 2=Enter, 3=ShiftEnter)</param>
public record TextLineLayout(
    float WidthRatio, int Alignment,
    bool HasBullet = false, float FontSize = 12f, int BreakType = 0);

public record ShapeLayoutInfo(
    float Left, float Top, float Width, float Height,
    ShapeCategory Category,
    string Name,
    IReadOnlyList<TextLineLayout>? TextLines = null,
    float BoundHeight = 0f,
    float Rotation = 0f);

public record SlideLayout(
    float SlideWidth, float SlideHeight,
    IReadOnlyList<ShapeLayoutInfo> Shapes);
