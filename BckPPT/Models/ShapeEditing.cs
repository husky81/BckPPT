using System.Globalization;

namespace BckPPT.Models;

/// <summary>그룹 안까지 펼친 도형 목록의 한 줄. <see cref="Address"/>는 다른 편집 API에 그대로 넘길 수 있다.</summary>
public record ShapeEntry(
    string Address,
    string Name,
    string GroupPath,   // "Group 171" 처럼 '/'로 이은 조상 그룹. 최상위면 빈 문자열.
    int    Type,        // MsoShapeType
    float  Left,
    float  Top,
    float  Width,
    float  Height,
    bool   HasText,
    string TextPreview,
    int    Id = 0);   // PowerPoint Shape.Id. 슬라이드 안에서 고유하고 복제해도 유지된다.

/// <summary>값이 있는 항목만 바꾼다. 크기를 먼저 바꾼 뒤 위치를 맞추므로 뒤집힌 도형도 제자리에 둔다.</summary>
public record ShapeGeometry(float? Left = null, float? Top = null, float? Width = null, float? Height = null);

/// <summary>값이 있는 항목만 도형 전체 텍스트에 적용한다. 색은 "#RRGGBB".</summary>
public record TextFormatPatch(
    float?  Size     = null,
    bool?   Bold     = null,
    string? Color    = null,
    float?  Spacing  = null,
    string? FontName = null);

/// <summary>도형 텍스트 첫 런의 서식. 색은 "#RRGGBB", 자간(Spacing)은 pt.</summary>
public record TextFormatInfo(string FontName, float Size, float Spacing, bool Bold, string Color);

public enum TextShapeKind { Rectangle, RoundedRectangle, Diamond, Ellipse, Triangle, Chevron, Pentagon, RightArrow }

public enum TextAlign { Left, Center, Right }

/// <summary>
/// 텍스트 도형 추가 요청. <see cref="AutoWidth"/>면 글자 폭에 맞춰 너비를 정하고,
/// <see cref="RightEdge"/>가 있으면 늘어난 뒤 오른쪽 끝을 그 위치에 맞춘다(범례·배지 정렬용).
/// <see cref="LineColor"/>가 없으면 테두리를 그리지 않는다.
/// </summary>
public record TextShapeRequest(
    string        Text,
    float         Left,
    float         Top,
    float         Width,
    float         Height,
    string?       Name       = null,
    TextShapeKind Kind       = TextShapeKind.Rectangle,
    string?       FillColor  = null,   // null이면 채우기 없음
    string        FontName   = "",     // 빈 값이면 테마 글꼴
    float         FontSize   = 10f,
    bool          Bold       = false,
    string        FontColor  = "#000000",
    bool          AutoWidth  = false,
    float?        RightEdge  = null,
    float         Padding    = 8f,
    TextAlign     Align      = TextAlign.Left,
    string?       LineColor  = null,
    float         LineWidth  = 0.75f);

public enum ArrowHeads { None, End, Start, Both }

/// <summary>직선(화살표). 좌표는 pt.</summary>
public record LineRequest(
    float      X1,
    float      Y1,
    float      X2,
    float      Y2,
    string     Color  = "#000000",
    float      Width  = 1f,
    ArrowHeads Arrows = ArrowHeads.None,
    bool       Dashed = false,
    string?    Name   = null);

/// <summary>병합할 셀 범위(1부터, 양 끝 포함).</summary>
public record TableMerge(int Row, int Column, int Row2, int Column2);

/// <summary>셀 하나의 서식. 값이 있는 항목만 표 기본값을 덮어쓴다.</summary>
public record TableCellStyle(int Row, int Column, string? Fill = null, string? FontColor = null, bool? Bold = null,
    TextAlign? Align = null, float? FontSize = null);

/// <summary>
/// 표 추가 요청. <see cref="Cells"/>는 행마다 열 값(병합으로 가려지는 칸은 빈 값). 머리 행은 <see cref="HeaderFill"/>,
/// 본문은 <see cref="BandFill"/>로 한 줄 걸러 칠한다. 테두리는 <see cref="BorderColor"/> 가는 선.
/// </summary>
public record TableRequest(
    float                                  Left,
    float                                  Top,
    float                                  Width,
    IReadOnlyList<IReadOnlyList<string>>   Cells,
    IReadOnlyList<float>?                  ColumnWidths = null,
    float                                  RowHeight    = 20f,
    float                                  FontSize     = 10f,
    string                                 FontName     = "",
    int                                    HeaderRows   = 1,
    string?                                HeaderFill   = "#1F3864",
    string                                 HeaderFontColor = "#FFFFFF",
    string                                 FontColor    = "#000000",
    string?                                BodyFill     = null,
    string?                                BandFill     = null,
    string?                                BorderColor  = "#BFBFBF",
    float                                  BorderWidth  = 0.5f,
    IReadOnlyList<TextAlign>?              ColumnAligns = null,
    IReadOnlyList<TableMerge>?             Merges       = null,
    IReadOnlyList<TableCellStyle>?         CellStyles   = null,
    IReadOnlyList<float>?                  RowHeights   = null,
    float                                  CellPadding  = 3f,
    string?                                Name         = null);

/// <summary>만든 표의 실제 배치. 행 높이는 글자에 맞춰 늘어날 수 있어 그린 뒤 읽은 값이다.</summary>
public record TableLayout(string Name, IReadOnlyList<float> RowTops, IReadOnlyList<float> RowHeights,
    IReadOnlyList<float> ColumnLefts, IReadOnlyList<float> ColumnWidths);

/// <summary>"#RRGGBB" ↔ PowerPoint COM RGB(BGR 순서 정수) 변환.</summary>
public static class OleColor
{
    public static int FromHex(string hex)
    {
        string h = hex.Trim().TrimStart('#');
        if (h.Length != 6 || !int.TryParse(h, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb))
            throw new FormatException($"색은 #RRGGBB 형식이어야 합니다: '{hex}'");
        int r = (rgb >> 16) & 0xFF, g = (rgb >> 8) & 0xFF, b = rgb & 0xFF;
        return r | (g << 8) | (b << 16);
    }

    public static string ToHex(int ole)
        => $"#{ole & 0xFF:X2}{(ole >> 8) & 0xFF:X2}{(ole >> 16) & 0xFF:X2}";
}
