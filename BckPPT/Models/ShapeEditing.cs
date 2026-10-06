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
    string TextPreview);

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

public enum TextShapeKind { Rectangle, RoundedRectangle }

/// <summary>
/// 텍스트 도형 추가 요청. <see cref="AutoWidth"/>면 글자 폭에 맞춰 너비를 정하고,
/// <see cref="RightEdge"/>가 있으면 늘어난 뒤 오른쪽 끝을 그 위치에 맞춘다(범례·배지 정렬용).
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
    float         Padding    = 8f);

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
