using System.Text.Json.Serialization;

namespace BckPPT.Models;

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ShapeRole
{
    Unknown,
    // 텍스트
    Chapter, Title,
    Outline1, Outline2, Outline3, Outline4, Outline5,
    Outline6, Outline7, Outline8, Outline9,
    ListItem, Body, Caption, PageNumber, Footer,
    // 이미지
    ContentImage, DecorativeImage, BulletIcon
}

public static class ShapeRoleExtensions
{
    public static string ToAbbr(this ShapeRole role) => role switch
    {
        ShapeRole.Chapter        => "단원",
        ShapeRole.Title          => "제목",
        ShapeRole.Outline1       => "개요1",
        ShapeRole.Outline2       => "개요2",
        ShapeRole.Outline3       => "개요3",
        ShapeRole.Outline4       => "개요4",
        ShapeRole.Outline5       => "개요5",
        ShapeRole.Outline6       => "개요6",
        ShapeRole.Outline7       => "개요7",
        ShapeRole.Outline8       => "개요8",
        ShapeRole.Outline9       => "개요9",
        ShapeRole.ListItem       => "목록",
        ShapeRole.Body           => "본문",
        ShapeRole.Caption        => "PicT",
        ShapeRole.PageNumber     => "쪽번",
        ShapeRole.Footer         => "하단",
        ShapeRole.ContentImage   => "그림",
        ShapeRole.DecorativeImage => "장식",
        ShapeRole.BulletIcon     => "아콘",
        _                        => "?"
    };

    public static string ToDisplayName(this ShapeRole role) => role switch
    {
        ShapeRole.Chapter         => "단원",
        ShapeRole.Title           => "제목",
        ShapeRole.Outline1        => "개요 1",
        ShapeRole.Outline2        => "개요 2",
        ShapeRole.Outline3        => "개요 3",
        ShapeRole.Outline4        => "개요 4",
        ShapeRole.Outline5        => "개요 5",
        ShapeRole.Outline6        => "개요 6",
        ShapeRole.Outline7        => "개요 7",
        ShapeRole.Outline8        => "개요 8",
        ShapeRole.Outline9        => "개요 9",
        ShapeRole.ListItem        => "목록 항목",
        ShapeRole.Body            => "본문",
        ShapeRole.Caption         => "그림 제목",
        ShapeRole.PageNumber      => "페이지 번호",
        ShapeRole.Footer          => "바닥글",
        ShapeRole.ContentImage    => "콘텐츠 이미지",
        ShapeRole.DecorativeImage => "장식용 이미지",
        ShapeRole.BulletIcon      => "불릿 아이콘",
        _                         => "미분류"
    };

    public static bool IsImageRole(this ShapeRole role)
        => role is ShapeRole.ContentImage or ShapeRole.DecorativeImage or ShapeRole.BulletIcon;

    public static bool IsTextRole(this ShapeRole role)
        => role != ShapeRole.Unknown && !role.IsImageRole();
}

public record ShapeRoleLabel(
    string    PresentationName,
    int       SlideIndex,
    string    ShapeName,
    ShapeRole Role,
    DateTime  LabeledAt,
    // feature 스냅샷
    string    Category,
    float     FontSize,
    bool      IsBold,
    float     WidthRatio,
    float     HeightRatio,
    float     TopRatio,
    float     LeftRatio,
    int       LineCount,
    bool      HasBullet,
    int       HalfWidthLength);
