namespace CellBox2.BckPPT.Models;

public record SlideListItem(
    int SlideIndex,
    string Title,
    int ShapeCount,
    int OverflowCount,
    bool IsAnalyzing = false)
{
    /// <summary>SlideIndex == 0 이면 "전체" 항목.</summary>
    public bool IsAll => SlideIndex == 0;

    public string DisplayText => IsAll ? "전체"
        : string.IsNullOrWhiteSpace(Title)
            ? $"슬라이드 {SlideIndex}"
            : $"{SlideIndex}. {Title}";

    public string OverflowBadge => OverflowCount > 0 ? $"\u26a0 ({OverflowCount})" : "";
    public bool HasOverflow => OverflowCount > 0;
}
