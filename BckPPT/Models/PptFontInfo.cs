namespace CellBox2.BckPPT.Models;

public record PptFontInfo(
    int    SlideIndex,
    string ShapeName,
    string TextPreview,
    string FontName,
    float  FontSize,
    bool   IsBold,
    bool   IsAllowed)
{
    public string FontDisplay  => $"{FontName} {FontSize}pt{(IsBold ? " B" : "")}";
    public string AllowedIcon  => IsAllowed ? "\u2705" : "\u26a0\ufe0f";
}
