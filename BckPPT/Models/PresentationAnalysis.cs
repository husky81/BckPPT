namespace CellBox2.BckPPT.Models;

public record PresentationAnalysis(
    IReadOnlyList<SlideListItem> Slides,
    IReadOnlyList<PptSlideInfo>  Shapes,
    IReadOnlyList<PptFontInfo>   Fonts);
