namespace CellBox2.BckPPT.Models;

public record SlideAnalysisProgress(
    AnalysisProgressKind         Kind,
    int                          SlideIndex,
    int                          TotalSlides,
    SlideListItem?               Slide      = null,
    IReadOnlyList<PptSlideInfo>? ShapeBatch = null,
    IReadOnlyList<PptFontInfo>?  FontBatch  = null
);
