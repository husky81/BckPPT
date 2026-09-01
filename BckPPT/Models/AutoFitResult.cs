namespace CellBox2.BckPPT.Models;

public record AutoFitDetail(
    int    SlideIndex,
    string ShapeName,
    float  OriginalFontSize,
    float  FinalFontSize,
    bool   WasTrimmed);

public record AutoFitResult(
    int                         TotalFixed,
    IReadOnlyList<AutoFitDetail> Details);
