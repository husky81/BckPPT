namespace BckPPT.PowerPoint;

public sealed record LayoutIssue(int SlideIndex, string ShapeName, float OverflowAmount, int LineCount);

public sealed record LayoutValidationResult(
    IReadOnlyList<LayoutIssue> BeforeCorrection,
    IReadOnlyList<LayoutIssue> RemainingIssues,
    int CorrectedCount)
{
    public bool IsValid => RemainingIssues.Count == 0;
}
