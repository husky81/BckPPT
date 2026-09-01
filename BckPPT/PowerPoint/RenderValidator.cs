using BckPPT.Services;
using BckPPT.Infrastructure;

namespace BckPPT.PowerPoint;

public sealed class RenderValidator(IPptComService pptService, TemporaryArtifactStore artifacts)
{
    private const int MaxAutomaticCorrections = 5;

    public async Task<LayoutValidationResult> ValidateAndCorrectAsync(
        string presentationPath,
        IReadOnlySet<int> targetSlides,
        CancellationToken cancellationToken)
    {
        var before = await AnalyzeIssuesAsync(presentationPath, targetSlides, cancellationToken);
        var corrected = 0;
        foreach (var issue in before.Take(MaxAutomaticCorrections))
        {
            cancellationToken.ThrowIfCancellationRequested();
            await pptService.FixShapeAsync(presentationPath, issue.SlideIndex, issue.ShapeName);
            corrected++;
        }
        var remaining = corrected == 0
            ? before
            : await AnalyzeIssuesAsync(presentationPath, targetSlides, cancellationToken);
        return new LayoutValidationResult(before, remaining, corrected);
    }

    public async Task<IReadOnlyList<string>> RenderSlidesAsync(
        string presentationPath,
        IReadOnlySet<int> targetSlides,
        CancellationToken cancellationToken)
    {
        var paths = new List<string>();
        foreach (var slideIndex in targetSlides.Order())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = await pptService.GetSlideThumbnailAsync(presentationPath, slideIndex);
            if (!string.IsNullOrWhiteSpace(path)) paths.Add(artifacts.Track(path));
        }
        return paths;
    }

    private async Task<IReadOnlyList<LayoutIssue>> AnalyzeIssuesAsync(
        string path, IReadOnlySet<int> slides, CancellationToken cancellationToken)
    {
        var analysis = await Task.Run(() => pptService.AnalyzePresentation(path, cancellationToken: cancellationToken), cancellationToken);
        return analysis.Shapes
            .Where(shape => slides.Contains(shape.SlideIndex) && shape.IsOverflow)
            .Select(shape => new LayoutIssue(shape.SlideIndex, shape.ShapeName, shape.OverflowAmount, shape.LineCount))
            .ToList();
    }
}
