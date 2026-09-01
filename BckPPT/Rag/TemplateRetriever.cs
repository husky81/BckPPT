using System.Text.RegularExpressions;

namespace BckPPT.Rag;

public sealed class TemplateRetriever(TemplateRepository repository)
{
    public IReadOnlyList<TemplateCandidate> Search(string query, int limit = 5)
    {
        var queryTokens = Tokenize(query);
        if (queryTokens.Count == 0) return [];
        return repository.ReadAll()
            .Select(slide => ToCandidate(slide, queryTokens))
            .Where(candidate => candidate.Score > 0)
            .OrderByDescending(candidate => candidate.Score)
            .ThenBy(candidate => candidate.FilePath, StringComparer.OrdinalIgnoreCase)
            .Take(limit)
            .ToList();
    }

    private static TemplateCandidate ToCandidate(IndexedSlide slide, HashSet<string> queryTokens)
    {
        var titleTokens = Tokenize(slide.Title);
        var allTokens = Tokenize(slide.Title + " " + slide.Text + " " + slide.Category);
        var overlap = queryTokens.Count(token => allTokens.Contains(token));
        var titleOverlap = queryTokens.Count(token => titleTokens.Contains(token));
        var coverage = (double)overlap / queryTokens.Count;
        var titleBoost = (double)titleOverlap / queryTokens.Count * 0.25;
        var reuseBoost = slide.TextShapeCount is >= 2 and <= 12 ? 0.05 : 0;
        var score = Math.Min(1, coverage * 0.7 + titleBoost + reuseBoost);
        var reason = $"키워드 {overlap}/{queryTokens.Count} 일치 · {slide.Category} · 텍스트 영역 {slide.TextShapeCount}개";
        return new TemplateCandidate(slide.FilePath, slide.SlideIndex, slide.Category, slide.Title, slide.Text, score, reason);
    }

    internal static HashSet<string> Tokenize(string text) => Regex.Matches(text.ToLowerInvariant(), @"[\p{L}\p{N}]{2,}")
        .Select(match => match.Value).ToHashSet(StringComparer.OrdinalIgnoreCase);
}
