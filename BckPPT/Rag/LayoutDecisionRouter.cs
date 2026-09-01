using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Rag;

public sealed class LayoutDecisionRouter(IPptComService pptService, TemplateRetriever retriever)
{
    private const double CurrentDocumentThreshold = 0.42;
    private const double LibraryThreshold = 0.32;

    public async Task<LayoutDecision> DecideAsync(
        string query, string currentDocumentPath, CancellationToken cancellationToken)
    {
        var analysis = await Task.Run(() => pptService.AnalyzePresentation(
            currentDocumentPath, cancellationToken: cancellationToken), cancellationToken);
        var current = SearchCurrent(query, currentDocumentPath, analysis).FirstOrDefault();
        if (current?.Score >= CurrentDocumentThreshold)
            return new LayoutDecision(LayoutDecisionKind.CurrentDocument, current, current.Score,
                "현재 문서의 디자인 문법과 내용 구조를 우선 재사용합니다.");

        var library = retriever.Search(query, 1).FirstOrDefault();
        if (library?.Score >= LibraryThreshold)
            return new LayoutDecision(LayoutDecisionKind.TemplateLibrary, library, library.Score,
                "개인 라이브러리에서 더 적합한 슬라이드를 찾았습니다.");

        return new LayoutDecision(LayoutDecisionKind.NewLayout, null,
            Math.Max(current?.Score ?? 0, library?.Score ?? 0), "재사용 기준 점수 미달로 신규 레이아웃이 필요합니다.");
    }

    private static IReadOnlyList<TemplateCandidate> SearchCurrent(
        string query, string path, PresentationAnalysis analysis)
    {
        var queryTokens = TemplateRetriever.Tokenize(query);
        return analysis.Shapes.GroupBy(shape => shape.SlideIndex).Select(group =>
        {
            var text = string.Join("\n", group.Select(shape => shape.FullText));
            var tokens = TemplateRetriever.Tokenize(text);
            var overlap = queryTokens.Count == 0 ? 0 : queryTokens.Count(token => tokens.Contains(token));
            var score = queryTokens.Count == 0 ? 0 : Math.Min(1, (double)overlap / queryTokens.Count * 0.8 + 0.08);
            return new TemplateCandidate(path, group.Key, "current", group.FirstOrDefault()?.TextPreview ?? "", text,
                score, $"현재 문서 키워드 {overlap}/{queryTokens.Count} 일치");
        }).OrderByDescending(candidate => candidate.Score).ToList();
    }
}
