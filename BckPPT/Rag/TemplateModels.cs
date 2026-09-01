namespace BckPPT.Rag;

public sealed record IndexedSlide(
    string FilePath, int SlideIndex, string FileHash, DateTime FileModifiedUtc,
    string Category, string Title, string Text, int ParagraphCount,
    int TextShapeCount, int ImageCount, int TableCount);

public sealed record TemplateCandidate(
    string FilePath, int SlideIndex, string Category, string Title,
    string Text, double Score, string Reason);

public enum LayoutDecisionKind { CurrentDocument, TemplateLibrary, NewLayout }

public sealed record LayoutDecision(
    LayoutDecisionKind Kind, TemplateCandidate? Candidate, double Confidence, string Reason);
