using CellBox2.BckPPT.Models;

namespace CellBox2.BckPPT.Services;

public interface IPptComService
{
    bool IsPowerPointRunning { get; }

    IReadOnlyList<PptPresentation> GetOpenPresentations();
    PresentationAnalysis AnalyzePresentation(
        string presentationPath,
        IProgress<SlideAnalysisProgress>? slideProgress = null,
        CancellationToken cancellationToken = default);

    Task<AutoFitResult> AutoFitAsync(string presentationPath, IProgress<string>? progress = null);
    Task FixShapeAsync(string presentationPath, int slideIndex, string shapeName,
        IProgress<string>? progress = null);

    Task NavigateToShapeAsync(string presentationPath, int slideIndex, string shapeName);
    void NavigateToSlide(string presentationPath, int slideIndex);
    void OpenPptxInPowerPoint(string filePath);

    int  ReplaceText(string presentationPath, string find, string replace,
        IReadOnlyList<(int slideIndex, string shapeName)>? targets = null);
    int  NormalizeFonts(string presentationPath);
    void ReplaceImage(string presentationPath, int slideIndex, float targetLeft, float targetTop, string newImagePath);
    void SetShapeText(string presentationPath, int slideIndex, string shapeName, string newText);

    /// <summary>설정된 ReplaceRules를 프레젠테이션 전체에 일괄 적용. 변경된 항목 목록 반환.</summary>
    IReadOnlyList<(int SlideIndex, string ShapeName, string OldText, string NewText)>
        ApplyRules(string presentationPath, IReadOnlyList<ReplaceRule> rules);

    /// <summary>단일 shape의 현재 텍스트·메트릭을 읽어 PptSlideInfo로 반환. 찾지 못하면 null.</summary>
    PptSlideInfo? GetShapeInfo(string presentationPath, int slideIndex, string shapeName);

    SlideLayout GetSlideLayout(string presentationPath, int slideIndex);

    /// <summary>슬라이드를 PNG 썸네일로 내보내고 임시 파일 경로를 반환. 실패 시 null.</summary>
    Task<string?> GetSlideThumbnailAsync(string presentationPath, int slideIndex);

    IReadOnlyList<RecentPptxFile> GetRecentFiles();

    /// <summary>텍스트 상자의 각 시각적 줄과 줄넘김 유형(자동/Enter/Shift+Enter)을 분석.</summary>
    IReadOnlyList<TextLineInfo> GetTextLineBreaks(string presentationPath, int slideIndex, string shapeName);
}
