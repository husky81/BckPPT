using BckPPT.Models;

namespace BckPPT.Services;

public interface IPptComService
{
    bool IsPowerPointRunning { get; }

    IReadOnlyList<PptPresentation> GetOpenPresentations();
    /// <summary>PowerPoint 활성 창의 문서와 보고 있는 쪽. PowerPoint가 없거나 창이 없으면 null(새로 띄우지 않는다).</summary>
    ActivePresentationInfo? GetActivePresentation();
    /// <summary>문서 창을 그 쪽으로 옮긴다(창이 있을 때만).</summary>
    void GoToSlide(string presentationPath, int slideIndex);
    /// <summary>창이 있는 문서 목록(BckPPT가 숨은 창으로 연 기준본 등은 빠진다). PowerPoint가 없으면 빈 목록.</summary>
    IReadOnlyList<OpenPresentationInfo> ListOpenWindows();
    /// <summary>그 문서의 창을 PowerPoint 앞으로 가져온다.</summary>
    void ActivatePresentation(string presentationPath);
    PresentationAnalysis AnalyzePresentation(
        string presentationPath,
        IProgress<SlideAnalysisProgress>? slideProgress = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// PowerPoint가 해석한 도형 계층을 슬라이드별 Markdown으로 내보낸다.
    /// 각 슬라이드는 PNG 미리보기와 함께 저장하여 이미지·차트 등 비텍스트 요소도 보존한다.
    /// </summary>
    MarkdownExportResult ExportMarkdown(string presentationPath, string outputPath,
        CancellationToken cancellationToken = default);

    Task<AutoFitResult> AutoFitAsync(string presentationPath, IProgress<string>? progress = null);
    Task FixShapeAsync(string presentationPath, int slideIndex, string shapeName,
        IProgress<string>? progress = null);

    Task NavigateToShapeAsync(string presentationPath, int slideIndex, string shapeName);
    void NavigateToSlide(string presentationPath, int slideIndex);
    void OpenPptxInPowerPoint(string filePath);

    /// <summary>문서를 창 없이 열어 두고 핸들을 해제하면 닫는다. 이미 열린 문서는 그대로 둔다.</summary>
    IDisposable KeepOpen(string presentationPath, bool readOnly);
    /// <summary>문서가 PowerPoint 창에 열려 있는지(사용자가 보고 있는 문서인지).</summary>
    bool IsOpenInWindow(string presentationPath);
    /// <summary>
    /// 열린 창에서 바로 작업한다. 핸들을 쥐고 있는 동안 이 문서는 저장하지 않고(저장은 사용자가 한다),
    /// 쓰기 전마다 사용자가 도형·글을 선택했는지 보고 그랬으면 <see cref="PptUserInterventionException"/>으로 멈춘다.
    /// 새로 만든 쪽은 화면에 띄운다.
    /// </summary>
    IDisposable WatchUserActivity(string presentationPath);
    /// <summary>창의 선택을 푼다. 바로 작업에서 요청을 시작할 때, 앞서 멈춘 뒤 남은 선택으로 다시 멈추지 않게 한다.</summary>
    void ClearSelection(string presentationPath);
    /// <summary>끊긴 세션이 창 없이 남긴 문서(작업 폴더의 사본, alsoPaths)를 저장하지 않고 닫는다. 닫은 파일 이름을 돌려준다.</summary>
    IReadOnlyList<string> CloseOrphanedPresentations(string workDirectory, IEnumerable<string> alsoPaths);

    /// <summary>슬라이드 번호와 제목만 빠르게 읽는다(도형 측정 없음).</summary>
    IReadOnlyList<(int Index, string Title)> ListSlides(string presentationPath);

    int  ReplaceText(string presentationPath, string find, string replace,
        IReadOnlyList<(int slideIndex, string shapeName)>? targets = null);
    int  NormalizeFonts(string presentationPath);
    void ReplaceImage(string presentationPath, int slideIndex, float targetLeft, float targetTop, string newImagePath);
    void SetShapeText(string presentationPath, int slideIndex, string shapeName, string newText);
    int DuplicateSlide(string presentationPath, int sourceSlideIndex, int insertAfterSlideIndex);
    int ImportSlide(string sourcePath, int sourceSlideIndex, string targetPath, int insertAfterSlideIndex);

    /// <summary>설정된 ReplaceRules를 프레젠테이션 전체에 일괄 적용. 변경된 항목 목록 반환.</summary>
    IReadOnlyList<(int SlideIndex, string ShapeName, string OldText, string NewText)>
        ApplyRules(string presentationPath, IReadOnlyList<ReplaceRule> rules);

    /// <summary>
    /// 문단 단위로 텍스트를 교체·삽입한다. markRed가 켜지면 손댄 문단만 빨간색으로 표시하고
    /// 원래 색을 스냅샷으로 돌려준다.
    /// </summary>
    ParagraphEditResult SetParagraphTexts(
        string presentationPath, IReadOnlyList<ParagraphEdit> edits, bool markRed);

    /// <summary>빨간색 표시를 스냅샷의 원래 글자색으로 되돌린다.</summary>
    void RestoreParagraphColors(
        string presentationPath, IReadOnlyList<ParagraphColorSnapshot> snapshots);

    /// <summary>단일 shape의 현재 텍스트·메트릭을 읽어 PptSlideInfo로 반환. 찾지 못하면 null.</summary>
    PptSlideInfo? GetShapeInfo(string presentationPath, int slideIndex, string shapeName);

    SlideLayout GetSlideLayout(string presentationPath, int slideIndex);

    /// <summary>슬라이드를 PNG 썸네일로 내보내고 임시 파일 경로를 반환. 실패 시 null.</summary>
    Task<string?> GetSlideThumbnailAsync(string presentationPath, int slideIndex);

    IReadOnlyList<RecentPptxFile> GetRecentFiles();

    /// <summary>텍스트 상자의 각 시각적 줄과 줄넘김 유형(자동/Enter/Shift+Enter)을 분석.</summary>
    IReadOnlyList<TextLineInfo> GetTextLineBreaks(string presentationPath, int slideIndex, string shapeName);

    // 아래 도형 API의 address는 ShapeAddress 규약("이름", "이름#2", "그룹/이름")을 따른다.
    // 그룹 이름도 받는다(그룹 통째로 지우기·옮기기). 같은 이름의 그룹 안 도형이 있으면 그 도형이 먼저다.
    // 위의 shapeName 인자들도 같은 규약을 받으며 그룹 안 도형까지 찾는다.

    /// <summary>그룹 안까지 펼친 도형 목록. 겹치는 이름은 주소에 "#n"이 붙는다.</summary>
    IReadOnlyList<ShapeEntry> ListShapes(string presentationPath, int slideIndex);
    void   RenameShape(string presentationPath, int slideIndex, string address, string newName);
    void   SetShapeGeometry(string presentationPath, int slideIndex, string address, ShapeGeometry geometry);
    void   DeleteShape(string presentationPath, int slideIndex, string address);
    string AddTextShape(string presentationPath, int slideIndex, TextShapeRequest request);
    /// <summary>여러 도형을 한 번에 추가한다(문서 열기·저장 한 번). 만든 이름을 순서대로 돌려준다.</summary>
    IReadOnlyList<string> AddTextShapes(string presentationPath, int slideIndex, IReadOnlyList<TextShapeRequest> requests);
    string AddLine(string presentationPath, int slideIndex, LineRequest request);
    TableLayout AddTable(string presentationPath, int slideIndex, TableRequest request);
    /// <summary>최상위 도형들을 그룹으로 묶고 그룹 이름을 돌려준다. 그룹 안 도형은 묶을 수 없다.</summary>
    string GroupShapes(string presentationPath, int slideIndex, IReadOnlyList<string> addresses, string? name);
    /// <summary>그룹을 풀고 풀린 도형 이름을 돌려준다.</summary>
    IReadOnlyList<string> UngroupShape(string presentationPath, int slideIndex, string address);
    /// <summary>keep에 든 도형(이름·주소)만 남기고 최상위 도형을 모두 지운다. 지운 수를 돌려준다.</summary>
    int ClearSlide(string presentationPath, int slideIndex, IReadOnlyCollection<string> keep);
    int    SetTextColor(string presentationPath, int slideIndex, string address, string find, string color);
    void   SetTextFormat(string presentationPath, int slideIndex, string address, TextFormatPatch patch);
    TextFormatInfo GetTextFormat(string presentationPath, int slideIndex, string address);
    int    ReplaceTextInShape(string presentationPath, int slideIndex, string address, string find, string replace);
}
