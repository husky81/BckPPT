namespace BckPPT.Models;

/// <summary>PDF 내보내기 요청. Slides가 비면 전체 쪽.</summary>
public sealed record PdfExportRequest(
    string PresentationPath,
    string OutputPath,
    IReadOnlyList<int>? Slides = null,
    bool IncludeHidden = false,
    bool Overwrite = false,
    TimeSpan? Timeout = null);

/// <summary>PDF 내보내기 결과.</summary>
/// <param name="Pages">PDF에 든 쪽 수(숨긴 쪽 제외 반영).</param>
/// <param name="TotalSlides">문서 전체 쪽 수.</param>
/// <param name="HiddenSkipped">고른 쪽 중 숨긴 쪽이라 빠진 수.</param>
/// <param name="FromOpenDocument">PowerPoint에 이미 열려 있던 문서를 그대로 내보냈는지.</param>
/// <param name="OpenWindows">이미 열린 문서였을 때 그 창 수(0이면 창 없이 메모리에 남은 문서).</param>
/// <param name="HadUnsavedChanges">이미 열린 문서에 저장 안 된 고침이 있었는지(있으면 PDF에 그 고침이 들어갔다).</param>
/// <param name="StartedPowerPoint">PowerPoint가 떠 있지 않아 새로 띄웠는지.</param>
/// <param name="QuitPowerPoint">새로 띄운 PowerPoint를 끝에 닫았는지.</param>
public sealed record PdfExportResult(
    string OutputPath,
    int Pages,
    int TotalSlides,
    int HiddenSkipped,
    long Bytes,
    TimeSpan Elapsed,
    bool FromOpenDocument,
    int OpenWindows,
    bool HadUnsavedChanges,
    bool StartedPowerPoint,
    bool QuitPowerPoint,
    string? Note = null);
