namespace BckPPT.Models;

/// <summary>Markdown 내보내기 결과. 경고는 렌더 이미지에서만 보존된 요소를 뜻한다.</summary>
public record MarkdownExportResult(string OutputPath, int SlideCount, int WarningCount);
