namespace BckPPT.Models;

/// <summary>PowerPoint에서 지금 보고 있는 문서(활성 창). 상태줄이 따라간다.</summary>
public record ActivePresentationInfo(string Name, string Path, int SlideCount, int? CurrentSlide);

/// <summary>PowerPoint 창에 열린 문서 하나. <see cref="IsSaved"/>가 false면 저장된 적 없는 새 문서(경로 없음)다.</summary>
public record OpenPresentationInfo(string Name, string Path, int SlideCount, bool IsActive, bool IsSaved);

public record PptPresentation(string Name, string Path, int SlideCount)
{
    public string SlideCountText => $"{SlideCount}장";
    public override string ToString() => Name;
}
