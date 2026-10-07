namespace BckPPT.Models;

/// <summary>PowerPoint에서 지금 보고 있는 문서(활성 창). 상태줄이 따라간다.</summary>
public record ActivePresentationInfo(string Name, string Path, int SlideCount, int? CurrentSlide);

public record PptPresentation(string Name, string Path, int SlideCount)
{
    public string SlideCountText => $"{SlideCount}장";
    public override string ToString() => Name;
}
