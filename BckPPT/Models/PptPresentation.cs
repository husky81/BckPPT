namespace CellBox2.BckPPT.Models;

public record PptPresentation(string Name, string Path, int SlideCount)
{
    public string SlideCountText => $"{SlideCount}장";
    public override string ToString() => Name;
}
