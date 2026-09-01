using CommunityToolkit.Mvvm.ComponentModel;

namespace CellBox2.BckPPT.Models;

public partial class TextMatchItem : ObservableObject
{
    public int    SlideIndex  { get; init; }
    public string ShapeName   { get; init; } = "";
    public string ContextText { get; init; } = "";

    [ObservableProperty]
    private bool _isSelected = true;
}
