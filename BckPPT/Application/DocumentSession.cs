using CommunityToolkit.Mvvm.ComponentModel;
using BckPPT.Models;

namespace BckPPT.Application;

public partial class DocumentSession : ObservableObject
{
    [ObservableProperty]
    private PptPresentation? _presentation;

    [ObservableProperty]
    private int? _selectedSlideIndex;

    public string TargetText => Presentation is null
        ? "대상 문서가 선택되지 않았습니다."
        : $"{Presentation.Name} · {(SelectedSlideIndex is null ? "슬라이드 미선택" : $"{SelectedSlideIndex}쪽")}";

    partial void OnPresentationChanged(PptPresentation? value) => OnPropertyChanged(nameof(TargetText));
    partial void OnSelectedSlideIndexChanged(int? value) => OnPropertyChanged(nameof(TargetText));
}
