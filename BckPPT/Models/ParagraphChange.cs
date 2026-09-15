using CommunityToolkit.Mvvm.ComponentModel;

namespace BckPPT.Models;

public enum ParagraphChangeKind
{
    /// <summary>기존 문단의 텍스트를 교체.</summary>
    Replace,
    /// <summary>기존 문단을 빈 문단으로.</summary>
    Clear,
    /// <summary>기준 문단 뒤에 새 문단을 삽입 (서식은 기준 문단에서 상속).</summary>
    Insert
}

/// <summary>미리보기 창의 한 행. 사용자가 항목별로 켜고 끌 수 있다.</summary>
public partial class ParagraphChange : ObservableObject
{
    public required string Id             { get; init; }
    public required int    SlideIndex     { get; init; }
    public required string ShapeName      { get; init; }
    public required int    ParagraphIndex { get; init; }
    public required string OldText        { get; init; }
    public required string NewText        { get; init; }

    public ParagraphChangeKind Kind { get; init; } = ParagraphChangeKind.Replace;

    [ObservableProperty]
    private bool _isSelected = true;

    /// <summary>넘침 예상 경고. 비어 있으면 문제 없음.</summary>
    public string OverflowWarning { get; set; } = "";

    public bool HasWarning => OverflowWarning.Length > 0;

    public string KindLabel => Kind switch
    {
        ParagraphChangeKind.Insert => "문단 추가",
        ParagraphChangeKind.Clear  => "문단 비움",
        _                          => "치환"
    };

    public string Location => $"{SlideIndex}쪽 · {ShapeName}";

    public string OldDisplay => OldText.Length == 0 ? ParagraphEntry.EmptyMarker : Show(OldText);
    public string NewDisplay => NewText.Length == 0 ? ParagraphEntry.EmptyMarker : Show(NewText);

    /// <summary>Shift+Enter(\v)를 눈에 보이게.</summary>
    private static string Show(string s) => s.Replace("\v", "\n");
}
