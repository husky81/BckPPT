using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace BckPPT.Models;

/// <summary>TreeView용 shape 노드. 트리 구조(부모-자식) + 드래그 이동 지원.</summary>
public partial class PptNode : ObservableObject
{
    public PptSlideInfo Shape { get; }
    public ObservableCollection<PptNode> Children { get; } = [];

    [ObservableProperty]
    private PptNode? _parent;

    [ObservableProperty]
    private bool _isExpanded = true;

    [ObservableProperty]
    private bool _isSelected;

    [ObservableProperty]
    private bool _isVisible = true;

    /// <summary>안정 키: {slide}:{shapeName}</summary>
    public string NodeId => $"{Shape.SlideIndex}:{Shape.ShapeName}";

    /// <summary>역할 약어 (UI 바인딩용, 변경 시 알림 발생).</summary>
    public string RoleAbbr => Shape.Role.ToAbbr();

    /// <summary>접기/펼치기 삼각형 글리프. 자식이 없으면 빈 문자열.</summary>
    public string ExpanderGlyph => Children.Count > 0
        ? (IsExpanded ? "\u25BE" : "\u25B8")  // ▾ / ▸
        : "";

    partial void OnIsExpandedChanged(bool value)
        => OnPropertyChanged(nameof(ExpanderGlyph));

    /// <summary>Shape의 Role을 변경하고 UI에 알림.</summary>
    public void SetRole(ShapeRole role)
    {
        Shape.Role = role;
        Shape.IsUserLabeled = true;
        OnPropertyChanged(nameof(RoleAbbr));
    }

    public PptNode(PptSlideInfo shape, PptNode? parent = null)
    {
        Shape = shape;
        _parent = parent;
        Children.CollectionChanged += (_, _) => OnPropertyChanged(nameof(ExpanderGlyph));
    }

    /// <summary>이 노드를 현재 부모에서 제거.</summary>
    public void Detach()
    {
        if (Parent is not null)
        {
            Parent.Children.Remove(this);
            Parent = null;
        }
    }

    /// <summary>target이 이 노드의 자손인지 검사 (자기 자신 아래로 드롭 방지).</summary>
    public bool IsDescendant(PptNode target)
    {
        var cur = target;
        while (cur is not null)
        {
            if (cur == this) return true;
            cur = cur.Parent;
        }
        return false;
    }
}
