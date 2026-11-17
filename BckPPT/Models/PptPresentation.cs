namespace BckPPT.Models;

/// <summary>PowerPoint에서 지금 보고 있는 문서(활성 창). 상태줄이 따라간다.</summary>
public record ActivePresentationInfo(string Name, string Path, int SlideCount, int? CurrentSlide);

/// <summary>PowerPoint 창에 열린 문서 하나. <see cref="IsSaved"/>가 false면 저장된 적 없는 새 문서(경로 없음)다.</summary>
public record OpenPresentationInfo(string Name, string Path, int SlideCount, bool IsActive, bool IsSaved, bool IsReadOnly = false);

public record PptPresentation(string Name, string Path, int SlideCount)
{
    public string SlideCountText => $"{SlideCount}장";
    public override string ToString() => Name;
}

/// <summary>「제한된 보기」(Protected View) 창. 이 창의 문서는 Presentations에 없어 다른 도구가 읽지 못한다.</summary>
public record ProtectedViewInfo(string Name, string Path);

public enum OpenedState
{
    /// <summary>편집할 수 있게 열렸다.</summary>
    Editable,
    /// <summary>열렸지만 읽기 전용이다(다른 곳에서 쓰는 중 등).</summary>
    ReadOnly,
    /// <summary>「제한된 보기」로 열렸다. 다른 도구가 이 문서를 찾지 못한다.</summary>
    ProtectedView,
    /// <summary>기다리는 동안 문서가 올라오지 않았다.</summary>
    NotYet,
}

/// <summary>open_presentation 결과. HadDownloadMark: 인터넷에서 받은 파일 표시(Zone.Identifier)가 있었는지.</summary>
public record OpenPresentationResult(OpenedState State, bool HadDownloadMark, bool Unblocked, TimeSpan Waited, int SlideCount = 0);
