namespace BckPPT.Models;

public enum ParagraphEditKind
{
    /// <summary>해당 문단의 텍스트를 교체.</summary>
    Replace,
    /// <summary>해당 문단 뒤에 새 문단을 삽입.</summary>
    Insert,
    /// <summary>해당 문단을 통째로 삭제. 되돌리기에서 삽입을 취소할 때 쓴다.</summary>
    Delete
}

/// <summary>COM 계층에 전달하는 문단 단위 편집 지시.</summary>
public record ParagraphEdit(
    int               SlideIndex,
    string            ShapeName,
    int               ParagraphIndex,  // 1-based. Insert면 "이 문단 뒤"를 뜻한다.
    string            NewText,
    ParagraphEditKind Kind = ParagraphEditKind.Replace);

/// <summary>
/// 편집 결과. <see cref="Snapshots"/>는 빨간색 표시를 되돌리는 데 쓰고,
/// <see cref="Failures"/>는 조용히 넘어가지 않도록 사용자에게 보여준다.
/// </summary>
public record ParagraphEditResult(
    IReadOnlyList<ParagraphColorSnapshot> Snapshots,
    IReadOnlyList<string>                 Failures);
