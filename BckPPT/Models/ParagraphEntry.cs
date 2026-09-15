namespace BckPPT.Models;

/// <summary>
/// 문단 단위 텍스트 항목. <see cref="Id"/>는 외부 에이전트와 주고받는 안정 키다.
/// 형식은 <c>S{슬라이드}-{도형순번}-{문단번호}</c> (예: <c>S1-12-3</c>).
/// 마크다운 표 구분자와 충돌하지 않도록 <c>|</c> 대신 하이픈을 쓴다.
/// </summary>
public record ParagraphEntry(
    string Id,
    int    SlideIndex,
    string ShapeName,       // 테이블 셀은 "Table 1 [R2C3]" 형식
    int    ShapeSeq,        // 창을 여는 시점에 확정되는 도형 순번
    int    ParagraphIndex,  // 1-based (PowerPoint COM 규약)
    string Text)
{
    public bool IsEmpty => string.IsNullOrWhiteSpace(Text);

    /// <summary>목록 표시용. 빈 문단은 플레이스홀더로 보여 ID를 잃지 않게 한다.</summary>
    public string Display => IsEmpty ? EmptyMarker : Text;

    public const string EmptyMarker = "(빈 문단)";

    public static string MakeId(int slideIndex, int shapeSeq, int paragraphIndex)
        => $"S{slideIndex}-{shapeSeq}-{paragraphIndex}";
}
