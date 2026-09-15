namespace BckPPT.Models;

/// <summary>
/// 빨간색 표시를 걸기 전의 원래 글자색. "표시 해제"에서 복원에 사용한다.
/// 테마색이 걸린 문단은 RGB만 되돌리면 테마 연결이 끊기므로 함께 보관한다.
/// </summary>
public record ParagraphColorSnapshot(
    int    SlideIndex,
    string ShapeName,
    int    ParagraphIndex,
    int    Rgb,               // PowerPoint COM은 BGR 순서
    int    ObjectThemeColor); // msoNotThemeColor = 0
