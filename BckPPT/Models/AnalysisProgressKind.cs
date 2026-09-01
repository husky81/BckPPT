namespace CellBox2.BckPPT.Models;

public enum AnalysisProgressKind
{
    SlideStart,      // 슬라이드 진입 즉시 — skeleton 추가용
    ShapeBatch,      // 33ms 또는 20개마다 batch flush
    SlideCompleted,  // 슬라이드 shape 루프 완료 — skeleton 확정
}
