using BckPPT.Models;

namespace BckPPT.Services;

/// <summary>
/// 치환을 적용하기 <b>전에</b> 줄 넘침을 예측한다.
/// PowerPoint에 쓰고 나서야 넘침을 발견하면 되돌리는 비용이 크기 때문이다.
///
/// 새 글꼴 메트릭을 재지 않고, 이미 분석해 둔 값만 쓴다.
/// - <see cref="PptSlideInfo.MaxLineHw"/>: 그 상자에서 실제로 한 줄에 들어갔던 최대 반각 길이
/// - <see cref="PptSlideInfo.BoundHeight"/> / <see cref="PptSlideInfo.LineCount"/>: 줄 높이
/// </summary>
public static class OverflowEstimator
{
    /// <summary>
    /// 도형의 문단들이 <paramref name="newParagraphs"/>로 바뀌었을 때 넘침 여부를 판정한다.
    /// 판정할 수 없으면 빈 문자열을 돌려준다 (경고를 지어내지 않는다).
    /// </summary>
    public static string Estimate(PptSlideInfo shape, IReadOnlyList<string> newParagraphs)
    {
        if (shape.MaxLineHw <= 0 || shape.LineCount <= 0 ||
            shape.BoundHeight <= 0 || shape.BoxHeight <= 0)
            return "";

        float lineHeight = shape.BoundHeight / shape.LineCount;
        int   estLines   = newParagraphs.Sum(p => EstimateLines(p, shape.MaxLineHw));
        if (estLines <= 0) return "";

        float predicted = estLines * lineHeight;
        float tolerance = PptRules.Current.OverflowTolerance;
        if (predicted <= shape.BoxHeight + tolerance) return "";

        int curLines = shape.LineCount;
        return estLines > curLines
            ? $"{curLines}줄 → {estLines}줄 예상 (상자 초과)"
            : "상자 초과 예상";
    }

    /// <summary>한 문단이 몇 줄로 접힐지. 문단 안의 Shift+Enter(\v)는 줄을 강제로 나눈다.</summary>
    private static int EstimateLines(string paragraph, int maxLineHw)
    {
        int total = 0;
        foreach (var segment in paragraph.Split('\v'))
        {
            int hw = PptSlideInfo.CalcHalfWidth(segment);
            total += hw <= 0 ? 1 : (int)Math.Ceiling(hw / (double)maxLineHw);
        }
        return total;
    }
}
