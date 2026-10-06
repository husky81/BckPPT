using BckPPT.Models;

namespace BckPPT.Agent;

/// <summary>
/// "이 상자의 문단을 이 목록으로 바꿔라"를 <see cref="ParagraphEdit"/> 목록으로 바꾼다.
/// 겹치는 문단은 교체해 서식(불릿·들여쓰기·글꼴)을 살리고, 남으면 뒤에서부터 지우고, 모자라면 마지막 문단 뒤에 넣는다.
/// </summary>
public static class ParagraphFitPlanner
{
    /// <summary>문단 안 줄바꿈(Shift+Enter)은 PowerPoint의 수직 탭(\v)이다. 모델이 \n으로 보내도 같게 본다.</summary>
    public static string NormalizeParagraph(string text) =>
        text.Replace("\r\n", "\v").Replace('\n', '\v').Replace('\r', '\v');

    public static IReadOnlyList<ParagraphEdit> Plan(
        int slideIndex, string address, int currentCount, IReadOnlyList<string> desired)
    {
        if (desired.Count == 0) throw new ArgumentException("문단이 하나 이상 있어야 합니다.", nameof(desired));
        currentCount = Math.Max(currentCount, 1);   // 빈 상자도 문단 하나는 있다.

        var edits = new List<ParagraphEdit>();
        int shared = Math.Min(currentCount, desired.Count);
        for (int i = 0; i < shared; i++)
            edits.Add(new ParagraphEdit(slideIndex, address, i + 1, NormalizeParagraph(desired[i])));

        // 삭제는 SetParagraphTexts가 큰 인덱스부터 처리하므로 순서와 무관하다.
        for (int p = desired.Count + 1; p <= currentCount; p++)
            edits.Add(new ParagraphEdit(slideIndex, address, p, "", ParagraphEditKind.Delete));

        // 삽입은 모두 "원래 마지막 문단 뒤"에 넣는다. 같은 인덱스끼리는 나열 순서대로 실행되어
        // 나중 것이 앞에 끼어들므로, 거꾸로 나열해야 최종 순서가 desired와 같다.
        for (int i = desired.Count - 1; i >= currentCount; i--)
            edits.Add(new ParagraphEdit(slideIndex, address, currentCount, NormalizeParagraph(desired[i]),
                ParagraphEditKind.Insert));
        return edits;
    }

    /// <summary>TextRange.Text 기준 문단 수. 문단 구분은 \r이고 빈 텍스트도 문단 하나다.</summary>
    public static int CountParagraphs(string fullText) =>
        fullText.Length == 0 ? 1 : fullText.TrimEnd('\r').Split('\r').Length;
}
