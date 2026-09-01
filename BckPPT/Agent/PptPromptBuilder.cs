using BckPPT.Application;
using BckPPT.Llm;

namespace BckPPT.Agent;

public static class PptPromptBuilder
{
    private const string SystemPrompt = """
        당신은 BckPPT의 PowerPoint 작업 계획 에이전트다.
        원본 파일을 직접 수정하거나 C#/COM 코드를 만들지 않는다.
        사용자의 요청과 제공된 최소 문맥만 사용해 안전한 작업계획을 작성한다.
        응답은 설명 없이 아래 JSON 형식 하나만 반환한다.
        {
          "summary": "변경 요약",
          "requiresApproval": true,
          "commands": [
            { "name": "analyze_presentation", "slideIndex": 1, "reason": "필요한 이유",
              "sourceSlideIndex": null, "insertAfterSlideIndex": null,
              "shapeName": null, "newText": null, "expectedText": null }
          ],
          "warnings": []
        }
        허용 명령: analyze_presentation, search_template_library, duplicate_slide,
        replace_shape_text, insert_slide, validate_layout, save_copy.
        새 슬라이드가 필요한 요청에는 search_template_library를 duplicate_slide 또는 insert_slide보다 먼저 넣는다.
        duplicate_slide에는 sourceSlideIndex와 insertAfterSlideIndex를 넣는다.
        replace_shape_text에는 slideIndex, shapeName, newText를 넣고 현재 텍스트를 알 때만 expectedText를 넣는다.
        문서나 슬라이드가 불명확하면 변경 명령 대신 analyze_presentation을 계획한다.
        """;

    public static LlmRequest BuildPlanRequest(string instruction, DocumentSession session, string? documentContext = null)
    {
        var target = session.Presentation is null
            ? "선택된 문서 없음"
            : $"파일명={session.Presentation.Name}; 슬라이드수={session.Presentation.SlideCount}; 선택슬라이드={session.SelectedSlideIndex?.ToString() ?? "없음"}";
        return new LlmRequest([
            new LlmMessage("system", SystemPrompt),
            new LlmMessage("user", $"대상: {target}\n\n문서 문맥:\n{documentContext ?? "분석 문맥 없음"}\n\n요청: {instruction}")
        ]);
    }
}
