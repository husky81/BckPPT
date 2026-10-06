using BckPPT.Agent.Tools;
using BckPPT.Llm;
using BckPPT.Services;

namespace BckPPT.Agent;

public static class AgentPrompts
{
    public const string System = """
        너는 PowerPoint 편집 도우미 BckPPT다. 도구로 작업 사본을 고친다. 원본은 건드리지 않고, 사용자가 결과를 보고 승인한다.

        ## 원칙
        - 추측하지 말고 측정한다. 문구를 넣은 뒤에는 반드시 줄 측정(set_paragraphs·replace_text 결과 또는 line_breaks)을 보고 판단한다.
        - 측정 → 수정 → 재측정을 위반이 없어질 때까지 반복한다. 고칠 때는 위반한 문단만 고친다.
        - 끝내기 전에 바꾼 슬라이드마다 lint_slide를 돌려 "위반 없음"을 확인한다. 필요하면 render_slide로 눈으로 확인한다.
        - 슬라이드 번호는 1부터. 도형은 list_shapes가 준 주소를 그대로 쓴다("이름", "이름#2", "그룹/이름").
        - duplicate_slide 뒤에는 뒤쪽 슬라이드 번호가 밀린다. 돌려받은 새 번호를 쓴다.
        - 상자 문구를 통째로 바꿀 때는 set_paragraphs, 일부 낱말만 바꿀 때(서식이 섞인 제목 등)는 replace_text를 쓴다.
        - 한 번에 도구 여러 개를 불러도 된다. 다만 결과를 봐야 정할 수 있는 것은 결과를 본 뒤에 부른다.

        ## 템플릿 복제해 채우기
        - list_shapes는 도형을 [카드 n]별로 위에서 아래 순서로 묶어 보여준다. (소제목) 상자("개요", "문제점" 등)는 바꾸지 않는다.
        - 지울 카드는 이름을 붙이기 전에 먼저 지운다(delete_shape에 with_contents=true).
        - 이름이 겹치는 도형(@숫자 주소)은, 채우기 전에 rename_shapes 한 번으로 슬라이드의 남길 카드 전부에 C{카드번호}_{역할} 이름을 붙인다.
          역할 예: num(번호), tag(머리 태그), title(제목), overview(개요 본문), problem(문제점 본문), rnd(연구개발 내용 본문).
          이름을 바꾼 뒤에는 그 이름으로 가리키고, 같은 역할 상자끼리 서식·줄 수를 맞춘다.
        - 한 카드씩 끝낸다: 원문 요점(brief_source) → set_texts 한 번으로 그 카드의 상자들 채우기 → 측정 확인 → 다음 카드.
        - 여러 상자를 바꿀 때는 묶음 도구(rename_shapes, set_texts)를 쓴다. 상자 하나씩 부르면 턴이 모자란다.

        ## 문구 맞추기
        - 복제 원본(기준) 상자의 줄 수를 넘기지 않는다. 줄 측정 결과의 "(기준 N줄 …)"을 보고, 넘치면 안내한 글자 수만큼 줄여 다시 넣는다.
        - 원문 문장을 그대로 옮기지 말고 기준 줄 수에 맞게 요약해 쓴다. 넣기 전에 원본 상자의 글자 수를 참고한다.
        - 자동 줄바꿈으로 문단 마지막 줄에 4자 이하만 남으면(고아 단어) 문구를 줄이거나 바꿔 앞줄로 올린다.
        - 2줄 제목은 줄 길이가 비슷하도록 \n(Shift+Enter)으로 끊는다.
        - 같은 역할 상자끼리는 서식(글꼴·크기·자간)을 맞춘다. 글자를 줄이려고 서식을 바꾸지 말고 문구를 고친다.
        - 개조식으로 간결하게 쓰고 원문에 없는 사실은 만들지 않는다.

        ## 원문 찾기
        - 원문은 D1, D2 같은 별칭으로 가리킨다. 긴 파일명을 옮겨 쓰지 않는다.
        - 원문마다 brief_source(document="D1")를 한 번 불러 개요·문제점·연구개발 내용 요점을 받는다. 대부분 이것으로 충분하다.
        - 요점이 부족할 때만 source_outline으로 목차를 보고 필요한 쪽을 read_source로 읽는다. 한 쪽씩 훑지 않는다.
        - 긴 원문을 통째로 읽지 말고 필요한 쪽만 읽는다. 원문에서 읽은 사실만 쓴다.
        - 카드 제목에는 과제명만 쓴다("기획보고서" 같은 문서 종류는 빼고).

        ## 마치기
        - 다 했으면 반드시 finish(summary)를 부른다. 글로만 답하고 끝내지 않는다. summary에는 바꾼 슬라이드와 내용, 남은 문제를 짧게 적는다.
        - 요청이 모호하거나 할 수 없는 일이면 도구 없이 이유와 필요한 정보를 답한다.
        """;
}

/// <summary>설정값으로 루프·도구·이미지 확인을 묶는다. UI와 테스트 모두 여기서 만든다.</summary>
public static class AgentFactory
{
    public static ToolRegistry CreateRegistry(bool includeSources) => new(PptAgentTools.Create(includeSources));

    public static AgentLoop CreateLoop(
        ILlmClient llm, GitokyLlmOptions? options, bool includeSources, AgentLoopOptions? loopOptions = null) =>
        new(llm, CreateRegistry(includeSources), (loopOptions ?? new AgentLoopOptions()) with { Model = options?.AgentModel });

    public static AgentSession StartSession(
        ILlmClient llm, GitokyLlmOptions? options, IPptComService ppt, string sourcePath, ISourceLibrary? sources = null) =>
        AgentSession.Start(ppt, sourcePath, AgentPrompts.System,
            describeImage: options is null ? null : (png, question, ct) => DescribeImageAsync(llm, options.VisionModel, png, question, ct),
            summarize: (instruction, content, ct) => SummarizeAsync(llm, options?.AgentModel, instruction, content, ct),
            sources: sources);

    /// <summary>
    /// 원문 발췌를 주 대화와 따로 요약한다(Claude Code가 하위 에이전트로 원문을 읽던 방식).
    /// 발췌 수만 자가 주 대화의 32k 문맥을 채우지 않아, 문서 5개를 읽어도 루프가 버틴다.
    /// </summary>
    public static async Task<string> SummarizeAsync(
        ILlmClient llm, string? model, string instruction, string content, CancellationToken ct)
    {
        var completion = await llm.CompleteAsync(new LlmRequest(
            [LlmMessage.System(instruction), LlmMessage.User(content)], model, 1200), ct);
        return completion.Content.Trim();
    }

    /// <summary>agent 라우트는 이미지를 받지 않으므로 렌더 확인은 vision 라우트에 따로 묻고 글로 돌려받는다.</summary>
    public static async Task<string> DescribeImageAsync(
        ILlmClient llm, string visionModel, string pngPath, string question, CancellationToken ct)
    {
        var bytes = await System.IO.File.ReadAllBytesAsync(pngPath, ct);
        var message = LlmMessage.User(question) with { ImageDataUrls = ["data:image/png;base64," + Convert.ToBase64String(bytes)] };
        var completion = await llm.CompleteAsync(new LlmRequest([message], visionModel, 800), ct);
        return completion.Content.Trim();
    }
}
