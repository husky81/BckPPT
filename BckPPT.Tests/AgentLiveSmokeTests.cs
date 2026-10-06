using System.IO;
using BckPPT.Agent;
using BckPPT.Llm;
using BckPPT.Services;
using Xunit.Abstractions;

namespace BckPPT.Tests;

/// <summary>
/// 실제 PowerPoint + gitoky로 루프 전체를 돌린다. BCKPPT_LIVE_PPTX(대상 파일)와 GITOKY_LLM_API_KEY가
/// 있을 때만 실행하고, BCKPPT_LIVE_REQUEST로 요청을 바꿀 수 있다. 결과 사본은 지우고 대화 기록은 출력으로 남긴다.
/// </summary>
public class AgentLiveSmokeTests(ITestOutputHelper output)
{
    [Fact]
    public async Task 실서버_카드_채우기_스모크()
    {
        var pptx = Environment.GetEnvironmentVariable("BCKPPT_LIVE_PPTX");
        if (string.IsNullOrEmpty(pptx) || Environment.GetEnvironmentVariable("GITOKY_LLM_API_KEY") is not { Length: > 0 }) return;

        var options = GitokyLlmOptions.FromEnvironment(AppContext.BaseDirectory);
        using var llm = new GitokyLlmClient(options);
        var ppt = new PptComService();
        using var session = AgentFactory.StartSession(llm, options, ppt, pptx);
        var loop = AgentFactory.CreateLoop(llm, options, includeSources: false);
        var request = Environment.GetEnvironmentVariable("BCKPPT_LIVE_REQUEST")
            ?? "1쪽을 복제해 바로 뒤에 넣고, 새 쪽의 두 카드를 각각 '배터리 열폭주 조기 감지 센서 기술'과 " +
               "'수소 연료전지 스택 내구성 향상 기술' 과제로 바꿔 제목·개요·연구개발 내용을 채워줘. 줄 수는 원본을 넘기지 마.";

        var progress = new SyncProgress(e => output.WriteLine($"[{e.Kind}] {e.ToolName} {Clip(e.Text)}"));
        try
        {
            var result = await loop.RunAsync(session, request, progress);
            output.WriteLine($"== {result.Reason} turns={result.Turns} tokens={result.Usage.TotalTokens}");
            output.WriteLine(result.Summary);
            foreach (var issue in result.RemainingIssues) output.WriteLine("남은 위반: " + issue.Message);
            output.WriteLine("log: " + session.Log.FilePath);
        }
        finally
        {
            session.Discard();
        }
    }

    private static string Clip(string text) => text.Length <= 600 ? text : text[..600] + "…";

    private sealed class SyncProgress(Action<AgentEvent> report) : IProgress<AgentEvent>
    {
        public void Report(AgentEvent value) => report(value);
    }
}
