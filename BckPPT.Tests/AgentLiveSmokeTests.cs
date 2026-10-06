using System.IO;
using BckPPT.Agent;
using BckPPT.Application;
using BckPPT.Infrastructure;
using BckPPT.Llm;
using BckPPT.PowerPoint;
using BckPPT.Rag;
using BckPPT.Services;
using Xunit.Abstractions;

namespace BckPPT.Tests;

/// <summary>
/// 실제 PowerPoint + gitoky로 채팅창과 같은 경로(AgentWorkflow)를 끝까지 돌린다.
/// BCKPPT_LIVE_PPTX(대상 파일)와 GITOKY_LLM_API_KEY가 있을 때만 실행한다.
/// BCKPPT_LIVE_SOURCES(원문, '|'로 구분)와 BCKPPT_LIVE_REQUEST(요청)로 바꿀 수 있다.
/// 작업 사본은 버리고 대화 기록 경로를 출력에 남긴다.
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
        using var artifacts = new TemporaryArtifactStore();
        var dbDir = Path.Combine(Path.GetTempPath(), "bckppt-live-" + Guid.NewGuid().ToString("N")[..8]);
        using var workflow = new AgentWorkflow(llm, options, ppt, new RenderValidator(ppt, artifacts),
            new SourceLibrary(Path.Combine(dbDir, "sources.db")));

        var sources = Environment.GetEnvironmentVariable("BCKPPT_LIVE_SOURCES")?.Split('|', StringSplitOptions.RemoveEmptyEntries) ?? [];
        if (sources.Length > 0)
        {
            var (files, pages) = await workflow.AttachAsync(sources);
            output.WriteLine($"원문 {files}개 {pages}쪽 첨부");
        }
        var request = Environment.GetEnvironmentVariable("BCKPPT_LIVE_REQUEST")
            ?? "1쪽을 복제해 바로 뒤에 넣고, 새 쪽의 두 카드를 각각 '배터리 열폭주 조기 감지 센서 기술'과 " +
               "'수소 연료전지 스택 내구성 향상 기술' 과제로 바꿔 제목·개요·연구개발 내용을 채워줘. 줄 수는 원본을 넘기지 마.";

        var progress = new SyncProgress(e => output.WriteLine($"[{e.Kind}] {e.ToolName} {Clip(e.Text)}"));
        try
        {
            var result = await workflow.RunAsync(pptx, request, progress, CancellationToken.None);
            var run = result.Run;
            output.WriteLine($"== {run.Reason} turns={run.Turns} tokens={run.Usage.TotalTokens}");
            output.WriteLine(run.Summary);
            foreach (var issue in run.RemainingIssues) output.WriteLine("남은 위반: " + issue.Message);
            output.WriteLine($"썸네일 전 {result.BeforeThumbnails.Count} / 후 {result.AfterThumbnails.Count}");
            foreach (var change in result.Changes) output.WriteLine("변경: " + change);
            output.WriteLine("log: " + workflow.Session?.Log.FilePath);
            Assert.NotEqual(AgentStopReason.TokenLimit, run.Reason);
        }
        finally
        {
            workflow.Discard();
            try { Directory.Delete(dbDir, true); } catch { }
        }
    }

    private static string Clip(string text) => text.Length <= 600 ? text : text[..600] + "…";

    private sealed class SyncProgress(Action<AgentEvent> report) : IProgress<AgentEvent>
    {
        public void Report(AgentEvent value) => report(value);
    }
}
