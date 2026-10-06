using BckPPT.Agent.Tools;
using BckPPT.Llm;
using BckPPT.PowerPoint;

namespace BckPPT.Agent;

public sealed record AgentLoopOptions
{
    public string? Model { get; init; }
    public int MaxTurns { get; init; } = 30;
    public int MaxTokensPerTurn { get; init; } = 2048;

    /// <summary>도구 결과 한 건을 모델에 넘기는 최대 글자 수.</summary>
    public int MaxToolResultChars { get; init; } = 8000;

    /// <summary>
    /// 대화 전체 글자 수 상한. gitoky agent 라우트의 문맥은 32k 토큰이고 한글은 대략 1~1.5자/토큰이라
    /// 응답 여유를 남겨 이 정도로 잡는다. 넘으면 오래된 도구 결과부터 줄인다.
    /// </summary>
    public int ContextCharBudget { get; init; } = 36000;

    /// <summary>모델 문맥 한도(토큰). gitoky agent 라우트는 32,768.</summary>
    public int ContextWindowTokens { get; init; } = 32768;

    /// <summary>글자당 토큰 추정이 빗나갈 때를 위한 여유.</summary>
    public int ContextSafetyTokens { get; init; } = 2500;

    /// <summary>세션 누적 토큰 상한(입력+출력). 0이면 무제한.</summary>
    public int MaxSessionTokens { get; init; } = 400_000;

    /// <summary>finish 전에 건드린 슬라이드를 Lint로 검사해 위반이 있으면 되돌려 보내는 횟수.</summary>
    public int LintRejectionsBeforeFinish { get; init; } = 2;

    /// <summary>슬라이드를 바꾼 뒤 도구도 finish도 없이 답하면 이어가라고 되돌려 보내는 횟수.</summary>
    public int NudgesWithoutFinish { get; init; } = 2;
}

public enum AgentStopReason { Finished, AnsweredWithoutTools, TurnLimit, TokenLimit }

public sealed record AgentRunResult(
    AgentStopReason Reason,
    string Summary,
    int Turns,
    LlmUsage Usage,
    IReadOnlyList<LintIssue> RemainingIssues);

public enum AgentEventKind { Thinking, ToolStarted, ToolFinished, Message, Compacted }

public sealed record AgentEvent(AgentEventKind Kind, string Text, string? ToolName = null, bool IsError = false);

/// <summary>
/// 도구 호출 루프. 모델이 도구를 부르면 실행해 결과를 대화에 붙이고 다시 묻는다.
/// finish를 부르거나 도구 없이 답하면 끝난다. 상태(대화·작업 사본)는 <see cref="AgentSession"/>에 있어
/// 사용자 후속 요청은 같은 세션에서 <see cref="RunAsync"/>를 다시 부르면 이어진다.
/// </summary>
public sealed class AgentLoop(ILlmClient llm, ToolRegistry tools, AgentLoopOptions options)
{
    public const string ElidedPrefix = "[이전 도구 결과 생략";

    public const string ContinueNudge =
        "(자동 안내) 슬라이드를 바꾼 상태입니다. 남은 단계가 있으면 말하지 말고 지금 도구를 호출하세요. " +
        "다 끝났으면 finish(summary)를 호출하세요. 사용자에게 물어볼 것이 있으면 finish의 summary에 적으세요.";

    public async Task<AgentRunResult> RunAsync(
        AgentSession session, string userMessage,
        IProgress<AgentEvent>? progress = null, CancellationToken cancellationToken = default)
    {
        session.Append(LlmMessage.User(userMessage));
        var definitions = tools.Definitions();
        int lintRejections = 0;
        IReadOnlyList<LintIssue>? finishIssues = null;
        int nudges = 0;
        var runUsage = LlmUsage.Zero;
        session.Context.FinishSummary = null;

        for (int turn = 1; turn <= options.MaxTurns; turn++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            int budget = CharBudget(session);
            if (Compact(session.Messages, budget) is > 0 and var elided)
                progress?.Report(new(AgentEventKind.Compacted, $"오래된 도구 결과 {elided}건을 줄였습니다."));

            progress?.Report(new(AgentEventKind.Thinking, $"{turn}턴"));
            int sentChars = session.Messages.Sum(m => m.Content.Length);
            LlmCompletion completion;
            try
            {
                completion = await Ask(session, definitions, cancellationToken);
            }
            catch (LlmException ex) when (ex.Message.Contains("context window", StringComparison.OrdinalIgnoreCase))
            {
                // 추정이 빗나가 문맥을 넘으면 절반으로 줄이고 최근 결과만 남겨 한 번 더 묻는다.
                Compact(session.Messages, budget / 2, keepRecent: 2);
                progress?.Report(new(AgentEventKind.Compacted, "문맥 한도를 넘어 이전 결과를 더 줄이고 다시 묻습니다."));
                sentChars = session.Messages.Sum(m => m.Content.Length);
                completion = await Ask(session, definitions, cancellationToken);
            }
            if (completion.Usage.PromptTokens > 0) session.ObserveContext(sentChars, completion.Usage.PromptTokens);
            runUsage += completion.Usage;
            session.AddUsage(completion.Usage);
            session.Append(LlmMessage.Assistant(completion.Content, completion.ToolCalls));
            if (!string.IsNullOrWhiteSpace(completion.Content))
                progress?.Report(new(AgentEventKind.Message, completion.Content.Trim()));

            // qwen3는 "이제 lint_slide로 확인하겠습니다"라고 말만 하고 멈추곤 한다(실측). 슬라이드를 바꾼 뒤라면
            // 한 번은 이어가게 한다. 그래도 도구 없이 답하면 질문일 수 있으므로 끝내고, 바꾼 슬라이드의 판정만 보여준다.
            if (completion.ToolCalls.Count == 0 && nudges < options.NudgesWithoutFinish &&
                session.Context.Origins.Touched.Count > 0)
            {
                nudges++;
                // 남은 위반을 함께 보여줘야 무엇을 고칠지 안다. 실측에서 위반 15건을 남긴 채 "다음 단계"만 말하고 멈췄다.
                var pending = await LintTouchedAsync(session.Context, cancellationToken);
                session.Append(LlmMessage.User(pending.Count == 0
                    ? ContinueNudge
                    : ContinueNudge + "\n" + Truncate(LayoutLint.Describe(pending), 3000)));
                continue;
            }
            if (completion.ToolCalls.Count == 0)
                return Result(AgentStopReason.AnsweredWithoutTools, completion.Content.Trim(), turn, runUsage,
                    await LintTouchedAsync(session.Context, cancellationToken));

            foreach (var call in completion.ToolCalls)
            {
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new(AgentEventKind.ToolStarted, call.ArgumentsJson, call.Name));
                var outcome = await tools.ExecuteAsync(session.Context, call, cancellationToken);
                var content = outcome.Content;
                if (call.Name == "finish") finishIssues = null;   // 앞서 되돌려 보낸 판정은 이제 낡았다.

                if (call.Name == "finish" && !outcome.IsError && lintRejections < options.LintRejectionsBeforeFinish &&
                    (finishIssues = await LintTouchedAsync(session.Context, cancellationToken)) is { Count: > 0 } issues)
                {
                    // 위반이 남은 채 끝내려 하면 한두 번은 돌려보낸다. 그래도 남으면 사용자에게 그대로 보여준다.
                    lintRejections++;
                    session.Context.FinishSummary = null;
                    content = "아직 끝낼 수 없습니다. 고친 뒤 다시 finish 하세요.\n" + LayoutLint.Describe(issues);
                    outcome = outcome with { IsError = true };
                }

                session.Append(LlmMessage.Tool(call.Id, Truncate(content, options.MaxToolResultChars)));
                progress?.Report(new(AgentEventKind.ToolFinished, content, call.Name, outcome.IsError));
            }

            if (session.Context.FinishSummary is { } summary)
                return Result(AgentStopReason.Finished, summary, turn, runUsage,
                    finishIssues ?? await LintTouchedAsync(session.Context, cancellationToken));

            if (options.MaxSessionTokens > 0 && session.Usage.TotalTokens >= options.MaxSessionTokens)
                return Result(AgentStopReason.TokenLimit, "세션 토큰 상한에 도달해 멈췄습니다.", turn, runUsage,
                    await LintTouchedAsync(session.Context, cancellationToken));
        }

        return Result(AgentStopReason.TurnLimit, $"{options.MaxTurns}턴 상한에 도달해 멈췄습니다.", options.MaxTurns, runUsage,
            await LintTouchedAsync(session.Context, cancellationToken));

        AgentRunResult Result(AgentStopReason reason, string text, int turns, LlmUsage usage, IReadOnlyList<LintIssue> issues)
        {
            session.Log.Write("result", new { reason = reason.ToString(), summary = text, turns, issues = issues.Select(i => i.Message) });
            return new AgentRunResult(reason, text, turns, usage, issues);
        }
    }

    private Task<LlmCompletion> Ask(AgentSession session, IReadOnlyList<LlmToolDefinition> definitions, CancellationToken ct) =>
        llm.CompleteAsync(new LlmRequest(session.Messages, options.Model, options.MaxTokensPerTurn)
        {
            Tools = definitions,
            ToolChoice = "auto"
        }, ct);

    /// <summary>
    /// 이번 요청에 쓸 대화 글자 예산. 직전 응답의 실제 prompt_tokens로 글자당 토큰을 재서
    /// (도구 정의 등 고정 비용이 들어가 있어 보수적이다) 문맥 한도 − 응답 몫 − 여유 안에 들게 한다.
    /// 한글은 글자당 토큰이 1에 가까워 고정 글자 예산만으로는 32k를 넘었다(2026-10-06 실측).
    /// </summary>
    private int CharBudget(AgentSession session)
    {
        int targetTokens = options.ContextWindowTokens - options.MaxTokensPerTurn - options.ContextSafetyTokens;
        if (session.EstimateCharBudget(targetTokens) is not int estimate) return options.ContextCharBudget;
        return Math.Min(options.ContextCharBudget, Math.Max(4000, estimate));
    }

    private static async Task<IReadOnlyList<LintIssue>> LintTouchedAsync(AgentToolContext context, CancellationToken ct)
    {
        var issues = new List<LintIssue>();
        foreach (var slide in context.Origins.Touched.OrderBy(s => s))
        {
            try { issues.AddRange(await PptAgentTools.LintAsync(context, slide, null, ct)); }
            catch (OperationCanceledException) { throw; }
            catch { /* 측정 실패는 판정 불가로 보고 루프를 막지 않는다. */ }
        }
        return issues;
    }

    public static string Truncate(string text, int max) =>
        text.Length <= max ? text : text[..max] + $"\n…(결과 {text.Length - max}자 생략. 범위를 좁혀 다시 요청하세요)";

    /// <summary>
    /// 대화가 예산을 넘으면 오래된 도구 결과를 한 줄로 줄인다. 시스템·사용자 메시지와 최근 결과는 남긴다.
    /// 도구 메시지 자체는 지우지 않는다(tool_call_id 짝이 깨지면 API가 거부한다). 줄인 건수를 돌려준다.
    /// </summary>
    public static int Compact(List<LlmMessage> messages, int budget, int keepRecent = 6)
    {
        int total = messages.Sum(m => m.Content.Length);
        int elided = 0;
        for (int i = 0; i < messages.Count - keepRecent && total > budget; i++)
        {
            var m = messages[i];
            if (m.Role != "tool" || m.Content.StartsWith(ElidedPrefix, StringComparison.Ordinal)) continue;
            string head = m.Content[..Math.Min(80, m.Content.Length)].Replace('\n', ' ');
            string replacement = $"{ElidedPrefix}: {head}…]";
            total -= m.Content.Length - replacement.Length;
            messages[i] = m with { Content = replacement };
            elided++;
        }
        return elided;
    }
}
