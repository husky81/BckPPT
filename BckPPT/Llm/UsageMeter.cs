namespace BckPPT.Llm;

/// <summary>
/// 모델(라우트)별 호출 수와 토큰. 루프·요약·문안 작성·길이 맞춤을 모두 센다.
/// 비싼 모델은 한두 번만 쓰고 나머지는 qwen에 맡기는 구조라 모델별로 봐야 비용이 보인다.
/// </summary>
public sealed class UsageMeter
{
    private readonly Dictionary<string, (int Calls, LlmUsage Usage)> _byModel = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    public void Add(string model, LlmUsage usage)
    {
        lock (_lock)
        {
            var (calls, total) = _byModel.GetValueOrDefault(model, (0, LlmUsage.Zero));
            _byModel[model] = (calls + 1, total + usage);
        }
    }

    public IReadOnlyDictionary<string, (int Calls, LlmUsage Usage)> Snapshot()
    {
        lock (_lock) return new Dictionary<string, (int, LlmUsage)>(_byModel);
    }

    public LlmUsage Total
    {
        get { lock (_lock) return _byModel.Values.Aggregate(LlmUsage.Zero, (a, v) => a + v.Usage); }
    }

    /// <summary>두 스냅숏의 차이(이번 요청에서 쓴 양).</summary>
    public static IReadOnlyDictionary<string, (int Calls, LlmUsage Usage)> Delta(
        IReadOnlyDictionary<string, (int Calls, LlmUsage Usage)> before,
        IReadOnlyDictionary<string, (int Calls, LlmUsage Usage)> after) =>
        after.Select(kv =>
            {
                var (calls, usage) = before.GetValueOrDefault(kv.Key, (0, LlmUsage.Zero));
                return (kv.Key, Value: (kv.Value.Calls - calls,
                    new LlmUsage(kv.Value.Usage.PromptTokens - usage.PromptTokens, kv.Value.Usage.CompletionTokens - usage.CompletionTokens)));
            })
            .Where(x => x.Value.Item1 > 0)
            .ToDictionary(x => x.Key, x => x.Value);

    public static string Describe(IReadOnlyDictionary<string, (int Calls, LlmUsage Usage)> usage) =>
        usage.Count == 0 ? "LLM 호출 없음"
            : string.Join(", ", usage.OrderByDescending(kv => kv.Value.Usage.TotalTokens)
                .Select(kv => $"{kv.Key} {kv.Value.Calls}회 {kv.Value.Usage.TotalTokens:N0}토큰(입력 {kv.Value.Usage.PromptTokens:N0}/출력 {kv.Value.Usage.CompletionTokens:N0})"));
}

/// <summary>모든 호출의 사용량을 <see cref="UsageMeter"/>에 적는 감싸개.</summary>
public sealed class MeteredLlmClient(ILlmClient inner, UsageMeter meter, string defaultModel) : ILlmClient
{
    public UsageMeter Meter => meter;

    public async Task<LlmCompletion> CompleteAsync(LlmRequest request, CancellationToken cancellationToken = default)
    {
        var completion = await inner.CompleteAsync(request, cancellationToken);
        meter.Add(request.Model ?? defaultModel, completion.Usage);
        return completion;
    }

    public IAsyncEnumerable<LlmStreamEvent> StreamAsync(LlmRequest request, CancellationToken cancellationToken = default) =>
        inner.StreamAsync(request, cancellationToken);

    public Task TestConnectionAsync(CancellationToken cancellationToken = default) => inner.TestConnectionAsync(cancellationToken);
}
