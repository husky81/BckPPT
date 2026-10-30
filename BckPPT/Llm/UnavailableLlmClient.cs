namespace BckPPT.Llm;

/// <summary>키·설정이 없을 때 쓰는 자리. 화면은 이 타입을 보고 LLM 요청을 끄고 나머지 기능만 연다.</summary>
public sealed class UnavailableLlmClient(string reason) : ILlmClient
{
    public string Reason => reason;

    public async IAsyncEnumerable<LlmStreamEvent> StreamAsync(
        LlmRequest request,
        [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        await Task.CompletedTask;
        throw new LlmException(reason, "configuration_error");
#pragma warning disable CS0162
        yield break;
#pragma warning restore CS0162
    }

    public Task<LlmCompletion> CompleteAsync(LlmRequest request, CancellationToken cancellationToken = default) =>
        Task.FromException<LlmCompletion>(new LlmException(reason, "configuration_error"));

    public Task TestConnectionAsync(CancellationToken cancellationToken = default) =>
        Task.FromException(new LlmException(reason, "configuration_error"));
}
