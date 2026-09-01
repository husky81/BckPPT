namespace BckPPT.Llm;

public sealed class UnavailableLlmClient(string reason) : ILlmClient
{
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

    public Task TestConnectionAsync(CancellationToken cancellationToken = default) =>
        Task.FromException(new LlmException(reason, "configuration_error"));
}
