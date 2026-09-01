namespace BckPPT.Llm;

public interface ILlmClient
{
    IAsyncEnumerable<LlmStreamEvent> StreamAsync(
        LlmRequest request,
        CancellationToken cancellationToken = default);

    Task TestConnectionAsync(CancellationToken cancellationToken = default);
}
