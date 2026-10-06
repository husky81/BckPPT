namespace BckPPT.Llm;

public interface ILlmClient
{
    IAsyncEnumerable<LlmStreamEvent> StreamAsync(
        LlmRequest request,
        CancellationToken cancellationToken = default);

    /// <summary>비스트리밍 호출. 도구 루프는 응답 전체(도구 호출 포함)가 필요하므로 이것을 쓴다.</summary>
    Task<LlmCompletion> CompleteAsync(
        LlmRequest request,
        CancellationToken cancellationToken = default);

    Task TestConnectionAsync(CancellationToken cancellationToken = default);
}
