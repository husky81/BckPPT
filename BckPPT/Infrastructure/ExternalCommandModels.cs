using System.Collections.Concurrent;
using System.Text.Json.Serialization;

namespace BckPPT.Infrastructure;

public sealed record ExternalCommandRequest(
    string ProtocolVersion,
    string Command,
    string RequestId,
    string? FilePath = null,
    string? Instruction = null,
    string? SaveMode = null);

public sealed record ExternalCommandResponse(
    bool Accepted,
    string RequestId,
    string Status,
    string? Message = null,
    string? ResultPath = null);

public sealed class ExternalJobRegistry
{
    private readonly ConcurrentDictionary<string, ExternalCommandResponse> _jobs = new(StringComparer.OrdinalIgnoreCase);

    public void Set(ExternalCommandResponse response) => _jobs[response.RequestId] = response;
    public ExternalCommandResponse? Get(string requestId) => _jobs.GetValueOrDefault(requestId);
}
