using System.IO.Pipes;
using System.IO;
using System.Text;
using System.Text.Json;

namespace BckPPT.Infrastructure;

public sealed class ExternalCommandEndpoint(
    ExternalJobRegistry registry,
    Func<ExternalCommandRequest, Task> createJob,
    Action? showWindow = null) : IDisposable
{
    private readonly CancellationTokenSource _cts = new();
    private Task? _serverTask;
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    public void Start() => _serverTask = RunAsync(_cts.Token);

    private async Task RunAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            try
            {
                await using var pipe = new NamedPipeServerStream(SingleInstanceCoordinator.PipeName,
                    PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
                await pipe.WaitForConnectionAsync(cancellationToken);
                using var reader = new StreamReader(pipe, Encoding.UTF8, leaveOpen: true);
                await using var writer = new StreamWriter(pipe, new UTF8Encoding(false), leaveOpen: true) { AutoFlush = true };
                var line = await reader.ReadLineAsync(cancellationToken);
                var response = await HandleAsync(line, cancellationToken);
                await writer.WriteLineAsync(JsonSerializer.Serialize(response, JsonOptions).AsMemory(), cancellationToken);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { break; }
            catch { }
        }
    }

    private async Task<ExternalCommandResponse> HandleAsync(string? json, CancellationToken cancellationToken)
    {
        ExternalCommandRequest? request = null;
        try
        {
            request = JsonSerializer.Deserialize<ExternalCommandRequest>(json ?? "", JsonOptions)
                ?? throw new ArgumentException("요청 JSON이 비어 있습니다.");
            ExternalCommandValidator.Validate(request);
            if (request.Command == "get_job_status")
                return registry.Get(request.RequestId) ?? new(false, request.RequestId, "not_found", "작업을 찾을 수 없습니다.");
            if (request.Command == ExternalCommandValidator.ShowWindow)
            {
                showWindow?.Invoke();
                return new ExternalCommandResponse(true, request.RequestId, "completed", "채팅창을 띄웠습니다.");
            }
            var accepted = new ExternalCommandResponse(true, request.RequestId, "accepted", "채팅 세션으로 전달했습니다.");
            registry.Set(accepted);
            await createJob(request);
            return accepted;
        }
        catch (Exception ex)
        {
            return new ExternalCommandResponse(false, request?.RequestId ?? "", "rejected", ex.Message);
        }
    }

    public void Dispose()
    {
        _cts.Cancel();
        try { _serverTask?.Wait(TimeSpan.FromSeconds(2)); } catch { }
        _cts.Dispose();
    }
}
