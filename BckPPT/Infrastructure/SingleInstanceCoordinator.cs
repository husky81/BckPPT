using System.IO.Pipes;
using System.IO;
using System.Text;

namespace BckPPT.Infrastructure;

public sealed class SingleInstanceCoordinator : IDisposable
{
    public const string PipeName = "BckPPT.Agent.v1";
    private readonly FileStream? _lockHandle;
    public bool IsPrimary { get; }

    public SingleInstanceCoordinator()
    {
        var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BckSoft", "BckPPT");
        Directory.CreateDirectory(directory);
        try
        {
            _lockHandle = new FileStream(Path.Combine(directory, "instance.lock"), FileMode.OpenOrCreate,
                FileAccess.ReadWrite, FileShare.None);
            IsPrimary = true;
        }
        catch (IOException) { IsPrimary = false; }
    }

    public static async Task<string> SendAsync(string json, CancellationToken cancellationToken = default)
    {
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(7));
        var token = timeout.Token;
        await using var pipe = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut, PipeOptions.Asynchronous);
        await pipe.ConnectAsync(5000, token).ConfigureAwait(false);
        await using var writer = new StreamWriter(pipe, new UTF8Encoding(false), leaveOpen: true) { AutoFlush = true };
        using var reader = new StreamReader(pipe, Encoding.UTF8, leaveOpen: true);
        await writer.WriteLineAsync(json.AsMemory(), token).ConfigureAwait(false);
        return await reader.ReadLineAsync(token).ConfigureAwait(false) ?? "";
    }

    public void Dispose()
    {
        _lockHandle?.Dispose();
    }
}
