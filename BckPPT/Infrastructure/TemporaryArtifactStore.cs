using System.Collections.Concurrent;
using System.IO;

namespace BckPPT.Infrastructure;

public sealed class TemporaryArtifactStore : IDisposable
{
    private readonly ConcurrentBag<string> _paths = [];

    public string Track(string path)
    {
        _paths.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var path in _paths)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }
}
