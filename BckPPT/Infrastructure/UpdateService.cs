using System.Net.Http;
using System.Text.Json;
using System.IO;

namespace BckPPT.Infrastructure;

public sealed record UpdateManifest(string Version, string PackageUrl, string Sha256, string? ReleaseNotes = null);
public sealed record UpdateCheckResult(bool IsUpdateAvailable, UpdateManifest? Manifest, string Message);

public sealed class UpdateService : IDisposable
{
    private readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(10) };
    private readonly Uri? _manifestUri;

    public UpdateService()
    {
        var raw = Environment.GetEnvironmentVariable("BCKPPT_UPDATE_MANIFEST_URL") ?? ReadFromDotEnv();
        _manifestUri = Uri.TryCreate(raw, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps ? uri : null;
    }

    private static string? ReadFromDotEnv()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            var path = Path.Combine(directory.FullName, ".env");
            if (!File.Exists(path)) continue;
            var line = File.ReadLines(path).FirstOrDefault(item => item.TrimStart().StartsWith("BCKPPT_UPDATE_MANIFEST_URL="));
            return line?.Split('=', 2)[1].Trim();
        }
        return null;
    }

    public async Task<UpdateCheckResult> CheckAsync(Version currentVersion, CancellationToken cancellationToken = default)
    {
        if (_manifestUri is null) return new(false, null, "자동 업데이트 주소가 설정되지 않았습니다.");
        using var response = await _http.GetAsync(_manifestUri, cancellationToken);
        response.EnsureSuccessStatusCode();
        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        var manifest = await JsonSerializer.DeserializeAsync<UpdateManifest>(stream,
            new JsonSerializerOptions(JsonSerializerDefaults.Web), cancellationToken)
            ?? throw new InvalidDataException("업데이트 manifest가 비어 있습니다.");
        if (!Version.TryParse(manifest.Version, out var available))
            throw new InvalidDataException("업데이트 manifest의 version이 올바르지 않습니다.");
        if (!Uri.TryCreate(manifest.PackageUrl, UriKind.Absolute, out var packageUri) || packageUri.Scheme != Uri.UriSchemeHttps)
            throw new InvalidDataException("업데이트 packageUrl은 HTTPS 절대 주소여야 합니다.");
        if (manifest.Sha256.Length != 64 || !manifest.Sha256.All(Uri.IsHexDigit))
            throw new InvalidDataException("업데이트 SHA-256 값이 올바르지 않습니다.");
        return available > currentVersion
            ? new(true, manifest, $"BckPPT {available} 업데이트를 사용할 수 있습니다.")
            : new(false, manifest, "최신 버전입니다.");
    }

    public void Dispose() => _http.Dispose();
}
