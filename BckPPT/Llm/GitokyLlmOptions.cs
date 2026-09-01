using System.IO;

namespace BckPPT.Llm;

public sealed record GitokyLlmOptions(
    Uri BaseUri,
    string ApiKey,
    string Model,
    TimeSpan Timeout,
    int MaxRetries)
{
    public static GitokyLlmOptions FromEnvironment(string baseDirectory)
    {
        var values = DotEnvLoader.Load(FindDotEnv(baseDirectory));
        string Read(string name, string fallback = "") =>
            Environment.GetEnvironmentVariable(name) is { Length: > 0 } systemValue
                ? systemValue.Trim()
                : values.GetValueOrDefault(name, fallback).Trim();

        var key = Read("GITOKY_LLM_API_KEY");
        if (string.IsNullOrWhiteSpace(key))
            throw new InvalidOperationException(".env에 GITOKY_LLM_API_KEY를 설정해 주세요.");

        var baseUrl = Read("GITOKY_LLM_BASE_URL", "https://gitoky.com").TrimEnd('/');
        if (baseUrl.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            baseUrl = baseUrl[..^3];
        return new GitokyLlmOptions(
            new Uri(baseUrl + "/", UriKind.Absolute), key,
            Read("GITOKY_LLM_MODEL", "auto"), TimeSpan.FromSeconds(90), 2);
    }

    private static string FindDotEnv(string baseDirectory)
    {
        for (var directory = new DirectoryInfo(baseDirectory); directory is not null; directory = directory.Parent)
        {
            var candidate = Path.Combine(directory.FullName, ".env");
            if (File.Exists(candidate)) return candidate;
        }
        return Path.Combine(baseDirectory, ".env");
    }
}

internal static class DotEnvLoader
{
    public static IReadOnlyDictionary<string, string> Load(string path)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!File.Exists(path)) return result;
        foreach (var raw in File.ReadLines(path))
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith('#')) continue;
            var separator = line.IndexOf('=');
            if (separator <= 0) continue;
            result[line[..separator].Trim()] = line[(separator + 1)..].Trim().Trim('"', '\'');
        }
        return result;
    }
}
