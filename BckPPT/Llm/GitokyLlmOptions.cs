using System.IO;

namespace BckPPT.Llm;

public sealed record GitokyLlmOptions(
    Uri BaseUri,
    string ApiKey,
    string Model,
    TimeSpan Timeout,
    int MaxRetries,
    string AgentModel = "agent",
    string VisionModel = "vision",
    string? WriterModel = null)
{
    public static GitokyLlmOptions FromEnvironment(string baseDirectory)
    {
        var values = DotEnvLoader.Load(FindDotEnv(baseDirectory));
        string Read(string name, string fallback = "") =>
            Environment.GetEnvironmentVariable(name) is { Length: > 0 } systemValue
                ? systemValue.Trim()
                : values.GetValueOrDefault(name, fallback).Trim();

        // CellBox 계열 .env는 GITOKY_LLM_KEY를 쓴다. 같은 키를 그대로 옮겨 붙여도 동작하게 둘 다 읽는다.
        var key = Read("GITOKY_LLM_API_KEY") is { Length: > 0 } apiKey ? apiKey : Read("GITOKY_LLM_KEY");
        if (string.IsNullOrWhiteSpace(key))
            throw new InvalidOperationException(".env에 GITOKY_LLM_API_KEY를 설정해 주세요.");

        var baseUrl = Read("GITOKY_LLM_BASE_URL", "https://gitoky.com").TrimEnd('/');
        if (baseUrl.EndsWith("/v1", StringComparison.OrdinalIgnoreCase))
            baseUrl = baseUrl[..^3];
        return new GitokyLlmOptions(
            new Uri(baseUrl + "/", UriKind.Absolute), key,
            Read("GITOKY_LLM_MODEL", "auto"), TimeSpan.FromSeconds(90), 2,
            // 도구 루프는 tools를 지원하는 라우트, 렌더 확인은 이미지 입력 라우트가 필요하다.
            Read("GITOKY_LLM_AGENT_MODEL", "agent"), Read("GITOKY_LLM_VISION_MODEL", "vision"),
            // 카드 문안처럼 한두 번만 쓰는 글쓰기 호출. 비우면 agent 모델을 쓴다(예: quality로 opus).
            Read("GITOKY_LLM_WRITER_MODEL") is { Length: > 0 } writer ? writer : null);
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
            // 빈 값은 건너뛴다. 템플릿을 붙여 넣다 남은 "KEY=" 줄이 앞의 실제 값을 지우지 않게.
            var value = line[(separator + 1)..].Trim().Trim('"', '\'');
            if (value.Length > 0) result[line[..separator].Trim()] = value;
        }
        return result;
    }
}
