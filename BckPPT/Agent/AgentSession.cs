using System.IO;
using System.Text.Json;
using BckPPT.Agent.Tools;
using BckPPT.Llm;
using BckPPT.Services;

namespace BckPPT.Agent;

/// <summary>
/// 한 대상 PPT에 대한 에이전트 작업. 시작할 때 작업 사본을 만들고 모든 도구는 그 사본만 고친다.
/// 승인하면 원본 옆에 결과 사본으로 남기고, 거절하면 작업 사본을 지운다. 원본은 어느 쪽이든 그대로다.
/// </summary>
public sealed class AgentSession : IDisposable
{
    private readonly List<LlmMessage> _messages = new();

    private AgentSession(string sourcePath, string workingPath, AgentToolContext context, AgentSessionLog log)
    {
        SourcePath = sourcePath;
        WorkingPath = workingPath;
        Context = context;
        Log = log;
    }

    public Guid Id { get; } = Guid.NewGuid();
    public string SourcePath { get; }
    public string WorkingPath { get; }
    public AgentToolContext Context { get; }
    public AgentSessionLog Log { get; }
    public LlmUsage Usage { get; private set; } = LlmUsage.Zero;
    public List<LlmMessage> Messages => _messages;
    public string? ResultPath { get; private set; }

    public static string DefaultWorkDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BckSoft", "BckPPT", "work");

    public static string DefaultLogDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BckSoft", "BckPPT", "sessions");

    public static AgentSession Start(
        IPptComService ppt, string sourcePath, string systemPrompt,
        string? workDirectory = null, string? logDirectory = null,
        Func<string, string, CancellationToken, Task<string>>? describeImage = null,
        ISourceLibrary? sources = null)
    {
        sourcePath = Path.GetFullPath(sourcePath);
        if (!File.Exists(sourcePath)) throw new FileNotFoundException("대상 PPT 파일을 찾을 수 없습니다.", sourcePath);
        workDirectory ??= DefaultWorkDirectory;
        Directory.CreateDirectory(workDirectory);
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var workingPath = Path.Combine(workDirectory,
            $"{Path.GetFileNameWithoutExtension(sourcePath)}_작업_{stamp}_{Guid.NewGuid().ToString("N")[..6]}{Path.GetExtension(sourcePath)}");
        File.Copy(sourcePath, workingPath, overwrite: false);

        var context = new AgentToolContext(ppt, workingPath) { DescribeImage = describeImage, Sources = sources };
        var log = new AgentSessionLog(Path.Combine(logDirectory ?? DefaultLogDirectory, $"{stamp}_{Path.GetFileNameWithoutExtension(workingPath)}.jsonl"));
        var session = new AgentSession(sourcePath, workingPath, context, log);
        log.Write("start", new { source = sourcePath, working = workingPath });
        session.Append(LlmMessage.System(systemPrompt));
        return session;
    }

    public void Append(LlmMessage message)
    {
        _messages.Add(message);
        Log.Write("message", new
        {
            role = message.Role,
            content = AgentSessionLog.Clip(message.Content),
            toolCalls = message.ToolCalls?.Select(c => new { c.Id, c.Name, c.ArgumentsJson }),
            message.ToolCallId
        });
    }

    public void AddUsage(LlmUsage usage) => Usage += usage;

    /// <summary>작업 사본을 원본 옆 결과 파일로 확정한다. 원본은 덮어쓰지 않는다.</summary>
    public string Approve()
    {
        if (ResultPath is not null) return ResultPath;
        var directory = Path.GetDirectoryName(SourcePath)!;
        var stem = Path.Combine(directory, $"{Path.GetFileNameWithoutExtension(SourcePath)}_BckPPT_{DateTime.Now:yyyyMMdd_HHmmss}");
        var extension = Path.GetExtension(SourcePath);
        var candidate = stem + extension;
        for (var suffix = 2; File.Exists(candidate); suffix++) candidate = $"{stem}_{suffix}{extension}";
        File.Copy(WorkingPath, candidate, overwrite: false);
        ResultPath = candidate;
        Log.Write("approve", new { result = candidate });
        return candidate;
    }

    /// <summary>작업 사본을 버린다. 확정한 결과 파일은 남긴다.</summary>
    public void Discard()
    {
        Log.Write("discard", new { });
        try { if (File.Exists(WorkingPath)) File.Delete(WorkingPath); } catch { }
    }

    public void Dispose() => Log.Dispose();
}

/// <summary>
/// 세션 대화·도구 호출·결과를 JSONL로 남긴다(재현·디버깅용). 원문 전체가 남지 않도록 긴 내용은 자른다.
/// 로그 실패는 작업을 막지 않는다.
/// </summary>
public sealed class AgentSessionLog : IDisposable
{
    private const int MaxLoggedChars = 2000;
    private readonly StreamWriter? _writer;

    public AgentSessionLog(string? path)
    {
        if (path is null) return;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            _writer = new StreamWriter(path, append: true) { AutoFlush = true };
            FilePath = path;
        }
        catch { _writer = null; }
    }

    public string? FilePath { get; }

    public static string Clip(string text) =>
        text.Length <= MaxLoggedChars ? text : text[..MaxLoggedChars] + $"…(+{text.Length - MaxLoggedChars}자)";

    public void Write(string kind, object data)
    {
        if (_writer is null) return;
        try
        {
            lock (_writer)
                _writer.WriteLine(JsonSerializer.Serialize(new { ts = DateTimeOffset.Now, kind, data }, ToolArgs.Json));
        }
        catch { }
    }

    public void Dispose() => _writer?.Dispose();
}
