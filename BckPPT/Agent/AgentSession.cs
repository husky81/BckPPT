using System.IO;
using System.Text.Json;
using BckPPT.Agent.Tools;
using BckPPT.Llm;
using BckPPT.PowerPoint;
using BckPPT.Services;

namespace BckPPT.Agent;

/// <summary>
/// 한 대상 PPT에 대한 에이전트 작업. 두 방식이 있다.
/// - 작업 사본: 사본을 만들어 고치고, 승인하면 원본 옆 결과 사본으로 남기고 거절하면 사본을 지운다. 원본은 그대로다.
/// - 바로 작업(<see cref="InPlace"/>): 사용자가 PowerPoint 창에 열어 둔 문서를 그 자리에서 고친다. 새로 만든 쪽만 고치고,
///   저장·삭제는 하지 않는다(사용자가 한다). 판정 기준은 시작할 때 디스크의 파일을 복사해 둔 기준본이다.
/// </summary>
public sealed class AgentSession : IDisposable
{
    private readonly List<LlmMessage> _messages = new();
    private readonly List<IDisposable> _openHandles = new();
    private Action<string, TimeSpan>? _slowCallHandler;

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

    /// <summary>열린 PowerPoint 창에서 바로 작업한다. 작업 파일이 곧 원본이다.</summary>
    public bool InPlace { get; private init; }

    /// <summary>작업 사본에서 뺀 편집 제한 내장 글꼴. 결과 파일에서는 설치된 글꼴로 대체되어 보인다.</summary>
    public IReadOnlyList<RestrictedFontRemover.RemovedFont> RemovedFonts { get; private init; } = [];

    public static string DefaultWorkDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BckSoft", "BckPPT", "work");

    public static string DefaultLogDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BckSoft", "BckPPT", "sessions");

    public static AgentSession Start(
        IPptComService ppt, string sourcePath, string systemPrompt,
        string? workDirectory = null, string? logDirectory = null,
        Func<string, string, CancellationToken, Task<string>>? describeImage = null,
        Func<string, string, CancellationToken, Task<string>>? summarize = null,
        Func<string, string, CancellationToken, Task<string>>? write = null,
        ISourceLibrary? sources = null,
        bool inPlace = false)
    {
        sourcePath = Path.GetFullPath(sourcePath);
        if (!File.Exists(sourcePath)) throw new FileNotFoundException("대상 PPT 파일을 찾을 수 없습니다.", sourcePath);
        workDirectory ??= DefaultWorkDirectory;
        Directory.CreateDirectory(workDirectory);
        // 끊긴 세션이 창 없이 남긴 사본·원본을 먼저 닫는다. 남아 있으면 PowerPoint가 느려져 같은 단계에서 또 시간 초과가 났다.
        IReadOnlyList<string> orphans = [];
        try { orphans = ppt.CloseOrphanedPresentations(workDirectory, [sourcePath]); }
        catch { /* 정리는 보조 작업이다. 실패해도 새 작업은 시작한다. */ }
        var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var suffix = Guid.NewGuid().ToString("N")[..6];
        string workingPath, baselinePath;
        IReadOnlyList<RestrictedFontRemover.RemovedFont> removedFonts = [];
        if (inPlace)
        {
            // 열린 창의 문서를 그대로 고친다. 줄 수 판정은 디스크의 파일(마지막 저장본)을 복사한 기준본과 비교한다.
            workingPath = sourcePath;
            baselinePath = Path.Combine(workDirectory,
                $"{Path.GetFileNameWithoutExtension(sourcePath)}_기준_{stamp}_{suffix}{Path.GetExtension(sourcePath)}");
            File.Copy(sourcePath, baselinePath, overwrite: false);
        }
        else
        {
            workingPath = Path.Combine(workDirectory,
                $"{Path.GetFileNameWithoutExtension(sourcePath)}_작업_{stamp}_{suffix}{Path.GetExtension(sourcePath)}");
            File.Copy(sourcePath, workingPath, overwrite: false);
            baselinePath = sourcePath;
            // 편집을 막는 내장 글꼴(미리 보기·인쇄 전용)이 있으면 작업 사본에서만 뺀다. 없으면 PowerPoint가 읽기 전용으로 연다.
            try { removedFonts = RestrictedFontRemover.Remove(workingPath); }
            catch { /* 판정할 수 없는 파일이면 그대로 두고 PowerPoint에 맡긴다. */ }
        }

        var context = new AgentToolContext(ppt, workingPath)
        {
            DescribeImage = describeImage, Summarize = summarize, Write = write, Sources = sources, BaselinePath = baselinePath,
            InPlace = inPlace
        };
        var log = new AgentSessionLog(Path.Combine(logDirectory ?? DefaultLogDirectory, $"{stamp}_{Path.GetFileNameWithoutExtension(workingPath)}.jsonl"));
        var session = new AgentSession(sourcePath, workingPath, context, log) { RemovedFonts = removedFonts, InPlace = inPlace };
        if (removedFonts.Count > 0)
            log.Write("restricted_fonts_removed", removedFonts.Select(f => $"{f.Typeface} ({f.Style}, fsType 0x{f.FsType:X4})"));
        // 세션 동안 작업 사본과 원본(판정 기준)을 열어 둔다. 31MB 착수보고 PPT에서 도구마다 열고 닫으면 너무 느리다.
        try
        {
            if (inPlace)
            {
                session._openHandles.Add(ppt.WatchUserActivity(workingPath));
                session._openHandles.Add(ppt.KeepOpen(baselinePath, readOnly: true));
            }
            else
            {
                session._openHandles.Add(ppt.KeepOpen(workingPath, readOnly: false));
                session._openHandles.Add(ppt.KeepOpen(sourcePath, readOnly: true));
            }
        }
        catch (Exception ex) { log.Write("keep_open_failed", new { ex.Message }); }
        if (orphans.Count > 0) log.Write("closed_orphans", orphans);
        session._slowCallHandler = (caller, elapsed) => log.Write("slow_com", new { caller, seconds = Math.Round(elapsed.TotalSeconds, 1) });
        PptComService.SlowCall += session._slowCallHandler;
        log.Write("start", new { source = sourcePath, working = workingPath, inPlace, baseline = baselinePath });
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

    private (int Chars, int Tokens)? _firstObservation, _lastObservation;

    /// <summary>보낸 대화 글자 수와 실제 입력 토큰을 기록한다. 문맥 예산 추정에 쓴다.</summary>
    public void ObserveContext(int sentChars, int promptTokens)
    {
        if (promptTokens <= 0) return;
        _firstObservation ??= (sentChars, promptTokens);
        _lastObservation = (sentChars, promptTokens);
    }

    /// <summary>
    /// 입력 토큰 ≈ 고정 비용(도구 정의 등) + 대화 글자 수 / 글자당 토큰.
    /// 관측 두 개의 차이로 기울기를 재고, 아직 차이가 작으면 고정 비용까지 글자에 얹은 보수적 비율만 쓴다.
    /// </summary>
    public int? EstimateCharBudget(int targetTokens)
    {
        if (_lastObservation is not var (chars, tokens)) return null;
        if (_firstObservation is var (firstChars, firstTokens) && tokens - firstTokens >= 1000 && chars > firstChars)
        {
            double charsPerToken = (double)(chars - firstChars) / (tokens - firstTokens);
            double overhead = Math.Max(0, tokens - chars / charsPerToken);
            return (int)((targetTokens - overhead) * charsPerToken);
        }
        return (int)(targetTokens * ((double)chars / tokens));
    }

    /// <summary>작업 사본을 원본 옆 결과 파일로 확정한다. 원본은 덮어쓰지 않는다.</summary>
    public string Approve()
    {
        if (ResultPath is not null) return ResultPath;
        if (InPlace)
        {
            // 바로 작업은 결과 사본이 없다. 저장은 사용자가 PowerPoint에서 한다.
            ReleaseOpenDocuments();
            DeleteBaseline();
            ResultPath = SourcePath;
            Log.Write("approve", new { result = SourcePath, inPlace = true });
            return SourcePath;
        }
        ReleaseOpenDocuments();   // 열린 채로는 복사가 막힐 수 있다
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
        Log.Write("discard", new { InPlace });
        ReleaseOpenDocuments();
        // 바로 작업이면 작업 파일이 곧 원본이다. 절대 지우지 않는다(추가된 쪽도 사용자가 직접 지운다).
        if (InPlace) { DeleteBaseline(); return; }
        try { if (File.Exists(WorkingPath)) File.Delete(WorkingPath); } catch { }
    }

    private void DeleteBaseline()
    {
        if (Context.BaselinePath is not { } baseline || string.Equals(baseline, SourcePath, StringComparison.OrdinalIgnoreCase)) return;
        try { if (File.Exists(baseline)) File.Delete(baseline); } catch { }
    }

    private void ReleaseOpenDocuments()
    {
        foreach (var handle in _openHandles) handle.Dispose();
        _openHandles.Clear();
    }

    public void Dispose()
    {
        if (_slowCallHandler is not null) PptComService.SlowCall -= _slowCallHandler;
        ReleaseOpenDocuments();
        Log.Dispose();
    }
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
