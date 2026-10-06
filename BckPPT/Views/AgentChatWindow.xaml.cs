using System.IO;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using BckPPT.Agent;
using BckPPT.Application;
using BckPPT.Llm;
using BckPPT.Models;
using BckPPT.Services;
using BckPPT.Rag;
using BckPPT.Infrastructure;

namespace BckPPT.Views;

public partial class AgentChatWindow : Window, INotifyPropertyChanged
{
    private readonly DocumentSession _session;
    private readonly IPptComService _pptService;
    private readonly ILlmClient _llmClient;
    private readonly ApprovalWorkflow _approvalWorkflow;
    private readonly AgentWorkflow _agentWorkflow;
    private readonly PptIndexer _pptIndexer;
    private readonly ExternalJobRegistry _externalJobs;
    private readonly UpdateService _updateService;
    private readonly Action _openEditor;
    private readonly string? _initialDirectory;
    private CancellationTokenSource? _requestCts;
    private string _conversationText = "파일을 이 창에 놓거나 위 버튼으로 문서를 연결한 뒤 작업을 요청하세요.";
    private string _streamingText = "";
    private bool _isRequesting;
    private PptJob? _currentJob;
    private AgentTurnResult? _agentResult;

    public string TargetText => _session.TargetText;
    public string StatusText { get; private set; } = "준비됨";
    public string ConversationText { get => _conversationText; private set { _conversationText = value; Notify(); } }
    public string StreamingText { get => _streamingText; private set { _streamingText = value; Notify(); } }
    public bool IsRequesting { get => _isRequesting; private set { _isRequesting = value; Notify(); Notify(nameof(CanSend)); Notify(nameof(CanApply)); } }
    public bool CanSend => !IsRequesting;
    public bool CanPreview => _currentJob?.State == PptJobState.Planned;
    public bool CanApply => _currentJob?.State == PptJobState.PreviewReady && !IsRequesting;
    public Visibility JobControlsVisibility => _currentJob is null || _currentJob.State is PptJobState.Cancelled
        ? Visibility.Collapsed : Visibility.Visible;
    public IReadOnlyList<string> PreviewImages => _agentResult?.BeforeThumbnails ?? _currentJob?.BeforeThumbnails ?? [];
    public IReadOnlyList<string> ResultImages => _agentResult?.AfterThumbnails ?? _currentJob?.AfterThumbnails ?? [];
    public Visibility AgentControlsVisibility => _agentWorkflow.Session is null ? Visibility.Collapsed : Visibility.Visible;
    public string SourcesText => _agentWorkflow.AttachedSources.Count == 0
        ? "첨부 원문 없음"
        : $"첨부 원문 {_agentWorkflow.AttachedSources.Count}개: " +
          string.Join(", ", _agentWorkflow.AttachedSources.Select(Path.GetFileName));
    public Visibility PreviewImagesVisibility => PreviewImages.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ResultImagesVisibility => ResultImages.Count > 0 ? Visibility.Visible : Visibility.Collapsed;

    public AgentChatWindow(DocumentSession session, IPptComService pptService, ILlmClient llmClient,
        ApprovalWorkflow approvalWorkflow, AgentWorkflow agentWorkflow, PptIndexer pptIndexer, ExternalJobRegistry externalJobs,
        UpdateService updateService, Action openEditor, string? initialDirectory = null)
    {
        _initialDirectory = initialDirectory;
        _session = session;
        _pptService = pptService;
        _llmClient = llmClient;
        _approvalWorkflow = approvalWorkflow;
        _agentWorkflow = agentWorkflow;
        _pptIndexer = pptIndexer;
        _externalJobs = externalJobs;
        _updateService = updateService;
        _openEditor = openEditor;
        DataContext = this;
        session.PropertyChanged += (_, _) => Dispatcher.Invoke(() => RefreshBindings());
        InitializeComponent();
        Loaded += async (_, _) => await TestLlmConnectionAsync();
        Loaded += async (_, _) => await CheckForUpdatesAsync();
        Loaded += (_, _) =>
        {
            var health = EnvironmentHealthService.Check();
            if (!health.IsPowerPointInstalled) SetStatus(health.Message);
        };
        Closed += (_, _) =>
        {
            _requestCts?.Cancel();
            // 승인하지 않은 작업 사본은 남기지 않는다.
            _agentWorkflow.Discard();
        };
    }

    private void OpenFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "PowerPoint 파일|*.pptx;*.pptm;*.ppt|모든 파일|*.*" };
        if (_initialDirectory is not null) dialog.InitialDirectory = _initialDirectory;
        if (dialog.ShowDialog(this) == true) OpenPresentation(dialog.FileName);
    }

    private void Connect_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var open = _pptService.GetOpenPresentations();
            if (open.Count == 0) { SetStatus("실행 중인 PowerPoint 문서가 없습니다."); return; }
            _session.Presentation = open[0];
            SetStatus(open.Count == 1 ? "현재 문서에 연결했습니다." : $"열린 문서 {open.Count}개 중 첫 문서에 연결했습니다.");
        }
        catch (Exception ex) { SetStatus($"연결 실패: {ex.Message}"); }
    }

    private void OpenEditor_Click(object sender, RoutedEventArgs e) => _openEditor();

    private async void IndexFolder_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog { Title = "PPT 템플릿 라이브러리 폴더 선택", Multiselect = false };
        if (dialog.ShowDialog(this) != true) return;
        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        IsRequesting = true;
        try
        {
            var progress = new Progress<string>(SetStatus);
            var result = await _pptIndexer.IndexFolderAsync(dialog.FolderName, progress, _requestCts.Token);
            SetStatus($"인덱싱 완료 · 파일 {result.Files}개/슬라이드 {result.Slides}장 갱신 · {result.Skipped}개 변경 없음");
        }
        catch (OperationCanceledException) { SetStatus("인덱싱을 취소했습니다."); }
        catch (Exception ex) { SetStatus($"인덱싱 실패: {ex.Message}"); }
        finally { IsRequesting = false; }
    }

    private async void Send_Click(object sender, RoutedEventArgs e) => await SendAsync();

    public void AcceptExternalRequest(ExternalCommandRequest request)
    {
        try
        {
            ExternalCommandValidator.Validate(request);
            if (request.Command != "create_job") return;
            Activate();
            OpenPresentation(request.FilePath!);
            InstructionBox.Text = request.Instruction!;
            _externalJobs.Set(new ExternalCommandResponse(true, request.RequestId, "planning", "작업계획을 생성하는 중입니다."));
            _ = SendPlanAsync(request.RequestId);
        }
        catch (Exception ex)
        {
            _externalJobs.Set(new ExternalCommandResponse(false, request.RequestId, "failed", ex.Message));
            SetStatus($"외부 요청 거부: {ex.Message}");
        }
    }

    private async void InstructionBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter || Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) return;
        e.Handled = true;
        await SendAsync();
    }

    private Task SendAsync() => SendAgentAsync();

    private void Cancel_Click(object sender, RoutedEventArgs e) => _requestCts?.Cancel();

    /// <summary>
    /// 도구 루프로 작업한다. 같은 대상에 대한 후속 요청은 같은 작업 사본과 대화에서 이어진다.
    /// 결과는 "결과 저장"으로 승인해야 원본 옆 파일로 남는다.
    /// </summary>
    private async Task SendAgentAsync()
    {
        var instruction = InstructionBox.Text.Trim();
        if (instruction.Length == 0 || IsRequesting) return;
        if (_session.Presentation?.Path is not { Length: > 0 } target)
        {
            SetStatus("먼저 대상 PPT를 열거나 PowerPoint에 연결해 주세요.");
            return;
        }

        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        IsRequesting = true;
        _currentJob = null;
        bool continuing = _agentWorkflow.Session is not null;
        var log = new StringBuilder(continuing ? ConversationText + "\n\n" : "");
        log.Append($"요청\n{instruction}\n");
        ConversationText = log.ToString();
        StreamingText = "";
        SetStatus(continuing ? "같은 작업 사본에서 이어서 작업합니다..." : "작업 사본을 만들고 시작합니다 · 원본은 저장된 내용 기준이며 바뀌지 않습니다.");
        InstructionBox.Clear();

        var progress = new Progress<AgentEvent>(e =>
        {
            switch (e.Kind)
            {
                case AgentEventKind.Thinking: SetStatus($"생각하는 중 · {e.Text}"); break;
                case AgentEventKind.ToolStarted:
                    SetStatus($"도구 실행 · {e.ToolName}");
                    log.Append($"· {e.ToolName} {Clip(e.Text, 120)}\n");
                    break;
                case AgentEventKind.ToolFinished:
                    log.Append(e.IsError ? "   ⚠ " : "   → ")
                       .Append(Clip(FirstLines(e.Text, e.ToolName is "lint_slide" or "finish" ? 6 : 2), 300)).Append('\n');
                    break;
                case AgentEventKind.Message: log.Append($"{e.Text}\n"); break;
                case AgentEventKind.Compacted: log.Append($"   ({e.Text})\n"); break;
            }
            ConversationText = log.ToString();
        });

        try
        {
            _agentResult = await _agentWorkflow.RunAsync(target, instruction, progress, _requestCts.Token);
            var run = _agentResult.Run;
            if (!continuing && _agentWorkflow.Session?.RemovedFonts is { Count: > 0 } fonts)
                log.Append("\n(안내) 편집이 제한된 내장 글꼴 ")
                   .Append(string.Join(", ", fonts.Select(f => f.Typeface).Distinct()))
                   .Append("을 작업 사본에서 제거했습니다. 이 PC에 없는 글꼴이면 대체 글꼴로 보이고 줄바꿈이 원래와 다를 수 있습니다.\n");
            // 도구 없이 끝난 답은 이미 메시지로 찍혔으므로 다시 쓰지 않는다.
            log.Append(run.Reason == AgentStopReason.AnsweredWithoutTools ? "\n" : $"\n결과\n{run.Summary}\n");
            if (run.RemainingIssues.Count > 0)
                log.Append($"남은 레이아웃 위반 {run.RemainingIssues.Count}건\n")
                   .Append(string.Join("\n", run.RemainingIssues.Select(i => "- " + i.Message))).Append('\n');
            ConversationText = log.ToString();
            SetStatus(run.Reason switch
            {
                AgentStopReason.Finished or AgentStopReason.AnsweredWithoutTools =>
                    "완료 · 미리보기를 확인하고 '결과 저장' 또는 후속 요청을 해 주세요.",
                AgentStopReason.TurnLimit => "턴 상한에 도달해 멈췄습니다 · 이어서 요청하거나 결과를 확인해 주세요.",
                _ => "토큰 상한에 도달해 멈췄습니다."
            } + $" · {run.Turns}턴");
            if (_agentResult.UsageByModel is { Count: > 0 } byModel)
                log.Append("LLM 사용: ").Append(UsageMeter.Describe(byModel)).Append('\n');
            ConversationText = log.ToString();
        }
        catch (OperationCanceledException) { SetStatus("작업을 중단했습니다. 지금까지의 변경은 작업 사본에 남아 있습니다."); }
        catch (LlmException ex) { SetStatus(ToUserMessage(ex) + " · 작업 사본은 유지됩니다. 다시 요청하면 이어갑니다."); }
        catch (Exception ex) { SetStatus($"작업 실패: {ex.Message}"); }
        finally
        {
            IsRequesting = false;
            NotifyJobState();
        }
    }

    private void ApproveAgent_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var result = _agentWorkflow.Approve();
            ConversationText += $"\n\n저장함: {result}";
            SetStatus("결과 사본을 저장했습니다 · 원본은 바뀌지 않았습니다. 다음 요청은 원본에서 새로 시작합니다.");
            _agentResult = null;
        }
        catch (Exception ex) { SetStatus($"저장 실패: {ex.Message} · PowerPoint에서 작업 사본을 닫고 다시 시도해 주세요."); }
        NotifyJobState();
    }

    private void OpenWorkingCopy_Click(object sender, RoutedEventArgs e)
    {
        if (_agentWorkflow.Session?.WorkingPath is not { } path) return;
        try
        {
            _pptService.OpenPptxInPowerPoint(path);
            SetStatus("작업 사본을 PowerPoint에서 열었습니다 · 저장 전 결과입니다.");
        }
        catch (Exception ex) { SetStatus($"열기 실패: {ex.Message}"); }
    }

    private void DiscardAgent_Click(object sender, RoutedEventArgs e)
    {
        _agentWorkflow.Discard();
        _agentResult = null;
        ConversationText += "\n\n작업 사본을 버렸습니다.";
        SetStatus("작업을 버렸습니다 · 원본은 그대로입니다.");
        NotifyJobState();
    }

    private async void AttachSources_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "원문 첨부 (기획보고서·RFP 등)",
            Filter = "원문|*.pdf;*.pptx;*.md;*.markdown;*.txt|모든 파일|*.*",
            Multiselect = true
        };
        if (_initialDirectory is not null) dialog.InitialDirectory = _initialDirectory;
        if (dialog.ShowDialog(this) == true) await AttachSourcesAsync(dialog.FileNames);
    }

    private async Task AttachSourcesAsync(IReadOnlyList<string> paths)
    {
        if (IsRequesting) return;
        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        IsRequesting = true;
        try
        {
            SetStatus($"원문 {paths.Count}개를 색인하는 중...");
            var (files, pages) = await _agentWorkflow.AttachAsync(paths, _requestCts.Token);
            SetStatus($"원문 {files}개 · {pages}쪽 색인 완료 · 다음 요청부터 에이전트가 검색합니다.");
        }
        catch (OperationCanceledException) { SetStatus("원문 색인을 취소했습니다."); }
        catch (Exception ex) { SetStatus($"원문 첨부 실패: {ex.Message}"); }
        finally
        {
            IsRequesting = false;
            Notify(nameof(SourcesText));
        }
    }

    private static string FirstLines(string text, int count) =>
        string.Join(" / ", text.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0).Take(count));

    private static string Clip(string text, int max) => text.Length <= max ? text : text[..max] + "…";

    /// <summary>외부 요청(create_job)용 단발 계획 흐름. ApprovalWorkflow의 미리보기·적용 단계를 그대로 쓴다.</summary>
    private async Task SendPlanAsync(string? externalRequestId = null)
    {
        var instruction = InstructionBox.Text.Trim();
        if (instruction.Length == 0 || IsRequesting) return;

        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        IsRequesting = true;
        StreamingText = "";
        ConversationText = $"요청\n{instruction}\n\n작업 계획";
        SetStatus("작업계획을 생성하는 중...");
        var output = new StringBuilder();

        try
        {
            string? documentContext = null;
            if (_session.Presentation is { } target)
            {
                SetStatus("PPT 구조를 분석해 최소 문맥을 구성하는 중...");
                var analysis = await Task.Run(() => _pptService.AnalyzePresentation(
                    target.Path, cancellationToken: _requestCts.Token), _requestCts.Token);
                documentContext = PptContextAssembler.Assemble(_session, analysis);
            }
            await foreach (var evt in _llmClient.StreamAsync(
                PptPromptBuilder.BuildPlanRequest(instruction, _session, documentContext), _requestCts.Token))
            {
                if (evt.Kind == LlmStreamEventKind.Started)
                    SetStatus($"LLM 연결됨 · {evt.Route ?? "경로 미상"} · 요청 {evt.RequestId ?? "ID 없음"}");
                if (evt.Kind == LlmStreamEventKind.ContentDelta && !string.IsNullOrEmpty(evt.Content))
                {
                    output.Append(evt.Content);
                    StreamingText = output.ToString();
                }
            }
            var plan = PptTaskPlanDecoder.Decode(output.ToString());
            StreamingText = PptTaskPlanDecoder.ToDisplayJson(plan);
            _currentJob = _approvalWorkflow.Create(instruction, plan, _session);
            _currentJob.ExternalRequestId = externalRequestId;
            NotifyJobState();
            SetStatus("구조화 작업계획을 받았습니다. 변경 미리보기를 확인해 주세요.");
            if (externalRequestId is not null)
                _externalJobs.Set(new ExternalCommandResponse(true, externalRequestId, "awaiting_preview", "사용자 미리보기 승인을 기다립니다."));
            InstructionBox.Clear();
        }
        catch (OperationCanceledException) { SetStatus("요청을 취소했습니다."); SetExternalFailure(externalRequestId, "cancelled", "요청이 취소되었습니다."); }
        catch (LlmException ex) { SetStatus(ToUserMessage(ex)); SetExternalFailure(externalRequestId, "failed", ex.Message); }
        catch (FormatException ex) { SetStatus($"작업계획 검증 실패: {ex.Message} · 다시 만들어 주세요."); SetExternalFailure(externalRequestId, "failed", ex.Message); }
        catch (Exception ex) { SetStatus($"LLM 요청 실패: {ex.Message}"); SetExternalFailure(externalRequestId, "failed", ex.Message); }
        finally { IsRequesting = false; }
    }

    private async Task TestLlmConnectionAsync()
    {
        try
        {
            await _llmClient.TestConnectionAsync();
            SetStatus("GitOky LLM 연결됨");
        }
        catch (Exception ex) { SetStatus($"LLM 연결 안 됨: {ex.Message} · 전문 편집 기능은 사용할 수 있습니다."); }
    }

    private async Task CheckForUpdatesAsync()
    {
        try
        {
            var current = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0);
            var result = await _updateService.CheckAsync(current);
            if (result.IsUpdateAvailable) SetStatus(result.Message);
        }
        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"업데이트 확인 실패: {ex.Message}"); }
    }

    private async void Preview_Click(object sender, RoutedEventArgs e)
    {
        if (_currentJob is null) return;
        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        IsRequesting = true;
        try
        {
            SetStatus("PowerPoint에서 적용 전 미리보기를 렌더링하는 중...");
            StreamingText = await _approvalWorkflow.PreparePreviewAsync(_currentJob, _requestCts.Token);
            SetStatus("미리보기 준비됨 · 적용 전까지 원본은 변경되지 않습니다.");
            if (_currentJob.ExternalRequestId is { } requestId)
                _externalJobs.Set(new ExternalCommandResponse(true, requestId, "awaiting_approval", "적용 승인을 기다립니다."));
        }
        catch (OperationCanceledException) { SetStatus("미리보기 생성을 취소했습니다."); SetExternalFailure(_currentJob.ExternalRequestId, "cancelled", "미리보기가 취소되었습니다."); }
        catch (Exception ex) { SetStatus($"미리보기 실패: {ex.Message}"); SetExternalFailure(_currentJob.ExternalRequestId, "failed", ex.Message); }
        finally { IsRequesting = false; NotifyJobState(); }
    }

    private async void Apply_Click(object sender, RoutedEventArgs e)
    {
        if (_currentJob is null) return;
        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        IsRequesting = true;
        try
        {
            SetStatus("승인된 작업을 사본에 적용하는 중...");
            var result = await _approvalWorkflow.ApplyAsync(_currentJob, _requestCts.Token);
            var validation = _currentJob.Validation;
            StreamingText += $"\n\n완료: {result}\n" + FormatValidation(validation);
            SetStatus(validation?.IsValid == true
                ? "사본 적용 및 레이아웃 검증 완료 · 원본은 변경되지 않았습니다."
                : "사본 적용 완료 · 해결되지 않은 레이아웃 경고를 확인해 주세요.");
            if (_currentJob.ExternalRequestId is { } requestId)
                _externalJobs.Set(new ExternalCommandResponse(true, requestId, "completed", "작업이 완료되었습니다.", result));
        }
        catch (OperationCanceledException) { SetStatus("적용을 취소했습니다. 불완전한 사본은 삭제했습니다."); SetExternalFailure(_currentJob.ExternalRequestId, "cancelled", "적용이 취소되었습니다."); }
        catch (Exception ex) { SetStatus($"적용 실패: {ex.Message}"); SetExternalFailure(_currentJob.ExternalRequestId, "failed", ex.Message); }
        finally { IsRequesting = false; NotifyJobState(); }
    }

    private void CancelJob_Click(object sender, RoutedEventArgs e)
    {
        if (_currentJob is null) return;
        if (_currentJob.State == PptJobState.Applying) _requestCts?.Cancel();
        _approvalWorkflow.Cancel(_currentJob);
        if (_currentJob.ExternalRequestId is { } requestId)
            _externalJobs.Set(new ExternalCommandResponse(false, requestId, "cancelled", "사용자가 작업을 취소했습니다."));
        SetStatus("작업을 취소했습니다.");
        NotifyJobState();
    }

    private void NotifyJobState()
    {
        Notify(nameof(CanPreview));
        Notify(nameof(CanApply));
        Notify(nameof(JobControlsVisibility));
        Notify(nameof(PreviewImages));
        Notify(nameof(ResultImages));
        Notify(nameof(PreviewImagesVisibility));
        Notify(nameof(ResultImagesVisibility));
        Notify(nameof(AgentControlsVisibility));
    }

    private static string FormatValidation(BckPPT.PowerPoint.LayoutValidationResult? result)
    {
        if (result is null) return "레이아웃 검증 결과 없음";
        if (result.IsValid) return result.CorrectedCount == 0
            ? "레이아웃 검증: overflow 없음"
            : $"레이아웃 검증: {result.CorrectedCount}개 자동 보정, overflow 없음";
        var issues = string.Join(", ", result.RemainingIssues.Select(issue => $"{issue.SlideIndex}쪽/{issue.ShapeName}"));
        return $"레이아웃 검증: {result.CorrectedCount}개 자동 보정, {result.RemainingIssues.Count}개 경고 ({issues})";
    }

    private void SetExternalFailure(string? requestId, string status, string message)
    {
        if (requestId is not null)
            _externalJobs.Set(new ExternalCommandResponse(false, requestId, status, message));
    }

    private static string ToUserMessage(LlmException ex) => ex.ErrorCode switch
    {
        "quota_exceeded" => "GitOky LLM 사용 한도를 초과했습니다.",
        _ when ex.Message.Contains("quota", StringComparison.OrdinalIgnoreCase) =>
            $"GitOky LLM 일일 토큰 한도를 다 썼습니다. 한도가 초기화된 뒤 다시 요청해 주세요. ({ex.Message})",
        "timeout" => "LLM 응답 시간이 초과되었습니다. 다시 시도해 주세요.",
        _ when ex.StatusCode == 401 || ex.StatusCode == 403 => "GitOky LLM 키 또는 권한을 확인해 주세요.",
        _ => $"LLM 요청 실패: {ex.Message}"
    };

    private void Window_DragOver(object sender, DragEventArgs e) =>
        e.Effects = HasPresentationFile(e.Data) || DroppedSources(e.Data).Count > 0 ? DragDropEffects.Copy : DragDropEffects.None;

    /// <summary>PPT는 작업 대상으로 열고, PDF·MD·TXT와 폴더는 원문으로 첨부한다.</summary>
    private async void Window_Drop(object sender, DragEventArgs e)
    {
        var path = (e.Data.GetData(DataFormats.FileDrop) as string[])?.FirstOrDefault(IsPresentationFile);
        if (path is not null) OpenPresentation(path);
        var sources = DroppedSources(e.Data);
        if (sources.Count > 0) await AttachSourcesAsync(sources);
    }

    private static IReadOnlyList<string> DroppedSources(IDataObject data) =>
        (data.GetData(DataFormats.FileDrop) as string[] ?? [])
        .Where(p => Directory.Exists(p) || (!IsPresentationFile(p) && SourceLibrary.SupportedExtensions.Contains(Path.GetExtension(p))))
        .ToList();

    private void OpenPresentation(string path)
    {
        try
        {
            _pptService.OpenPptxInPowerPoint(path);
            _session.Presentation = new PptPresentation(Path.GetFileName(path), Path.GetFullPath(path), 0);
            SetStatus("PowerPoint에서 문서를 열었습니다.");
        }
        catch (Exception ex) { SetStatus($"파일 열기 실패: {ex.Message}"); }
    }

    private static bool HasPresentationFile(IDataObject data) =>
        (data.GetData(DataFormats.FileDrop) as string[])?.Any(IsPresentationFile) == true;

    private static bool IsPresentationFile(string path) =>
        new[] { ".pptx", ".pptm", ".ppt" }.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    private void SetStatus(string value) { StatusText = value; Notify(nameof(StatusText)); }
    private void RefreshBindings() { Notify(nameof(TargetText)); Notify(nameof(StatusText)); }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void Notify([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
