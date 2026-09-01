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
    private readonly PptIndexer _pptIndexer;
    private readonly ExternalJobRegistry _externalJobs;
    private readonly UpdateService _updateService;
    private readonly Action _openEditor;
    private CancellationTokenSource? _requestCts;
    private string _conversationText = "파일을 이 창에 놓거나 위 버튼으로 문서를 연결한 뒤 작업을 요청하세요.";
    private string _streamingText = "";
    private bool _isRequesting;
    private PptJob? _currentJob;

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
    public IReadOnlyList<string> PreviewImages => _currentJob?.BeforeThumbnails ?? [];
    public IReadOnlyList<string> ResultImages => _currentJob?.AfterThumbnails ?? [];
    public Visibility PreviewImagesVisibility => PreviewImages.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
    public Visibility ResultImagesVisibility => ResultImages.Count > 0 ? Visibility.Visible : Visibility.Collapsed;

    public AgentChatWindow(DocumentSession session, IPptComService pptService, ILlmClient llmClient,
        ApprovalWorkflow approvalWorkflow, PptIndexer pptIndexer, ExternalJobRegistry externalJobs,
        UpdateService updateService, Action openEditor)
    {
        _session = session;
        _pptService = pptService;
        _llmClient = llmClient;
        _approvalWorkflow = approvalWorkflow;
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
        Closed += (_, _) => _requestCts?.Cancel();
    }

    private void OpenFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "PowerPoint 파일|*.pptx;*.pptm;*.ppt|모든 파일|*.*" };
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
            _ = SendAsync(request.RequestId);
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

    private void Cancel_Click(object sender, RoutedEventArgs e) => _requestCts?.Cancel();

    private async Task SendAsync(string? externalRequestId = null)
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
        "timeout" => "LLM 응답 시간이 초과되었습니다. 다시 시도해 주세요.",
        _ when ex.StatusCode == 401 || ex.StatusCode == 403 => "GitOky LLM 키 또는 권한을 확인해 주세요.",
        _ => $"LLM 요청 실패: {ex.Message}"
    };

    private void Window_DragOver(object sender, DragEventArgs e) =>
        e.Effects = HasPresentationFile(e.Data) ? DragDropEffects.Copy : DragDropEffects.None;

    private void Window_Drop(object sender, DragEventArgs e)
    {
        var path = (e.Data.GetData(DataFormats.FileDrop) as string[])?.FirstOrDefault(IsPresentationFile);
        if (path is not null) OpenPresentation(path);
    }

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
