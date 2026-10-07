using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using BckPPT.Agent;
using BckPPT.Application;
using BckPPT.Llm;
using BckPPT.Models;
using BckPPT.Services;
using BckPPT.Rag;
using BckPPT.Infrastructure;

namespace BckPPT.Views;

/// <summary>
/// PowerPoint 옆에 붙어 쓰는 에이전트 콘솔. 메인 창에는 대상 PPT 한 줄(상태줄), 작업 흐름, 입력란만 둔다.
/// 나머지(파일 열기·원문 첨부·인덱싱·전문 편집창·설정)는 ⋯ 메뉴에 있다.
/// 대상은 PowerPoint에서 보고 있는 문서를 2초마다 확인해 따라간다(작업 중에는 확인하지 않는다).
/// </summary>
public partial class AgentChatWindow : Window, INotifyPropertyChanged
{
    private readonly DocumentSession _session;
    private readonly IPptComService _pptService;
    private readonly ILlmClient _llmClient;
    private readonly AgentWorkflow _agentWorkflow;
    private readonly PptIndexer _pptIndexer;
    private readonly ExternalJobRegistry _externalJobs;
    private readonly UpdateService _updateService;
    private readonly Action _openEditor;
    private readonly string? _initialDirectory;
    private readonly DispatcherTimer _followTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private readonly DispatcherTimer _spinTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
    private CancellationTokenSource? _requestCts;
    private bool _isRequesting;
    private bool _polling;
    private bool _followConversation = true;
    private bool _inPlaceNoticeShown;
    private string _statusDetail = "준비됨";
    private int? _currentSlide;
    private int _slideCount;
    private ConsoleTurn? _runningTurn;

    public ObservableCollection<ConsoleTurn> Turns { get; } = new();

    // ── 상태줄 ──
    public Brush LightBrush => (Brush)FindResource(
        _session.Presentation is null ? "ErrorBr" : IsRequesting || !StaDispatcher.Instance.IsHealthy ? "WarnBr" : "SuccessBr");
    public string LightTooltip => _session.Presentation is null
        ? "작업할 PowerPoint 문서가 없습니다 · PowerPoint에서 문서를 열거나 ⋯ > PPT 파일 열기"
        : IsRequesting ? "작업 중"
        : !StaDispatcher.Instance.IsHealthy ? "PowerPoint 응답을 기다리는 중"
        : $"작업할 수 있습니다 · {_statusDetail}";
    public string TargetTooltip => _session.Presentation?.Path ?? "대상 문서 없음";
    public string FileNameDisplay { get; private set; } = "PowerPoint 문서 없음";
    public string FolderDisplay { get; private set; } = "";
    public ObservableCollection<OpenFileItem> OpenFiles { get; } = new();
    public bool NoOpenFiles => OpenFiles.Count == 0;
    private bool _pinned;
    /// <summary>고정하면 PowerPoint에서 보고 있는 문서를 따라가지 않는다.</summary>
    public bool IsPinned { get => _pinned; private set { _pinned = value; Notify(); Notify(nameof(PinLabel)); UpdateStatusLayout(); } }
    public string PinLabel => IsPinned ? "📌 고정 풀기 (보고 있는 문서 따라가기)" : "📌 이 문서로 고정";
    public bool HasTarget => _session.Presentation is not null;
    public string SlideDisplay => _session.Presentation is null ? ""
        : _currentSlide is int s ? $"{s} / {_slideCount}" : _slideCount > 0 ? $"{_slideCount}쪽" : "";

    // ── 메뉴 ──
    public bool HasSources => _agentWorkflow.AttachedSources.Count > 0;
    public string AttachMenuText => HasSources ? $"원문 첨부 ({_agentWorkflow.AttachedSources.Count})…" : "원문 첨부…";
    public bool WorkInOpenWindow
    {
        get => _agentWorkflow.PreferInPlace;
        set { _agentWorkflow.PreferInPlace = value; PptRules.Current.WorkInOpenWindow = value; PptRules.Current.Save(); Notify(); }
    }
    public bool LlmFit
    {
        get => _agentWorkflow.LlmFit;
        set { _agentWorkflow.LlmFit = value; PptRules.Current.LlmLengthFit = value; PptRules.Current.Save(); Notify(); }
    }
    public bool StayOnTop
    {
        get => Topmost;
        set { Topmost = value; PptRules.Current.ChatTopmost = value; PptRules.Current.Save(); Notify(); }
    }
    private bool InPlaceSession => _agentWorkflow.Session?.InPlace == true;
    public bool HasSession => _agentWorkflow.Session is not null;
    public string ApproveLabel => InPlaceSession ? "작업 마침" : "결과 저장";
    public string DiscardLabel => InPlaceSession ? "대화 끝내기" : "작업 버리기";
    public bool ShowOpenWorkingCopy => HasSession && !InPlaceSession;

    // ── 입력 ──
    public bool IsRequesting
    {
        get => _isRequesting;
        private set
        {
            _isRequesting = value;
            Notify(); Notify(nameof(CanSend)); Notify(nameof(SendGlyph)); Notify(nameof(SendTooltip));
            RefreshLight();
            if (value) _spinTimer.Start(); else _spinTimer.Stop();
        }
    }
    public bool CanSend => !IsRequesting;
    public string SendGlyph => IsRequesting ? "■" : "↑";
    public string SendTooltip => IsRequesting ? "멈추기 (Esc)" : "보내기 (Enter)";
    public string EmptyHint => "PowerPoint에서 작업할 문서를 열어 두고 요청을 입력하세요.\n\n" +
                               "예) 35쪽 카드 템플릿으로 첨부한 보고서마다 카드를 하나씩 만들어 37쪽 뒤에 넣어줘\n" +
                               "예) 38~41쪽 공백줄 삭제해줘\n예) 이 md로 43쪽 뒤에 간트형 일정표 그려줘";
    public bool ShowEmptyHint => Turns.Count == 0;
    private bool _showJumpToEnd;
    public bool ShowJumpToEnd { get => _showJumpToEnd; private set { _showJumpToEnd = value; Notify(); } }

    public AgentChatWindow(DocumentSession session, IPptComService pptService, ILlmClient llmClient,
        ApprovalWorkflow approvalWorkflow, AgentWorkflow agentWorkflow, PptIndexer pptIndexer, ExternalJobRegistry externalJobs,
        UpdateService updateService, Action openEditor, string? initialDirectory = null)
    {
        _initialDirectory = initialDirectory;
        _session = session;
        _pptService = pptService;
        _llmClient = llmClient;
        _agentWorkflow = agentWorkflow;
        // 사용자 고급 모델로 카드 문안 받기는 화면에서 뺐다. 기능 코드는 남기되 꺼 둔다.
        _agentWorkflow.ExternalWriting = false;
        _agentWorkflow.LlmFit = PptRules.Current.LlmLengthFit;
        _agentWorkflow.PreferInPlace = PptRules.Current.WorkInOpenWindow;
        _pptIndexer = pptIndexer;
        _externalJobs = externalJobs;
        _updateService = updateService;
        _openEditor = openEditor;
        DataContext = this;
        session.PropertyChanged += (_, _) => Dispatcher.Invoke(RefreshTarget);
        Turns.CollectionChanged += (_, _) => Notify(nameof(ShowEmptyHint));
        InitializeComponent();
        RestoreBounds();

        _followTimer.Tick += async (_, _) => await FollowPowerPointAsync();
        _spinTimer.Tick += (_, _) => { _runningTurn?.RunningStep?.Tick(); _runningTurn?.Tick(); };
        SourceInitialized += (_, _) => UseDarkTitleBar();
        Loaded += async (_, _) =>
        {
            var health = EnvironmentHealthService.Check();
            if (!health.IsPowerPointInstalled) AddNote(health.Message, warning: true);
            RefreshTarget();
            FocusInput();
            _followTimer.Start();
            await FollowPowerPointAsync();
            await TestLlmConnectionAsync();
            await CheckForUpdatesAsync();
        };
        Closing += (_, _) => SaveBounds();
        Closed += (_, _) =>
        {
            _followTimer.Stop();
            _requestCts?.Cancel();
            // 작업 사본이면 지우고, 바로 작업이면 아무것도 지우지 않는다.
            _agentWorkflow.Discard();
        };
    }

    // ══ 요청 ══════════════════════════════════════════════════

    private async void SendOrStop_Click(object sender, RoutedEventArgs e)
    {
        if (IsRequesting) { StopRequest(); return; }
        await SendAgentAsync();
    }

    /// <summary>
    /// Enter·Ctrl+Enter는 보내기, Shift+Enter·Alt+Enter는 줄바꿈. 여러 줄 붙여넣기는 그대로 들어간다.
    /// </summary>
    private async void InstructionBox_KeyDown(object sender, KeyEventArgs e)
    {
        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        if (key != Key.Enter) return;
        var mods = Keyboard.Modifiers;
        if (mods.HasFlag(ModifierKeys.Shift) || mods.HasFlag(ModifierKeys.Alt))
        {
            // Shift+Enter는 TextBox가 줄을 바꾼다. Alt+Enter는 시스템 키라 직접 넣는다.
            if (!mods.HasFlag(ModifierKeys.Alt)) return;
            int caret = InstructionBox.CaretIndex;
            InstructionBox.Text = InstructionBox.Text.Insert(caret, Environment.NewLine);
            InstructionBox.CaretIndex = caret + Environment.NewLine.Length;
            e.Handled = true;
            return;
        }
        e.Handled = true;
        await SendAgentAsync();
    }

    private bool _userCancelled;

    private void StopRequest()
    {
        _userCancelled = true;
        _requestCts?.Cancel();
    }

    private void InstructionBox_TextChanged(object sender, TextChangedEventArgs e) =>
        Placeholder.Visibility = InstructionBox.Text.Length == 0 ? Visibility.Visible : Visibility.Collapsed;

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Escape) return;
        if (MenuPopup.IsOpen) { MenuPopup.IsOpen = false; e.Handled = true; }
        else if (IsRequesting) { StopRequest(); e.Handled = true; }
    }

    /// <summary>도구 루프로 작업한다. 같은 대상에 대한 후속 요청은 같은 세션에서 이어진다.</summary>
    private async Task<AgentTurnResult?> SendAgentAsync()
    {
        var instruction = InstructionBox.Text.Trim();
        if (instruction.Length == 0 || IsRequesting) return null;
        // 보내자마자 요청 줄과 기다리는 표시를 띄운다. 대상 찾기·문서 열기에도 몇 초 걸린다.
        var turn = new ConsoleTurn(instruction);
        Turns.Add(turn);
        _runningTurn = turn;
        IsRequesting = true;
        InstructionBox.Clear();
        turn.StartWaiting();
        ScrollToEndNow();

        // 요청에 PPT 경로가 있으면 그것을 대상으로 연다. 열린 문서의 이름 조각이면 그 문서로 바꾼다.
        if (AgentWorkflow.MentionedPresentationPath(instruction) is { } mentionedPptx && _session.Presentation?.Path is null)
            OpenPresentation(mentionedPptx);
        if (await SwitchToMentionedFileAsync(instruction) is { } switchNote) turn.Add(new NoteEntry(switchNote));
        if (_session.Presentation?.Path is not { Length: > 0 } target)
        {
            turn.AddOutcome(new NoteEntry("작업할 PowerPoint 문서가 없습니다. PowerPoint에서 문서를 열거나 ⋯ > PPT 파일 열기로 고른 뒤 다시 보내 주세요.", true));
            turn.Finish(0);
            _runningTurn = null;
            IsRequesting = false;
            InstructionBox.Text = instruction;   // 요청 문구는 되살려 둔다
            return null;
        }

        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        _userCancelled = false;
        // 사용자가 누르지 않았는데 취소된 일이 두 번 있었다(2026-10-07, 재현 안 됨). 그때 호출 경로를 남긴다.
        _requestCts.Token.Register(() =>
        {
            if (_userCancelled) return;
            try { File.AppendAllText(Path.Combine(AgentSession.DefaultLogDirectory, "cancel_trace.txt"), $"{DateTime.Now:HH:mm:ss}\n{Environment.StackTrace}\n\n"); }
            catch { }
        });
        bool continuing = _agentWorkflow.Session is not null;
        var progress = new Progress<AgentEvent>(ev => OnAgentEvent(turn, ev));

        AgentTurnResult? result = null;
        try
        {
            result = await _agentWorkflow.RunAsync(target, instruction, progress, _requestCts.Token);
            ShowResult(turn, result, continuing);
        }
        catch (OperationCanceledException) when (_requestCts.IsCancellationRequested)
        {
            AbortSteps(turn, "멈춤");
            turn.AddOutcome(new NoteEntry("작업을 멈췄습니다. 지금까지 바뀐 것은 그대로 있습니다.", true));
        }
        catch (OperationCanceledException ex)
        {
            // 사용자가 멈춘 것이 아니면 응답 대기 시간 초과다(HttpClient는 시간 초과도 취소 예외로 낸다).
            AbortSteps(turn, "시간 초과");
            turn.AddOutcome(new NoteEntry($"LLM 응답을 기다리다 시간이 초과되었습니다. 다시 요청하면 이어서 합니다. ({ex.GetBaseException().Message})", true));
        }
        catch (LlmException ex) { AbortSteps(turn, "중단"); turn.AddOutcome(new NoteEntry(ToUserMessage(ex), true)); }
        catch (PptUserInterventionException ex) { AbortSteps(turn, "멈춤"); turn.AddOutcome(new NoteEntry(ex.Message, true)); }
        catch (Exception ex) { AbortSteps(turn, "실패"); turn.AddOutcome(new NoteEntry($"작업 실패: {ex.Message}", true)); }
        finally
        {
            turn.Finish(result?.AfterSlides.Count ?? result?.AfterThumbnails.Count ?? 0);
            _runningTurn = null;
            IsRequesting = false;
            FocusInput();
            NotifySession();
            if (_followConversation) ScrollToEndNow();
        }
        return result;
    }

    /// <summary>
    /// 요청에 열린 문서의 이름 조각("추진일정 파일")이 있으면 그 문서로 바꾼다. 여러 개가 맞으면 바꾸지 않고 알린다.
    /// 경로를 적은 경우는 앞에서 이미 처리했다. 돌려준 문장은 이번 요청의 작업 흐름 맨 앞에 붙는다.
    /// </summary>
    private async Task<string?> SwitchToMentionedFileAsync(string instruction)
    {
        IReadOnlyList<OpenPresentationInfo> open;
        try { open = await Task.Run(_pptService.ListOpenWindows); } catch { return null; }
        var match = OpenFileMatcher.Match(instruction, open);
        if (match.Kind == FileMatchKind.Ambiguous)
            return $"'{match.Token}'에 맞는 문서가 {match.Candidates.Count}개라 대상을 바꾸지 않았습니다({string.Join(", ", match.Candidates.Select(c => c.Name))}). " +
                   "상태줄에서 고르거나 이름을 더 적어 주세요.";
        if (match is not { Kind: FileMatchKind.One, File: { } file }) return null;
        if (string.Equals(file.Path, _session.Presentation?.Path, StringComparison.OrdinalIgnoreCase)) return null;
        await SelectFileAsync(file.Name, file.Path, file.SlideCount);
        return $"요청에 적힌 '{match.Token}' → {file.Name}에서 작업합니다.";
    }

    /// <summary>대상을 바꾸고 PowerPoint에서도 그 문서를 앞으로 가져온다(대상 = 보고 있는 문서).</summary>
    private async Task SelectFileAsync(string name, string path, int slideCount)
    {
        bool hadSession = _agentWorkflow.Session is not null &&
                          !string.Equals(_agentWorkflow.Session.SourcePath, path, StringComparison.OrdinalIgnoreCase);
        _session.Presentation = new PptPresentation(name, path, slideCount);
        _slideCount = slideCount;
        _currentSlide = null;
        RefreshTarget();
        try { await Task.Run(() => _pptService.ActivatePresentation(path)); } catch { }
        // PowerPoint는 그 문서로 바뀐 채 뒤에 두고, 바로 이어서 입력할 수 있게 BckPPT를 앞으로 가져온다.
        Activate();
        FocusInput();
        if (hadSession) AddNote($"대상을 {name}(으)로 바꿨습니다. 다음 요청은 이 문서에서 새로 시작합니다.");
        await FollowPowerPointAsync();
    }

    // ── 상태줄 문서 목록 ──

    private async void FileName_Click(object sender, RoutedEventArgs e)
    {
        if (IsRequesting) return;
        IReadOnlyList<OpenPresentationInfo> open = [];
        try { open = await Task.Run(_pptService.ListOpenWindows); } catch { }
        OpenFiles.Clear();
        foreach (var o in open.OrderByDescending(o => o.IsSaved).ThenBy(o => o.Name, StringComparer.CurrentCulture))
            OpenFiles.Add(new OpenFileItem(o.Name, o.Path, o.IsSaved ? ConsoleFormatting.ShortenFolder(System.IO.Path.GetDirectoryName(o.Path)) : "",
                o.SlideCount, string.Equals(o.Path, _session.Presentation?.Path, StringComparison.OrdinalIgnoreCase), o.IsSaved));
        Notify(nameof(NoOpenFiles)); Notify(nameof(HasTarget)); Notify(nameof(PinLabel));
        FilePopup.IsOpen = true;
    }

    private async void OpenFileItem_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is not OpenFileItem item) return;
        FilePopup.IsOpen = false;
        if (!item.IsSaved)
        {
            AddNote($"{item.Name}은(는) 저장되지 않은 새 문서라 작업할 수 없습니다. PowerPoint에서 먼저 저장해 주세요.", warning: true);
            return;
        }
        await SelectFileAsync(item.Name, item.Path, item.SlideCount);
    }

    private void Pin_Click(object sender, RoutedEventArgs e)
    {
        FilePopup.IsOpen = false;
        IsPinned = !IsPinned;
        _statusDetail = IsPinned ? "고정한 문서에서 작업" : "보고 있는 문서를 따라감";
        RefreshLight();
        if (!IsPinned) _ = FollowPowerPointAsync();
    }

    /// <summary>진행 이벤트를 작업 흐름 항목으로 바꾼다.</summary>
    private void OnAgentEvent(ConsoleTurn turn, AgentEvent ev)
    {
        switch (ev.Kind)
        {
            case AgentEventKind.ToolStarted:
            {
                turn.StopWaiting();
                var line = ConsoleFormatting.Describe(ev.ToolName ?? "", ev.Text);
                if (!line.Hidden) turn.Add(new ToolStepEntry(line.Label, line.Detail, ev.ToolName ?? ""));
                break;
            }
            case AgentEventKind.ToolFinished:
            {
                if (turn.RunningStep is { } step && step.ToolName == ev.ToolName)
                    step.Complete(ConsoleFormatting.ResultLine(ev.Text, ev.IsError), ev.IsError);
                else if (ev.ToolName == "finish" && ev.IsError)
                    turn.Add(new NarrationEntry("레이아웃 검사에서 위반이 남아 다시 고칩니다."));
                // 도구가 끝나면 다시 모델을 기다린다.
                if (turn.RunningStep is null) turn.StartWaiting();
                break;
            }
            case AgentEventKind.Message:
                // 도구 실행 중의 메시지는 그 도구의 안쪽 진행(카드 문안 작성 2/4 등), 아니면 모델의 설명이다.
                if (turn.RunningStep is { } running) running.Progress(ev.Text);
                else turn.Add(new NarrationEntry(ev.Text));
                break;
            case AgentEventKind.Compacted:
                turn.Add(new NoteEntry(ev.Text));
                break;
        }
        if (_followConversation) ScrollToEndSoon();
    }

    private void ShowResult(ConsoleTurn turn, AgentTurnResult result, bool continuing)
    {
        var run = result.Run;
        // 도구 없이 끝난 답은 이미 설명으로 찍혔다. 마지막 설명을 답으로 옮긴다.
        if (run.Reason == AgentStopReason.AnsweredWithoutTools && turn.Activity.LastOrDefault() is NarrationEntry last && last.Text == run.Summary.Trim())
            turn.Activity.Remove(last);
        if (run.Reason == AgentStopReason.UserIntervened) turn.AddOutcome(new NoteEntry(run.Summary, true));
        else if (!string.IsNullOrWhiteSpace(run.Summary)) turn.AddOutcome(new AnswerEntry(run.Summary.Trim()));
        if (run.Reason == AgentStopReason.TurnLimit) turn.AddOutcome(new NoteEntry("턴 상한에 도달해 멈췄습니다. 이어서 요청하면 계속합니다.", true));
        if (run.Reason == AgentStopReason.TokenLimit) turn.AddOutcome(new NoteEntry("토큰 상한에 도달해 멈췄습니다.", true));

        if (InPlaceSession && !_inPlaceNoticeShown && result.Changes.Count > 0)
        {
            _inPlaceNoticeShown = true;
            turn.AddOutcome(new NoteEntry("열린 PowerPoint에서 바로 작업했습니다. 새로 만든 쪽만 고쳤고 저장은 하지 않았습니다. 저장은 PowerPoint에서 해 주세요."));
        }
        if (!continuing && _agentWorkflow.Session?.RemovedFonts is { Count: > 0 } fonts)
            turn.AddOutcome(new NoteEntry($"편집이 제한된 내장 글꼴({string.Join(", ", fonts.Select(f => f.Typeface).Distinct())})을 작업 사본에서 뺐습니다. 대체 글꼴로 보일 수 있습니다."));
        if (run.RemainingIssues.Count > 0)
            turn.AddOutcome(new NoteEntry($"남은 레이아웃 위반 {run.RemainingIssues.Count}건 · " +
                string.Join(" · ", run.RemainingIssues.Take(3).Select(i => i.Message)) + (run.RemainingIssues.Count > 3 ? " …" : ""), true));
        if (result.UsageByModel is { Count: > 0 } usage && usage.Values.Sum(v => v.Usage.TotalTokens) >= 50_000)
            turn.AddOutcome(new NoteEntry("LLM 사용 " + UsageMeter.Describe(usage)));

        foreach (var change in result.Changes) turn.Changes.Add(change);
        var slides = result.AfterSlides;
        for (int i = 0; i < result.AfterThumbnails.Count; i++)
            turn.Thumbnails.Add(new ThumbnailItem(result.AfterThumbnails[i], i < slides.Count ? slides[i] : null));
    }

    private static void AbortSteps(ConsoleTurn turn, string text)
    {
        foreach (var step in turn.Activity.OfType<ToolStepEntry>().Where(s => s.IsRunning)) step.Complete(text, failed: true);
    }

    private void ToggleActivity_Click(object sender, RoutedEventArgs e) => ((sender as FrameworkElement)?.DataContext as ConsoleTurn)?.ToggleActivity();
    private void ToggleChanges_Click(object sender, RoutedEventArgs e) => ((sender as FrameworkElement)?.DataContext as ConsoleTurn)?.ToggleChanges();

    private async void Thumbnail_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.DataContext is not ThumbnailItem { Slide: int slide }) return;
        if ((_agentWorkflow.Session?.WorkingPath ?? _session.Presentation?.Path) is not { } path) return;
        try { await Task.Run(() => _pptService.GoToSlide(path, slide)); } catch { }
    }

    /// <summary>외부 요청(create_job)은 에이전트 흐름으로 처리한다.</summary>
    public void AcceptExternalRequest(ExternalCommandRequest request) => _ = RunExternalAsync(request);

    private async Task RunExternalAsync(ExternalCommandRequest request)
    {
        try
        {
            ExternalCommandValidator.Validate(request);
            if (request.Command != "create_job") return;
            Activate();
            OpenPresentation(request.FilePath!);
            InstructionBox.Text = request.Instruction!;
            _externalJobs.Set(new ExternalCommandResponse(true, request.RequestId, "running", "작업 중입니다."));
            var result = await SendAgentAsync();
            _externalJobs.Set(result is null
                ? new ExternalCommandResponse(false, request.RequestId, "failed", "작업을 시작하지 못했습니다.")
                : new ExternalCommandResponse(true, request.RequestId, "completed", result.Run.Summary));
        }
        catch (Exception ex)
        {
            _externalJobs.Set(new ExternalCommandResponse(false, request.RequestId, "failed", ex.Message));
            AddNote($"외부 요청 거부: {ex.Message}", warning: true);
        }
    }

    // ══ 상태줄: PowerPoint 따라가기 ═════════════════════════════

    /// <summary>
    /// PowerPoint에서 보고 있는 문서와 쪽을 따라간다. 작업 중에는 묻지 않는다(대상이 바뀌면 안 되고 COM 호출이 겹친다).
    /// </summary>
    private async Task FollowPowerPointAsync()
    {
        if (_polling || IsRequesting || !StaDispatcher.Instance.IsHealthy) { RefreshLight(); return; }
        _polling = true;
        try
        {
            var active = await Task.Run(() => { try { return _pptService.GetActivePresentation(); } catch { return null; } });
            if (IsRequesting) return;
            if (active is not null && IsPinned && _session.Presentation is not null)
            {
                // 고정 중: 대상은 그대로. 보고 있는 문서가 대상이면 쪽만 따라간다.
                bool same = string.Equals(active.Path, _session.Presentation.Path, StringComparison.OrdinalIgnoreCase);
                _currentSlide = same ? active.CurrentSlide : null;
                _statusDetail = "고정한 문서에서 작업" + (same ? "" : " · PowerPoint에서는 다른 문서를 보고 있음");
            }
            else if (active is not null)
            {
                bool changed = !string.Equals(active.Path, _session.Presentation?.Path, StringComparison.OrdinalIgnoreCase);
                _currentSlide = active.CurrentSlide;
                _slideCount = active.SlideCount;
                if (changed)
                {
                    bool hadSession = _agentWorkflow.Session is not null;
                    _session.Presentation = new PptPresentation(active.Name, active.Path, active.SlideCount);
                    if (hadSession) AddNote($"대상이 {active.Name}(으)로 바뀌었습니다. 다음 요청은 이 문서에서 새로 시작합니다.");
                }
                _statusDetail = InPlaceSession || (WorkInOpenWindow && _agentWorkflow.Session is null)
                    ? "열린 창에서 바로 작업" : "작업 사본에서 작업";
            }
            else
            {
                _currentSlide = null;
                _statusDetail = _session.Presentation is null ? "PowerPoint 문서 없음" : "PowerPoint 창에 열려 있지 않음 · 작업 사본에서 작업";
            }
            RefreshTarget();
        }
        finally { _polling = false; }
    }

    private void RefreshTarget()
    {
        _slideCount = _slideCount == 0 ? _session.Presentation?.SlideCount ?? 0 : _slideCount;
        UpdateStatusLayout();
        Notify(nameof(SlideDisplay));
        Notify(nameof(TargetTooltip));
        Notify(nameof(HasTarget));
        RefreshLight();
    }

    private void RefreshLight() { Notify(nameof(LightBrush)); Notify(nameof(LightTooltip)); }

    private void StatusBar_SizeChanged(object sender, SizeChangedEventArgs e) => UpdateStatusLayout();

    /// <summary>
    /// 파일명은 선명하게, 경로는 작은 회색으로. 좁아지면 경로부터 숨기고 파일명 가운데를 줄인다.
    /// </summary>
    private void UpdateStatusLayout()
    {
        if (FileNameText is null) return;
        var path = _session.Presentation?.Path;
        string name = path is null ? "PowerPoint 문서 없음" : Path.GetFileName(path);
        string folder = path is null ? "" : ConsoleFormatting.ShortenFolder(Path.GetDirectoryName(path));
        // 신호등·여백·쪽 표시·⋯ 단추를 뺀 폭
        // 단추 여백·▾·(고정이면) 📌 몫도 뺀다.
        double available = StatusBar.ActualWidth - 9 - 9 - 8 - 30 - Measure(SlideDisplay, FileNameText) - 16 - 30 - (IsPinned ? 18 : 0);
        double nameWidth = Measure(name, FileNameText);
        double folderWidth = folder.Length == 0 ? 0 : Measure(folder, FolderText) + 8;
        if (nameWidth + folderWidth <= available) { FileNameDisplay = name; FolderDisplay = folder; }
        else
        {
            FolderDisplay = "";
            FileNameDisplay = ConsoleFormatting.MiddleEllipsis(name, available, s => Measure(s, FileNameText));
        }
        Notify(nameof(FileNameDisplay));
        Notify(nameof(FolderDisplay));
    }

    private double Measure(string text, TextBlock style) =>
        new FormattedText(text, CultureInfo.CurrentUICulture, FlowDirection.LeftToRight,
            new Typeface(style.FontFamily, style.FontStyle, style.FontWeight, style.FontStretch), style.FontSize, Brushes.White,
            VisualTreeHelper.GetDpi(this).PixelsPerDip).WidthIncludingTrailingWhitespace;

    // ══ ⋯ 메뉴 ═══════════════════════════════════════════════

    private void More_Click(object sender, RoutedEventArgs e)
    {
        Notify(nameof(AttachMenuText)); Notify(nameof(HasSources)); NotifySession();
        MenuPopup.IsOpen = !MenuPopup.IsOpen;
    }

    private void CloseMenu() { MenuPopup.IsOpen = false; FocusInput(); }

    /// <summary>입력창에 커서를 둔다(시작·문서 고르기·요청 끝·메뉴 닫기).</summary>
    private void FocusInput() =>
        Dispatcher.BeginInvoke(() => { InstructionBox.Focus(); Keyboard.Focus(InstructionBox); InstructionBox.CaretIndex = InstructionBox.Text.Length; },
            DispatcherPriority.Input);

    private void OpenFile_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        var dialog = new OpenFileDialog { Filter = "PowerPoint 파일|*.pptx;*.pptm;*.ppt|모든 파일|*.*" };
        if (_initialDirectory is not null) dialog.InitialDirectory = _initialDirectory;
        if (dialog.ShowDialog(this) == true) OpenPresentation(dialog.FileName);
    }

    private void OpenEditor_Click(object sender, RoutedEventArgs e) { CloseMenu(); _openEditor(); }

    private async void IndexFolder_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        var dialog = new OpenFolderDialog { Title = "PPT 템플릿 라이브러리 폴더 선택", Multiselect = false };
        if (dialog.ShowDialog(this) != true || IsRequesting) return;
        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        IsRequesting = true;
        try
        {
            var result = await _pptIndexer.IndexFolderAsync(dialog.FolderName, new Progress<string>(SetStatus), _requestCts.Token);
            AddNote($"템플릿 인덱싱 완료 · 파일 {result.Files}개/슬라이드 {result.Slides}장 갱신 · {result.Skipped}개 변경 없음");
        }
        catch (OperationCanceledException) { AddNote("인덱싱을 멈췄습니다."); }
        catch (Exception ex) { AddNote($"인덱싱 실패: {ex.Message}", warning: true); }
        finally { IsRequesting = false; }
    }

    private async void AttachSources_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        var dialog = new OpenFileDialog
        {
            Title = "원문 첨부 (기획보고서·RFP 등)",
            Filter = "원문|*.pdf;*.pptx;*.md;*.markdown;*.txt|모든 파일|*.*",
            Multiselect = true
        };
        if (_initialDirectory is not null) dialog.InitialDirectory = _initialDirectory;
        if (dialog.ShowDialog(this) == true) await AttachSourcesAsync(dialog.FileNames);
    }

    private void DetachSources_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        int count = _agentWorkflow.AttachedSources.Count;
        _agentWorkflow.DetachAll();
        AddNote($"첨부 원문 {count}개를 뺐습니다.");
        Notify(nameof(HasSources)); Notify(nameof(AttachMenuText));
    }

    private async Task AttachSourcesAsync(IReadOnlyList<string> paths)
    {
        if (IsRequesting) return;
        _requestCts?.Dispose();
        _requestCts = new CancellationTokenSource();
        IsRequesting = true;
        try
        {
            var (files, pages) = await _agentWorkflow.AttachAsync(paths, _requestCts.Token);
            AddNote($"원문 {files}개 · {pages}쪽을 첨부했습니다. 다음 요청부터 에이전트가 찾아 씁니다.");
        }
        catch (OperationCanceledException) { AddNote("원문 첨부를 멈췄습니다."); }
        catch (Exception ex) { AddNote($"원문 첨부 실패: {ex.Message}", warning: true); }
        finally
        {
            IsRequesting = false;
            Notify(nameof(HasSources)); Notify(nameof(AttachMenuText));
        }
    }

    private void Usage_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        var meter = _agentWorkflow.Session?.Context.Meter;
        AddNote(meter is null ? "이번 작업에서 아직 LLM을 쓰지 않았습니다." : "이번 작업의 LLM 사용 · " + UsageMeter.Describe(meter.Snapshot()));
    }

    private void OpenLogs_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        try
        {
            Directory.CreateDirectory(AgentSession.DefaultLogDirectory);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", AgentSession.DefaultLogDirectory) { UseShellExecute = true });
        }
        catch (Exception ex) { AddNote($"폴더를 열지 못했습니다: {ex.Message}", warning: true); }
    }

    private void ApproveAgent_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        try
        {
            bool inPlace = InPlaceSession;
            var result = _agentWorkflow.Approve();
            AddNote(inPlace ? "작업을 마쳤습니다. 저장은 PowerPoint에서 해 주세요." : $"결과를 저장했습니다: {result}");
        }
        catch (Exception ex) { AddNote($"저장 실패: {ex.Message} · PowerPoint에서 작업 사본을 닫고 다시 시도해 주세요.", warning: true); }
        NotifySession();
    }

    private void OpenWorkingCopy_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        if (_agentWorkflow.Session?.WorkingPath is not { } path) return;
        try { _pptService.OpenPptxInPowerPoint(path); }
        catch (Exception ex) { AddNote($"열기 실패: {ex.Message}", warning: true); }
    }

    private void DiscardAgent_Click(object sender, RoutedEventArgs e)
    {
        CloseMenu();
        bool inPlace = InPlaceSession;
        _agentWorkflow.Discard();
        // 바로 작업은 아무것도 지우지 않는다. 추가된 쪽은 사용자가 PowerPoint에서 지운다.
        AddNote(inPlace ? "대화를 끝냈습니다. 추가된 쪽은 PowerPoint에 그대로 있습니다." : "작업 사본을 버렸습니다. 원본은 그대로입니다.");
        NotifySession();
    }

    private void NotifySession()
    {
        Notify(nameof(HasSession)); Notify(nameof(ApproveLabel)); Notify(nameof(DiscardLabel)); Notify(nameof(ShowOpenWorkingCopy));
    }

    // ══ 대화 영역 ═════════════════════════════════════════════

    /// <summary>사용자 요청이 아닌 알림 한 줄(원문 첨부 완료 등).</summary>
    private void AddNote(string text, bool warning = false)
    {
        var turn = new ConsoleTurn("", system: true);
        turn.AddOutcome(new NoteEntry(text, warning));
        turn.Finish(0);
        Turns.Add(turn);
        if (_followConversation) ScrollToEndSoon();
    }

    private void SetStatus(string value) { _statusDetail = value; RefreshLight(); }

    private void ConversationScroll_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        bool atEnd = ConversationScroll.VerticalOffset >= ConversationScroll.ScrollableHeight - 24;
        // 내용이 늘어서 생긴 변화가 아니라 사용자가 스크롤한 경우에만 따라가기 여부를 바꾼다.
        if (e.ExtentHeightChange == 0) _followConversation = atEnd;
        else if (_followConversation) ConversationScroll.ScrollToEnd();
        ShowJumpToEnd = !_followConversation && !atEnd;
    }

    private void JumpToEnd_Click(object sender, RoutedEventArgs e) => ScrollToEndNow();

    private void ScrollToEndNow()
    {
        _followConversation = true;
        ShowJumpToEnd = false;
        ScrollToEndSoon();
    }

    private void ScrollToEndSoon() =>
        Dispatcher.BeginInvoke(() => ConversationScroll?.ScrollToEnd(), DispatcherPriority.Background);

    // ══ 파일 ═════════════════════════════════════════════════

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
            Activate();
            FocusInput();
        }
        catch (Exception ex) { AddNote($"파일 열기 실패: {ex.Message}", warning: true); }
    }

    private static bool HasPresentationFile(IDataObject data) =>
        (data.GetData(DataFormats.FileDrop) as string[])?.Any(IsPresentationFile) == true;

    private static bool IsPresentationFile(string path) =>
        new[] { ".pptx", ".pptm", ".ppt" }.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    // ══ 시작 확인 ═════════════════════════════════════════════

    private async Task TestLlmConnectionAsync()
    {
        try { await _llmClient.TestConnectionAsync(); SetStatus("LLM 연결됨"); }
        catch (Exception ex) { AddNote($"LLM에 연결하지 못했습니다: {ex.Message}", warning: true); }
    }

    private async Task CheckForUpdatesAsync()
    {
        try
        {
            var current = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0);
            var result = await _updateService.CheckAsync(current);
            if (result.IsUpdateAvailable) AddNote(result.Message);
        }
        catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"업데이트 확인 실패: {ex.Message}"); }
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

    // ══ 창 ════════════════════════════════════════════════════

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

    /// <summary>제목줄을 앱 배경(#1E1E1E)에 맞춘 어두운 색으로. Windows 10은 어두운 모드만, 11은 색까지 맞춘다.</summary>
    private void UseDarkTitleBar()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        int dark = 1;
        DwmSetWindowAttribute(hwnd, 20, ref dark, sizeof(int));          // DWMWA_USE_IMMERSIVE_DARK_MODE
        int caption = 0x001E1E1E;                                          // COLORREF(0x00BBGGRR)
        DwmSetWindowAttribute(hwnd, 35, ref caption, sizeof(int));       // DWMWA_CAPTION_COLOR (Windows 11)
        int text = 0x00D4D4D4;
        DwmSetWindowAttribute(hwnd, 36, ref text, sizeof(int));          // DWMWA_TEXT_COLOR
    }

    private void RestoreBounds()
    {
        Topmost = PptRules.Current.ChatTopmost;
        if (PptRules.Current.ChatWindowBounds?.Split(',') is not { Length: 4 } parts) return;
        var v = parts.Select(p => double.TryParse(p, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : double.NaN).ToArray();
        if (v.Any(double.IsNaN) || v[2] < MinWidth || v[3] < MinHeight) return;
        // 화면 밖이면(모니터 구성이 바뀜) 기본 위치를 쓴다.
        var area = new Rect(SystemParameters.VirtualScreenLeft, SystemParameters.VirtualScreenTop,
            SystemParameters.VirtualScreenWidth, SystemParameters.VirtualScreenHeight);
        if (!area.IntersectsWith(new Rect(v[0], v[1], v[2], v[3]))) return;
        WindowStartupLocation = WindowStartupLocation.Manual;
        (Left, Top, Width, Height) = (v[0], v[1], v[2], v[3]);
    }

    private void SaveBounds()
    {
        if (WindowState != WindowState.Normal) return;
        PptRules.Current.ChatWindowBounds = string.Join(",", new[] { Left, Top, Width, Height }.Select(d => d.ToString("0", CultureInfo.InvariantCulture)));
        try { PptRules.Current.Save(); } catch { }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void Notify([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
