using System.Collections.ObjectModel;
using System.Diagnostics;
using CommunityToolkit.Mvvm.ComponentModel;

namespace BckPPT.Views;

/// <summary>콘솔 한 줄. 종류마다 창의 DataTemplate이 모양을 정한다.</summary>
public abstract class ConsoleEntry : ObservableObject;

/// <summary>모델이 도구를 부르며 쓴 짧은 설명(무엇을 하려는지).</summary>
public sealed class NarrationEntry(string text) : ConsoleEntry
{
    public string Text { get; } = text;
}

/// <summary>회색 안내 한 줄(바로 작업 안내·남은 위반·사용량 등).</summary>
public sealed class NoteEntry(string text, bool warning = false) : ConsoleEntry
{
    public string Text { get; } = text;
    public bool Warning { get; } = warning;
}

public enum StepState { Running, Done, Failed }

/// <summary>
/// 도구 한 번. "› 빈 줄 정리 · 38–41쪽" 다음 줄에 결과("✓ 24줄 삭제")가 붙는다.
/// 실행 중에는 도는 표시와 경과 시간, 안쪽 진행("문안 작성 2/4")이 갱신된다.
/// </summary>
public sealed partial class ToolStepEntry : ConsoleEntry
{
    private static readonly string[] Spinner = ["◐", "◓", "◑", "◒"];
    private readonly Stopwatch _watch = Stopwatch.StartNew();
    private int _frame;

    public ToolStepEntry(string label, string detail, string toolName)
    {
        Label = label;
        Detail = detail;
        ToolName = toolName;
    }

    public string Label { get; }
    public string Detail { get; }
    public string ToolName { get; }
    public string Title => Detail.Length == 0 ? Label : $"{Label}  {Detail}";

    [ObservableProperty] private StepState _state = StepState.Running;
    [ObservableProperty] private string _result = "";
    [ObservableProperty] private string _icon = Spinner[0];

    public bool IsRunning => State == StepState.Running;
    public bool IsFailed => State == StepState.Failed;
    public TimeSpan Elapsed => _watch.Elapsed;

    /// <summary>실행 중일 때 0.25초마다 부른다. 도는 표시와 경과 시간을 갱신한다.</summary>
    public void Tick()
    {
        if (State != StepState.Running) return;
        Icon = Spinner[++_frame % Spinner.Length];
        if (_watch.Elapsed.TotalSeconds >= 2 && !_hasProgress) Result = $"{_watch.Elapsed.TotalSeconds:0}초";
    }

    private bool _hasProgress;

    /// <summary>도구 안쪽 진행(카드 문안 작성 2/4 등).</summary>
    public void Progress(string text)
    {
        _hasProgress = true;
        Result = $"{text} · {_watch.Elapsed.TotalSeconds:0}초";
    }

    public void Complete(string result, bool failed)
    {
        _watch.Stop();
        State = failed ? StepState.Failed : StepState.Done;
        Icon = failed ? "!" : "✓";
        Result = result;
        OnPropertyChanged(nameof(IsRunning));
        OnPropertyChanged(nameof(IsFailed));
    }
}

/// <summary>모델의 최종 답.</summary>
public sealed class AnswerEntry(string text) : ConsoleEntry
{
    public string Text { get; } = text;
}

/// <summary>바뀐 쪽 축소판. 누르면 PowerPoint가 그 쪽으로 간다.</summary>
public sealed record ThumbnailItem(string ImagePath, int? Slide)
{
    public string Caption => Slide is int s ? $"{s}" : "";
}

/// <summary>
/// 요청 하나(사용자 요청 → 작업 흐름 → 답·실행 내역·축소판). 끝나면 작업 흐름을 접고
/// "▸ 도구 5회 · 4쪽 수정 · 12.3초"만 보인다. 진행 중인 요청은 늘 펼쳐 둔다.
/// </summary>
public sealed partial class ConsoleTurn : ObservableObject
{
    private readonly Stopwatch _watch = Stopwatch.StartNew();

    public ConsoleTurn(string request, bool system = false)
    {
        Request = request;
        IsSystem = system;
    }

    public string Request { get; }
    /// <summary>사용자 요청이 아닌 알림(원문 첨부 완료 등). "나" 머리를 붙이지 않는다.</summary>
    public bool IsSystem { get; }
    public bool HasRequest => !IsSystem;

    public ObservableCollection<ConsoleEntry> Activity { get; } = new();
    public ObservableCollection<ConsoleEntry> Outcome { get; } = new();
    public ObservableCollection<string> Changes { get; } = new();
    public ObservableCollection<ThumbnailItem> Thumbnails { get; } = new();

    [ObservableProperty] private bool _isRunning = true;
    [ObservableProperty] private bool _activityExpanded = true;
    [ObservableProperty] private bool _changesExpanded;
    [ObservableProperty] private string _summary = "";

    public bool HasActivity => Activity.Count > 0;
    public bool ShowAgentHeader => HasActivity || Outcome.Count > 0;
    public bool ShowSummaryToggle => !IsRunning && Activity.Count > 0;
    public bool ShowActivity => IsRunning || ActivityExpanded;
    public bool HasChanges => Changes.Count > 0;
    public bool HasThumbnails => Thumbnails.Count > 0;
    public string ChangesHeader => $"실행 내역 {Changes.Count}건  {(ChangesExpanded ? "(접기)" : "(펼치기)")}";
    public string SummaryHeader => $"{Summary}  {(ActivityExpanded ? "(접기)" : "(펼치기)")}";

    public ToolStepEntry? RunningStep => Activity.OfType<ToolStepEntry>().LastOrDefault(s => s.IsRunning);

    public void Add(ConsoleEntry entry)
    {
        Activity.Add(entry);
        OnPropertyChanged(nameof(HasActivity));
        OnPropertyChanged(nameof(ShowAgentHeader));
    }

    public void AddOutcome(ConsoleEntry entry)
    {
        Outcome.Add(entry);
        OnPropertyChanged(nameof(ShowAgentHeader));
    }

    /// <summary>요청이 끝났다. 작업 흐름을 접고 요약을 만든다.</summary>
    public void Finish(int modifiedSlides)
    {
        _watch.Stop();
        int tools = Activity.OfType<ToolStepEntry>().Count();
        int failed = Activity.OfType<ToolStepEntry>().Count(s => s.IsFailed);
        var parts = new List<string> { $"도구 {tools}회" };
        if (failed > 0) parts.Add($"실패 {failed}");
        if (modifiedSlides > 0) parts.Add($"{modifiedSlides}쪽 수정");
        parts.Add($"{_watch.Elapsed.TotalSeconds:0.0}초");
        Summary = string.Join(" · ", parts);
        IsRunning = false;
        ActivityExpanded = false;
        OnPropertyChanged(nameof(HasChanges));
        OnPropertyChanged(nameof(HasThumbnails));
    }

    public void ToggleActivity() => ActivityExpanded = !ActivityExpanded;
    public void ToggleChanges() => ChangesExpanded = !ChangesExpanded;

    partial void OnIsRunningChanged(bool value) { OnPropertyChanged(nameof(ShowSummaryToggle)); OnPropertyChanged(nameof(ShowActivity)); }
    partial void OnActivityExpandedChanged(bool value) { OnPropertyChanged(nameof(ShowActivity)); OnPropertyChanged(nameof(SummaryHeader)); }
    partial void OnChangesExpandedChanged(bool value) => OnPropertyChanged(nameof(ChangesHeader));
    partial void OnSummaryChanged(string value) => OnPropertyChanged(nameof(SummaryHeader));
}
