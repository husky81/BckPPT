using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.ViewModels;

public partial class BckPptViewModel : ObservableObject
{
    private readonly IPptComService _pptComService;
    private readonly IUiDispatcher  _dispatcher;

    private IReadOnlyList<PptSlideInfo>  _allShapes = [];
    private IReadOnlyList<PptFontInfo>   _allFonts  = [];
    private IReadOnlyList<SlideListItem> _allSlides = [];
    private ObservableCollection<PptNode> _allRootNodes = [];

    // ── 파일 / 슬라이드 ──────────────────────────────────────────

    [ObservableProperty]
    private ObservableCollection<PptPresentation> _presentations = [];

    [ObservableProperty]
    private PptPresentation? _selectedPresentation;

    [ObservableProperty]
    private ObservableCollection<SlideListItem> _slideItems = [];

    [ObservableProperty]
    private SlideListItem? _selectedSlide;

    // ── Shape 목록 ───────────────────────────────────────────────

    [ObservableProperty]
    private ObservableCollection<PptSlideInfo> _currentShapes = [];

    [ObservableProperty]
    private PptSlideInfo? _selectedShape;

    // ── 트리 구조 ────────────────────────────────────────────────

    [ObservableProperty]
    private ObservableCollection<PptNode> _rootNodes = [];

    // ── 글꼴 필터 ────────────────────────────────────────────────

    [ObservableProperty]
    private ObservableCollection<string> _availableFonts = [];

    [ObservableProperty]
    private string? _selectedFontFilter;

    // ── 상태 ─────────────────────────────────────────────────────

    [ObservableProperty]
    private ObservableCollection<string> _statusLog = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsNotBusy))]
    private bool _isBusy;

    [ObservableProperty]
    private bool _isAnalyzing;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOverflows))]
    private int _totalOverflowCount;

    public bool IsNotBusy    => !IsBusy;
    public bool HasOverflows => TotalOverflowCount > 0;

    // ── 스트리밍 분석 ────────────────────────────────────────────

    private CancellationTokenSource? _analyzeCts;
    private readonly Dictionary<int, int> _skeletonIndex = [];
    private readonly List<PptSlideInfo>   _streamingShapes = [];
    private readonly List<PptFontInfo>    _streamingFonts  = [];

    [ObservableProperty]
    private string? _analysisProgressText;

    // ── Undo ─────────────────────────────────────────────────────

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(UndoCommand))]
    private bool _canUndo;

    private List<(int SlideIndex, string ShapeName, string OldText)>? _undoSnapshot;

    // ── 표시 ─────────────────────────────────────────────────────

    public System.Windows.Visibility EmptyTableVisibility =>
        CurrentShapes.Count == 0
            ? System.Windows.Visibility.Visible
            : System.Windows.Visibility.Collapsed;

    public System.Windows.Visibility SlideColumnVisibility =>
        SelectedSlide is not null && !SelectedSlide.IsAll
            ? System.Windows.Visibility.Collapsed
            : System.Windows.Visibility.Visible;

    // ── 생성자 ───────────────────────────────────────────────────

    public BckPptViewModel(IPptComService pptComService, IUiDispatcher dispatcher)
    {
        _pptComService = pptComService;
        _dispatcher    = dispatcher;
    }

    // ── 속성 변경 핸들러 ─────────────────────────────────────────

    partial void OnSelectedPresentationChanged(PptPresentation? value)
    {
        if (value is not null) _ = AnalyzeAsync(value.Path);
    }

    partial void OnSelectedSlideChanged(SlideListItem? value)
    {
        // 분석 중에는 RebuildSlideList() 호출을 막는다.
        // ApplyFilters() → RebuildSlideList() 가 SlideItems를 교체하면
        // _skeletonIndex가 가리키는 인덱스가 무효화되어 IndexOutOfRange 발생.
        if (!IsAnalyzing) ApplyFilters();
    }
    partial void OnSelectedFontFilterChanged(string? value)
    {
        if (!IsAnalyzing) ApplyFilters();
    }

    // ── 새로고침 ─────────────────────────────────────────────────

    [RelayCommand]
    private void Refresh()
    {
        IsBusy = true;
        try
        {
            var openList = _pptComService.GetOpenPresentations();
            Presentations = new ObservableCollection<PptPresentation>(openList);

            if (openList.Count > 0)
            {
                AddStatus($"✔ 열린 프레젠테이션 {openList.Count}개");
                SelectedPresentation = openList[0];
            }
            else
            {
                AddStatus("PowerPoint에 열린 파일 없음");
                SlideItems = [];
                ClearResults();
            }
        }
        catch (Exception ex) { AddStatus($"✖ 오류: {ex.Message}"); }
        finally { IsBusy = false; }
    }

    // ── 분석 ─────────────────────────────────────────────────────

    private async Task AnalyzeAsync(string path)
    {
        // 이전 분석 취소 (다른 파일 선택 시)
        _analyzeCts?.Cancel();
        _analyzeCts?.Dispose();
        var cts = new CancellationTokenSource();
        _analyzeCts = cts;

        IsAnalyzing = true;
        SelectedShape = null;
        CurrentShapes = [];
        _skeletonIndex.Clear();
        _streamingShapes.Clear();
        _streamingFonts.Clear();

        // "전체" skeleton만 먼저 표시
        SlideItems = [new SlideListItem(0, "", 0, 0)];
        AnalysisProgressText = "슬라이드 분석 중...";
        AddStatus("🔍 슬라이드 분석 중...");

        // 슬라이드/shape 완료마다 UI 갱신 (Progress<T> → UI 스레드 보장)
        var progress = new Progress<SlideAnalysisProgress>(p =>
        {
            if (cts.IsCancellationRequested) return;

            switch (p.Kind)
            {
                case AnalysisProgressKind.SlideStart:
                    _skeletonIndex[p.SlideIndex] = SlideItems.Count;
                    SlideItems.Add(new SlideListItem(
                        p.SlideIndex, $"슬라이드 {p.SlideIndex}", 0, 0,
                        IsAnalyzing: true));
                    AnalysisProgressText = $"슬라이드 {p.SlideIndex}/{p.TotalSlides} 분석 중...";
                    break;

                case AnalysisProgressKind.ShapeBatch:
                    if (p.ShapeBatch is not null) _streamingShapes.AddRange(p.ShapeBatch);
                    if (p.FontBatch  is not null) _streamingFonts.AddRange(p.FontBatch);
                    // "전체" 항목 누적 카운트 갱신 (SlideItems가 비어있는 경우 방어)
                    if (SlideItems.Count > 0)
                    {
                        SlideItems[0] = SlideItems[0] with
                        {
                            ShapeCount    = _streamingShapes.Count,
                            OverflowCount = _streamingShapes.Count(s => s.IsOverflow),
                        };
                    }
                    break;

                case AnalysisProgressKind.SlideCompleted:
                    if (p.Slide is not null &&
                        _skeletonIndex.TryGetValue(p.SlideIndex, out int idx) &&
                        idx < SlideItems.Count)  // skeleton이 교체되지 않았을 때만
                    {
                        var prev = SelectedSlide;
                        SlideItems[idx] = p.Slide;
                        // skeleton 선택 중이었으면 확정 항목으로 복원
                        if (prev?.SlideIndex == p.SlideIndex)
                            SelectedSlide = p.Slide;
                    }
                    break;
            }
        });

        try
        {
            var result = await Task.Run(
                () => _pptComService.AnalyzePresentation(path, progress, cts.Token),
                cts.Token);

            if (cts.IsCancellationRequested) return;

            _allShapes = result.Shapes;
            _allFonts  = result.Fonts;
            _allSlides = result.Slides;

            // 역할 분류
            var prsName = Path.GetFileName(path);
            var userLabels = ShapeRoleLabelStore.LoadAll();
            var roles = ShapeRoleClassifier.Classify(result, userLabels, prsName);
            foreach (var shape in _allShapes)
            {
                var key = (shape.SlideIndex, shape.ShapeName);
                if (roles.TryGetValue(key, out var role))
                {
                    shape.Role = role;
                    shape.IsUserLabeled = userLabels.ContainsKey((prsName, shape.SlideIndex, shape.ShapeName));
                }
            }

            int totalOverflow = result.Shapes.Count(s => s.IsOverflow);
            TotalOverflowCount = totalOverflow;

            var fontNames = _allFonts
                .Select(f => f.FontName).Distinct().OrderBy(n => n).ToList();
            fontNames.Insert(0, "");
            AvailableFonts     = new ObservableCollection<string>(fontNames);
            SelectedFontFilter = "";

            // 트리 구조 복원 또는 초기화 (빈 상자 제외, 중복 shapeName 제거)
            var nonEmpty = _allShapes.Where(s => !s.IsEmpty).ToList();
            _allRootNodes = TreeStructureStore.Load(prsName, nonEmpty)
                ?? new ObservableCollection<PptNode>(
                    nonEmpty
                        .GroupBy(s => $"{s.SlideIndex}:{s.ShapeName}")
                        .Select(g => new PptNode(g.First())));
            RootNodes = _allRootNodes;

            // 최종 목록 재구성 (필터 반영, 카운트 확정)
            RebuildSlideList();
            SelectedSlide = SlideItems.FirstOrDefault();

            AddStatus($"✔ 분석 완료 · 슬라이드 {result.Slides.Count}개, 글꼴 {fontNames.Count}종, overflow {totalOverflow}개");
        }
        catch (OperationCanceledException) { /* 정상 취소 */ }
        catch (Exception ex) { AddStatus($"✖ 분석 오류: {ex.Message}"); }
        finally
        {
            IsAnalyzing = false;
            AnalysisProgressText = null;
        }
    }

    // ── 필터 ─────────────────────────────────────────────────────

    private void ApplyFilters()
    {
        RebuildSlideList();

        IEnumerable<PptSlideInfo> source = _allShapes;

        bool hasFontFilter = !string.IsNullOrEmpty(SelectedFontFilter);
        if (!hasFontFilter)
            source = source.Where(s => !s.IsEmpty);

        if (SelectedSlide is not null && !SelectedSlide.IsAll)
            source = source.Where(s => s.SlideIndex == SelectedSlide.SlideIndex);

        if (hasFontFilter)
        {
            var matching = GetFontMatchingKeys(SelectedFontFilter!);
            source = source.Where(s => matching.Contains((s.SlideIndex, s.ShapeName)));
        }

        CurrentShapes = new ObservableCollection<PptSlideInfo>(source);

        // 트리뷰는 항상 _allRootNodes를 직접 사용 (드래그 조작이 바로 반영되도록)
        RootNodes = _allRootNodes;

        // 슬라이드 필터: TreeViewItem Visibility로 처리
        int? filterSlide = (SelectedSlide is not null && !SelectedSlide.IsAll)
            ? SelectedSlide.SlideIndex : null;
        ApplySlideFilter(_allRootNodes, filterSlide);

        OnPropertyChanged(nameof(EmptyTableVisibility));
        OnPropertyChanged(nameof(SlideColumnVisibility));
    }

    private void RebuildSlideList()
    {
        IEnumerable<SlideListItem> slides = _allSlides;

        if (!string.IsNullOrEmpty(SelectedFontFilter))
        {
            var slideIndices = _allFonts
                .Where(f => f.FontName == SelectedFontFilter)
                .Select(f => f.SlideIndex)
                .ToHashSet();
            slides = slides.Where(s => slideIndices.Contains(s.SlideIndex));
        }

        var items = new List<SlideListItem>
        {
            new(0, "", _allShapes.Count, _allShapes.Count(s => s.IsOverflow))
        };
        items.AddRange(slides);

        var prev = SelectedSlide;
        SlideItems  = new ObservableCollection<SlideListItem>(items);
        SelectedSlide = prev is not null
            ? items.FirstOrDefault(s => s.SlideIndex == prev.SlideIndex) ?? items[0]
            : items[0];
    }

    private HashSet<(int, string)> GetFontMatchingKeys(string fontName)
        => _allFonts
            .Where(f => f.FontName == fontName)
            .Select(f => (f.SlideIndex, f.ShapeName))
            .ToHashSet();

    // ── 트리 필터 ─────────────────────────────────────────────────

    private static void ApplySlideFilter(IEnumerable<PptNode> nodes, int? slideIndex)
    {
        foreach (var node in nodes)
        {
            node.IsVisible = slideIndex is null || node.Shape.SlideIndex == slideIndex.Value;
            ApplySlideFilter(node.Children, slideIndex);
        }
    }

    // ── 트리 구조 조작 ─────────────────────────────────────────────

    public enum DropPosition { Above, Inside, Below }

    /// <summary>노드를 target 기준으로 이동. UI/API 공용.</summary>
    public void MoveNode(PptNode source, PptNode target, DropPosition position)
    {
        if (source == target || source.IsDescendant(target)) return;

        // 기존 부모에서 제거
        if (source.Parent is not null)
            source.Parent.Children.Remove(source);
        else
            _allRootNodes.Remove(source);

        if (position == DropPosition.Inside)
        {
            target.Children.Add(source);
            source.Parent = target;
        }
        else
        {
            var parentList = target.Parent?.Children ?? _allRootNodes;
            int idx = parentList.IndexOf(target);
            if (position == DropPosition.Below) idx++;
            if (idx > parentList.Count) idx = parentList.Count;
            parentList.Insert(idx, source);
            source.Parent = target.Parent;
        }
    }

    /// <summary>nodeId (SlideIndex:ShapeName) 로 노드 검색.</summary>
    public PptNode? FindNode(string nodeId)
    {
        static PptNode? Search(IEnumerable<PptNode> nodes, string id)
        {
            foreach (var n in nodes)
            {
                if (n.NodeId == id) return n;
                var found = Search(n.Children, id);
                if (found is not null) return found;
            }
            return null;
        }
        return Search(_allRootNodes, nodeId);
    }

    public void SaveTreeStructure()
    {
        if (SelectedPresentation is null) return;
        var prsName = Path.GetFileName(SelectedPresentation.Path);
        TreeStructureStore.Save(prsName, _allRootNodes);
    }

    // ── 텍스트 치환용 데이터 ─────────────────────────────────────

    public List<(int Seq, int SlideIndex, string ShapeName, string FullText)> GetTextsForCurrentView()
        => CurrentShapes
            .Where(s => !s.IsEmpty)
            .Select((s, i) => (i + 1, s.SlideIndex, s.ShapeName, s.FullText))
            .ToList();

    /// <summary>텍스트 치환 적용. 적용 전 스냅샷을 저장해 Undo 가능하게 함.</summary>
    public async Task ApplyTextChangesAsync(
        List<(int SlideIndex, string ShapeName, string OldText, string NewText)> changes)
    {
        if (SelectedPresentation is null || changes.Count == 0) return;
        IsBusy = true;
        try
        {
            // Undo 스냅샷 저장
            _undoSnapshot = changes
                .Select(c => (c.SlideIndex, c.ShapeName, c.OldText))
                .ToList();
            CanUndo = true;

            string path = SelectedPresentation.Path;
            string name = SelectedPresentation.Name;
            int count = await Task.Run(() =>
            {
                foreach (var (slideIdx, shapeName, _, newText) in changes)
                    _pptComService.SetShapeText(path, slideIdx, shapeName, newText);
                return changes.Count;
            });

            // 이력 기록
            var histEntries = changes.Select(c => new ChangeHistoryEntry(
                DateTime.Now, name, "TextReplace",
                c.SlideIndex, c.ShapeName,
                c.OldText.Replace("\r", "\\n"),
                c.NewText.Replace("\r", "\\n")));
            ChangeHistoryService.AppendMany(histEntries);

            AddStatus($"✔ 텍스트 치환 {count}건 적용 (되돌리기 가능)");
            await AnalyzeAsync(path);
        }
        catch (Exception ex) { AddStatus($"✖ 치환 오류: {ex.Message}"); }
        finally { IsBusy = false; }
    }

    // ── 단일 Shape 새로고침 ──────────────────────────────────────

    public async Task RefreshShapeAsync(PptSlideInfo shape)
    {
        if (SelectedPresentation is null) return;
        try
        {
            string path    = SelectedPresentation.Path;
            var    updated = await Task.Run(() =>
                _pptComService.GetShapeInfo(path, shape.SlideIndex, shape.ShapeName));
            if (updated is null) return;

            // _allShapes 에서 해당 항목 교체
            var list = _allShapes.ToList();
            int idx  = list.FindIndex(s =>
                s.SlideIndex == shape.SlideIndex && s.ShapeName == shape.ShapeName);
            if (idx >= 0) list[idx] = updated;
            _allShapes = list;

            // CurrentShapes에서 동일 항목 교체
            int ci = CurrentShapes.IndexOf(shape);
            if (ci >= 0)
                CurrentShapes[ci] = updated;

            AddStatus($"✔ [{shape.ShapeName}] 새로고침");
        }
        catch (Exception ex) { AddStatus($"✖ 새로고침 오류: {ex.Message}"); }
    }

    // ── Undo ─────────────────────────────────────────────────────

    [RelayCommand(CanExecute = nameof(CanUndo))]
    private async Task Undo()
    {
        if (_undoSnapshot is null || SelectedPresentation is null) return;
        IsBusy = true;
        try
        {
            string path     = SelectedPresentation.Path;
            var    snapshot = _undoSnapshot;
            int count = await Task.Run(() =>
            {
                foreach (var (slideIdx, shapeName, oldText) in snapshot)
                    _pptComService.SetShapeText(path, slideIdx, shapeName, oldText);
                return snapshot.Count;
            });
            _undoSnapshot = null;
            CanUndo       = false;
            AddStatus($"↩ 되돌리기 완료 ({count}건)");
            await AnalyzeAsync(path);
        }
        catch (Exception ex) { AddStatus($"✖ 되돌리기 오류: {ex.Message}"); }
        finally { IsBusy = false; }
    }

    // ── AutoFit ──────────────────────────────────────────────────

    [RelayCommand]
    private async Task AutoFitAll()
    {
        if (SelectedPresentation is null) return;
        IsBusy = true;
        try
        {
            var progress = new Progress<string>(msg => AddStatus($"  {msg}"));
            var result = await _pptComService.AutoFitAsync(SelectedPresentation.Path, progress);

            // 이력 기록
            if (result.TotalFixed > 0)
            {
                var histEntries = result.Details.Select(d => new ChangeHistoryEntry(
                    DateTime.Now, SelectedPresentation.Name, "AutoFit",
                    d.SlideIndex, d.ShapeName,
                    $"{d.OriginalFontSize}pt",
                    $"{d.FinalFontSize}pt{(d.WasTrimmed ? " (잘림)" : "")}"));
                ChangeHistoryService.AppendMany(histEntries);
            }

            // 상세 리포트 — 수정된 항목 목록
            var sb = new StringBuilder();
            sb.Append($"✔ AutoFit 완료 · {result.TotalFixed}개 수정");
            if (result.TotalFixed > 0)
            {
                sb.AppendLine();
                foreach (var d in result.Details.Take(5))
                {
                    sb.Append($"  S{d.SlideIndex} {d.ShapeName}: " +
                              $"{d.OriginalFontSize}→{d.FinalFontSize}pt");
                    if (d.WasTrimmed) sb.Append(" [잘림]");
                    sb.AppendLine();
                }
                if (result.Details.Count > 5)
                    sb.Append($"  … 외 {result.Details.Count - 5}건");
            }
            AddStatus(sb.ToString().TrimEnd());
            await AnalyzeAsync(SelectedPresentation.Path);
        }
        catch (Exception ex) { AddStatus($"✖ AutoFit 오류: {ex.Message}"); }
        finally { IsBusy = false; }
    }

    // ── 규칙 일괄 치환 ────────────────────────────────────────────

    [RelayCommand]
    private async Task ApplyRules()
    {
        if (SelectedPresentation is null) return;
        var rules = PptRules.Current.ReplaceRules;
        if (rules.Count == 0)
        {
            AddStatus("치환 규칙이 없습니다. 규칙 편집기에서 추가하세요.");
            return;
        }

        IsBusy = true;
        try
        {
            string path = SelectedPresentation.Path;
            string name = SelectedPresentation.Name;
            var changed = await Task.Run(() =>
                _pptComService.ApplyRules(path, rules));

            if (changed.Count > 0)
            {
                var histEntries = changed.Select(c => new ChangeHistoryEntry(
                    DateTime.Now, name, "RuleReplace",
                    c.SlideIndex, c.ShapeName,
                    c.OldText.Replace("\r", "\\n"),
                    c.NewText.Replace("\r", "\\n")));
                ChangeHistoryService.AppendMany(histEntries);
            }

            AddStatus($"✔ 규칙 치환 완료 · {changed.Count}개 shape 변경");
            if (changed.Count > 0) await AnalyzeAsync(path);
        }
        catch (Exception ex) { AddStatus($"✖ 규칙 치환 오류: {ex.Message}"); }
        finally { IsBusy = false; }
    }

    // ── CSV 내보내기 ──────────────────────────────────────────────

    [RelayCommand]
    private async Task ExportCsv()
    {
        if (CurrentShapes.Count == 0) return;
        IsBusy = true;
        try
        {
            var shapes = CurrentShapes.ToList();
            string path = await Task.Run(() =>
            {
                string tempFile = Path.Combine(Path.GetTempPath(),
                    $"bckppt_export_{DateTime.Now:yyyyMMdd_HHmmss}.csv");

                var sb = new StringBuilder();
                sb.AppendLine("Slide,Shape,FontSize,HalfWidthLength,BoundH,BoxH,Lines,LineH,Overflow,TextPreview");
                foreach (var s in shapes)
                {
                    var preview = s.TextPreview.Replace("\"", "\"\"");
                    sb.AppendLine(
                        $"{s.SlideIndex},\"{s.ShapeName}\"," +
                        $"{s.FontSize},{s.HalfWidthLength}," +
                        $"{s.BoundHeight},{s.BoxHeight}," +
                        $"{s.LineCount},{s.LineHeight}," +
                        $"{(s.IsOverflow ? "Y" : "")}," +
                        $"\"{preview}\"");
                }
                File.WriteAllText(tempFile, sb.ToString(), System.Text.Encoding.UTF8);
                return tempFile;
            });

            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
            AddStatus($"✔ CSV 내보내기 완료 · {shapes.Count}행");
        }
        catch (Exception ex) { AddStatus($"✖ 내보내기 오류: {ex.Message}"); }
        finally { IsBusy = false; }
    }

    // ── Navigate ─────────────────────────────────────────────────

    [RelayCommand]
    private async Task NavigateToShape(PptSlideInfo? info)
    {
        if (info is null || SelectedPresentation is null) return;
        try
        {
            await _pptComService.NavigateToShapeAsync(
                SelectedPresentation.Path, info.SlideIndex, info.ShapeName);
        }
        catch (Exception ex) { AddStatus($"✖ 네비게이션 오류: {ex.Message}"); }
    }

    // ── 상태 로그 ────────────────────────────────────────────────

    public void AddStatus(string message)
    {
        var entry = $"{DateTime.Now:HH:mm:ss}  {message}";
        _dispatcher.Post(() =>
        {
            StatusLog.Insert(0, entry);
            while (StatusLog.Count > 8)
                StatusLog.RemoveAt(StatusLog.Count - 1);
        });
    }

    private void ClearResults()
    {
        _allShapes = [];
        _allFonts  = [];
        _allSlides = [];
        CurrentShapes      = [];
        AvailableFonts     = [];
        TotalOverflowCount = 0;
        _undoSnapshot      = null;
        CanUndo            = false;
    }
}
