using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using CellBox2.BckPPT.Models;
using CellBox2.BckPPT.Services;
using CellBox2.BckPPT.ViewModels;

namespace CellBox2.BckPPT.Views;

/// <summary>WPF 디스패처 구현 — ViewModel이 Application.Current에 직접 의존하지 않도록.</summary>
internal class WpfDispatcher : IUiDispatcher
{
    public void Post(Action action)
        => Application.Current?.Dispatcher.BeginInvoke(action);
}

public partial class BckPptWindow : Window
{
    public BckPptViewModel ViewModel { get; }

    // ── Debug API ──────────────────────────────────────────────────
    private DebugApiServer? _debugApi;

    // ── Shape 선택 debounce 상태 ──────────────────────────────────
    private CancellationTokenSource? _navCts;
    private (int SlideIndex, string ShapeName)? _lastNavigated;

    // ── 다중선택 상태 ────────────────────────────────────────────
    private PptNode? _selectionAnchor;
    private PptNode? _pendingSelectNode; // MouseUp에서 선택 적용용
    private bool _didDrag;

    // ── Drag & Drop 상태 ─────────────────────────────────────────
    private Point _dragStartPoint;
    private DropGuideAdorner? _dropAdorner;
    private bool _isMovingNode; // 노드 이동 중 네비게이션 억제
    private System.Windows.Documents.AdornerLayer? _adornerLayer;

    public BckPptWindow()
    {
        var pptComService = new PptComService();
        ViewModel = new BckPptViewModel(pptComService, new WpfDispatcher());
        DataContext = ViewModel;

        InitializeComponent();

        SourceInitialized += (_, _) =>
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            int value = 1;
            DwmSetWindowAttribute(hwnd, 20, ref value, sizeof(int));
        };

        var version = System.Reflection.Assembly.GetExecutingAssembly()
            .GetName().Version;
        Title = $"BckPPT v{version?.ToString(3) ?? "0.0.0"}";

        Loaded += (_, _) =>
        {
            ViewModel.RefreshCommand.Execute(null);
            _debugApi = new DebugApiServer(ViewModel, Dispatcher);
            _debugApi.Start();
        };
        Closing += (_, _) =>
        {
            _debugApi?.Stop();
            ViewModel.SaveTreeStructure();
        };
    }

    // ── 슬라이드 선택 → PPT 이동 ─────────────────────────────────

    private async void SlideList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ViewModel.SelectedSlide is not { } slide || slide.IsAll) return;
        if (ViewModel.SelectedPresentation is null) return;
        try
        {
            await Task.Run(() => new PptComService().NavigateToSlide(
                ViewModel.SelectedPresentation.Path, slide.SlideIndex));
        }
        catch { }
    }

    // ── TreeView 선택 → PPT 이동 (debounce) ───────────────────────

    private static readonly TimeSpan NormalDelay = TimeSpan.FromMilliseconds(300);
    private static readonly TimeSpan CellDelay   = TimeSpan.FromMilliseconds(700);

    private async void ShapeTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (_isMovingNode) return;
        if (e.NewValue is not PptNode node) return;
        if (ViewModel.SelectedPresentation is null) return;

        // 키보드 화살표 네비게이션 동기화:
        // 마우스 클릭은 PreviewMouseLeftButtonDown에서 이미 node.IsSelected = true 처리.
        // 화살표 키는 WPF 내부 SelectedItem만 바뀌므로, IsSelected가 false인 경우에만 동기화.
        if (!node.IsSelected)
        {
            ClearAllSelections();
            node.IsSelected = true;
            _selectionAnchor = node;
        }

        var shape = node.Shape;
        ViewModel.SelectedShape = shape;

        var key = (shape.SlideIndex, shape.ShapeName);
        if (_lastNavigated == key) return;

        _navCts?.Cancel();
        _navCts?.Dispose();
        var cts = new CancellationTokenSource();
        _navCts = cts;

        bool isCell = shape.ShapeName.Contains(" [R", StringComparison.Ordinal);
        var delay = isCell ? CellDelay : NormalDelay;

        try
        {
            await Task.Delay(delay, cts.Token);
            if (cts.IsCancellationRequested) return;
            _lastNavigated = key;
            string path = ViewModel.SelectedPresentation.Path;
            await new PptComService().NavigateToShapeAsync(path, shape.SlideIndex, shape.ShapeName);
            // PPT의 Activate() 호출로 포커스가 빼앗기므로 BckPPT로 즉시 복원.
            // 이렇게 해야 Up/Down 화살표가 TreeView에서 계속 동작한다.
            Activate();
            ShapeTree.Focus();
        }
        catch (OperationCanceledException) { }
        catch { }
    }

    // ── 텍스트 치환 ───────────────────────────────────────────────

    private async void TextReplace_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedPresentation is null) return;

        var entries = ViewModel.GetTextsForCurrentView();
        if (entries.Count == 0) return;

        var slideTitle = ViewModel.SelectedSlide?.DisplayText ?? "전체";
        var win = new TextEditorWindow(
            ViewModel.SelectedPresentation.Path, entries, slideTitle);
        win.Owner = this;

        if (win.ShowDialog() == true && win.AppliedChanges is { Count: > 0 })
        {
            await ViewModel.ApplyTextChangesAsync(win.AppliedChanges);
        }
    }

    // ── 규칙 편집 + 치환 ──────────────────────────────────────────

    private async void RulesReplace_Click(object sender, RoutedEventArgs e)
    {
        if (ViewModel.SelectedPresentation is null) return;

        var rulesWin = new RulesWindow { Owner = this };

        if (PptRules.Current.ReplaceRules.Count == 0)
        {
            if (rulesWin.ShowDialog() != true) return;
            PptRules.Reload();
        }

        var result = MessageBox.Show(
            $"{PptRules.Current.ReplaceRules.Count}개 규칙을 적용합니다.\n" +
            "편집하려면 '아니오'를 누르세요.",
            "규칙 치환",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question);

        if (result == MessageBoxResult.Cancel) return;
        if (result == MessageBoxResult.No)
        {
            rulesWin = new RulesWindow { Owner = this };
            if (rulesWin.ShowDialog() != true) return;
            PptRules.Reload();
        }

        await ViewModel.ApplyRulesCommand.ExecuteAsync(null);
    }

    // ── 레이아웃 보기 ─────────────────────────────────────────────

    private void LayoutMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi) return;
        if (mi.Parent is not ContextMenu cm) return;
        if (cm.PlacementTarget is not FrameworkElement target) return;
        if (target.DataContext is not SlideListItem slide || slide.IsAll) return;
        if (ViewModel.SelectedPresentation is null) return;

        var win = new LayoutWindow(
            ViewModel.SelectedPresentation.Path,
            ViewModel.SlideItems,
            slide.SlideIndex)
        { Owner = this };
        win.Show();
    }

    // ── 새로고침 시 선택 해제 ────────────────────────────────────────

    private void Refresh_Click(object sender, RoutedEventArgs e)
    {
        _lastNavigated = null;
        ViewModel.SelectedShape = null;
    }

    // ── 텍스트 확인 ─────────────────────────────────────────────────

    private async void TextInspect_Click(object sender, RoutedEventArgs e)
    {
        var shape = GetShapeFromMenuItem(sender);
        if (shape is null || shape.IsEmpty) return;
        if (ViewModel.SelectedPresentation is null) return;

        try
        {
            var svc  = new PptComService();
            var path = ViewModel.SelectedPresentation.Path;
            var lines = await Task.Run(() =>
                svc.GetTextLineBreaks(path, shape.SlideIndex, shape.ShapeName));

            if (lines.Count == 0) return;
            var win = new TextInspectWindow(shape.ShapeName, lines) { Owner = this };
            win.Show();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "텍스트 확인 오류",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ── 역할 지정 ──────────────────────────────────────────────────

    private void RoleAssign_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem mi || mi.Tag is not string tagStr) return;
        if (!Enum.TryParse<ShapeRole>(tagStr, out var role)) return;
        if (ViewModel.SelectedPresentation is null) return;

        var prsName = System.IO.Path.GetFileName(ViewModel.SelectedPresentation.Path);

        // 다중선택 지원: 선택된 모든 노드에 일괄 적용
        var selected = GetSelectedNodes();
        if (selected.Count == 0)
        {
            // 선택된 노드 없으면 우클릭 대상 사용
            var node = GetNodeFromMenuItem(sender);
            if (node is null) return;
            selected = [node];
        }

        var labels = new List<ShapeRoleLabel>();
        foreach (var node in selected)
        {
            var s = node.Shape;
            labels.Add(new ShapeRoleLabel(
                prsName, s.SlideIndex, s.ShapeName, role, DateTime.Now,
                s.IsImage ? "Image" : "Text",
                s.FontSize, false, 0f, 0f, 0f, 0f,
                s.LineCount, false, s.HalfWidthLength));
            node.SetRole(role);
        }
        ShapeRoleLabelStore.SaveMany(labels);
        ViewModel.SaveTreeStructure();

        ClearAllSelections();
    }

    // ── Shape 새로고침 ────────────────────────────────────────────

    private async void ShapeRefresh_Click(object sender, RoutedEventArgs e)
    {
        var shape = GetShapeFromMenuItem(sender);
        if (shape is null) return;
        await ViewModel.RefreshShapeAsync(shape);
    }

    // ── Dev Info 복사 ─────────────────────────────────────────────

    private void CopyDevInfo_Click(object sender, RoutedEventArgs e)
    {
        var shape = GetShapeFromMenuItem(sender);
        if (shape is null) return;

        var info = $"Slide:{shape.SlideIndex} Shape:\"{shape.ShapeName}\" " +
                   $"BoundH:{shape.BoundHeight} BoxH:{shape.BoxHeight} " +
                   $"Lines:{shape.LineCount} Overflow:{shape.IsOverflow}";
        Clipboard.SetText(info);
    }

    // ── 이력 보기 ─────────────────────────────────────────────────

    private void History_Click(object sender, RoutedEventArgs e)
    {
        var win = new HistoryWindow { Owner = this };
        win.Show();
    }

    // ══════════════════════════════════════════════════════════════
    //  드래그 앤 드롭 (TreeView)
    // ══════════════════════════════════════════════════════════════

    private void ShapeTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _dragStartPoint = e.GetPosition(null);
        _didDrag = false;
        _pendingSelectNode = null;

        var tvi = GetTreeViewItem(e.OriginalSource as DependencyObject);
        if (tvi?.DataContext is not PptNode node) return;

        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
        {
            // Ctrl+클릭: 즉시 토글
            node.IsSelected = !node.IsSelected;
            if (node.IsSelected) _selectionAnchor = node;
        }
        else if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && _selectionAnchor is not null)
        {
            // Shift+클릭: 즉시 범위 선택
            var flat = FlattenNodes(ViewModel.RootNodes).ToList();
            int anchorIdx = flat.IndexOf(_selectionAnchor);
            int targetIdx = flat.IndexOf(node);
            if (anchorIdx >= 0 && targetIdx >= 0)
            {
                int from = Math.Min(anchorIdx, targetIdx);
                int to   = Math.Max(anchorIdx, targetIdx);
                ClearAllSelections();
                for (int i = from; i <= to; i++)
                    flat[i].IsSelected = true;
            }
        }
        else if (node.IsSelected)
        {
            // 이미 선택된 항목 클릭 → 드래그 대기. MouseUp에서 단일 선택 처리.
            _pendingSelectNode = node;
        }
        else
        {
            // 일반 클릭: 단일 선택
            ClearAllSelections();
            node.IsSelected = true;
            _selectionAnchor = node;
        }
    }

    private void ShapeTree_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        // 드래그 없이 MouseUp → 단일 선택 적용
        if (!_didDrag && _pendingSelectNode is not null)
        {
            ClearAllSelections();
            _pendingSelectNode.IsSelected = true;
            _selectionAnchor = _pendingSelectNode;
        }
        _pendingSelectNode = null;
    }

    private void ShapeTree_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed) return;

        var diff = _dragStartPoint - e.GetPosition(null);
        if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance) return;

        var tvi = GetTreeViewItem(e.OriginalSource as DependencyObject);
        if (tvi?.DataContext is not PptNode node) return;

        // 선택된 노드가 없거나 현재 노드가 선택 안 됐으면 단일 드래그
        var selected = GetSelectedNodes();
        if (selected.Count == 0 || !node.IsSelected)
        {
            ClearAllSelections();
            node.IsSelected = true;
            selected = [node];
        }

        _didDrag = true;
        _pendingSelectNode = null;
        DragDrop.DoDragDrop(tvi, selected, DragDropEffects.Move);
        RemoveDropGuide();
    }

    private void ShapeTree_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = DragDropEffects.None;
        if (!e.Data.GetDataPresent(typeof(List<PptNode>))) { RemoveDropGuide(); return; }

        var tvi = GetTreeViewItem(e.OriginalSource as DependencyObject);
        if (tvi?.DataContext is not PptNode target) { RemoveDropGuide(); return; }

        var draggedList = (List<PptNode>)e.Data.GetData(typeof(List<PptNode>));
        if (draggedList.Any(d => d == target || d.IsDescendant(target))) { RemoveDropGuide(); return; }

        tvi.UpdateLayout();
        var pos = e.GetPosition(tvi);
        var height = tvi.ActualHeight;

        var dropPos = pos.Y < height * 0.33 ? DropGuideAdorner.DropPosition.Above
                    : pos.Y > height * 0.67 ? DropGuideAdorner.DropPosition.Below
                    : DropGuideAdorner.DropPosition.Inside;

        ShowDropGuide(tvi, dropPos);
        e.Effects = DragDropEffects.Move;
        e.Handled = true;
    }

    private void ShapeTree_DragLeave(object sender, DragEventArgs e)
        => RemoveDropGuide();

    private void ShapeTree_Drop(object sender, DragEventArgs e)
    {
        RemoveDropGuide();
        if (!e.Data.GetDataPresent(typeof(List<PptNode>))) return;

        var tvi = GetTreeViewItem(e.OriginalSource as DependencyObject);
        if (tvi?.DataContext is not PptNode target) return;

        var draggedList = (List<PptNode>)e.Data.GetData(typeof(List<PptNode>));
        if (draggedList.Any(d => d == target || d.IsDescendant(target))) return;

        tvi.UpdateLayout();
        var pos = e.GetPosition(tvi);
        var height = tvi.ActualHeight;

        var dropPos = pos.Y < height * 0.33
            ? BckPptViewModel.DropPosition.Above
            : pos.Y > height * 0.67
                ? BckPptViewModel.DropPosition.Below
                : BckPptViewModel.DropPosition.Inside;

        // 트리 순서 유지를 위해 flat 순서로 정렬
        var flat = FlattenNodes(ViewModel.RootNodes).ToList();
        var ordered = draggedList.OrderBy(d => flat.IndexOf(d)).ToList();

        _isMovingNode = true;
        try
        {
            foreach (var dragged in ordered)
                ViewModel.MoveNode(dragged, target, dropPos);
        }
        finally { _isMovingNode = false; }

        ViewModel.SaveTreeStructure();
        e.Handled = true;
    }

    // ── 드래그 가이드라인 ─────────────────────────────────────────

    private void ShowDropGuide(TreeViewItem item, DropGuideAdorner.DropPosition position)
    {
        _adornerLayer ??= System.Windows.Documents.AdornerLayer.GetAdornerLayer(ShapeTree);
        if (_adornerLayer is null) return;

        RemoveDropGuide();

        // TreeView 기준 좌표로 변환
        var transform = item.TransformToAncestor(ShapeTree);
        var itemRect = transform.TransformBounds(
            new Rect(0, 0, item.ActualWidth, item.ActualHeight));

        _dropAdorner = new DropGuideAdorner(ShapeTree)
        {
            Position = position,
            TargetRect = itemRect,
        };
        _adornerLayer.Add(_dropAdorner);
    }

    private void RemoveDropGuide()
    {
        if (_dropAdorner is not null && _adornerLayer is not null)
        {
            _adornerLayer.Remove(_dropAdorner);
            _dropAdorner = null;
        }
    }

    // ── Tab / Shift+Tab (들여쓰기) ────────────────────────────────

    private void ShapeTree_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Tab) return;
        if (ShapeTree.SelectedItem is not PptNode node) return;

        e.Handled = true;

        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
        {
            // Shift+Tab → 부모의 형제로 이동 (outdent)
            if (node.Parent is null) return;
            var parent = node.Parent;
            var grandParentList = parent.Parent?.Children ?? ViewModel.RootNodes;
            int parentIdx = grandParentList.IndexOf(parent);

            parent.Children.Remove(node);
            grandParentList.Insert(parentIdx + 1, node);
            node.Parent = parent.Parent;
        }
        else
        {
            // Tab → 바로 위 형제의 자식으로 이동 (indent)
            var siblingList = node.Parent?.Children ?? ViewModel.RootNodes;
            int idx = siblingList.IndexOf(node);
            if (idx <= 0) return;

            var prevSibling = siblingList[idx - 1];
            siblingList.Remove(node);
            prevSibling.Children.Add(node);
            node.Parent = prevSibling;
            prevSibling.IsExpanded = true;
        }

        ViewModel.SaveTreeStructure();
    }

    // ══════════════════════════════════════════════════════════════
    //  헬퍼
    // ══════════════════════════════════════════════════════════════

    /// <summary>ContextMenu의 MenuItem에서 PptNode를 추출.</summary>
    private static PptNode? GetNodeFromMenuItem(object sender)
    {
        if (sender is not MenuItem mi) return null;
        DependencyObject? current = mi;
        ContextMenu? cm = null;
        while (current is not null)
        {
            if (current is ContextMenu found) { cm = found; break; }
            current = LogicalTreeHelper.GetParent(current)
                      ?? VisualTreeHelper.GetParent(current);
        }
        if (cm?.PlacementTarget is not FrameworkElement target) return null;
        return target.DataContext as PptNode;
    }

    /// <summary>ContextMenu의 MenuItem에서 PptSlideInfo를 추출.</summary>
    private static PptSlideInfo? GetShapeFromMenuItem(object sender)
    {
        if (sender is not MenuItem mi) return null;

        // 서브메뉴 항목인 경우 부모 MenuItem의 부모 ContextMenu를 탐색
        DependencyObject? current = mi;
        ContextMenu? cm = null;
        while (current is not null)
        {
            if (current is ContextMenu found) { cm = found; break; }
            current = LogicalTreeHelper.GetParent(current)
                      ?? VisualTreeHelper.GetParent(current);
        }
        if (cm?.PlacementTarget is not FrameworkElement target) return null;

        return target.DataContext switch
        {
            PptNode node => node.Shape,
            PptSlideInfo shape => shape,
            _ => null
        };
    }

    // ── 접기/펼치기 클릭 ─────────────────────────────────────────

    private void Expander_Click(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is PptNode node && node.Children.Count > 0)
        {
            node.IsExpanded = !node.IsExpanded;
            e.Handled = true; // 선택 이벤트 전파 방지
        }
    }

    /// <summary>모든 노드의 선택 해제.</summary>
    private void ClearAllSelections()
    {
        foreach (var node in FlattenNodes(ViewModel.RootNodes))
            node.IsSelected = false;
    }

    /// <summary>트리를 DFS로 flat 순회.</summary>
    private static IEnumerable<PptNode> FlattenNodes(IEnumerable<PptNode> nodes)
    {
        foreach (var node in nodes)
        {
            yield return node;
            foreach (var child in FlattenNodes(node.Children))
                yield return child;
        }
    }

    /// <summary>선택된 모든 노드 수집.</summary>
    private List<PptNode> GetSelectedNodes()
        => FlattenNodes(ViewModel.RootNodes).Where(n => n.IsSelected).ToList();

    /// <summary>Visual tree를 거슬러 올라가 TreeViewItem을 찾음.</summary>
    private static TreeViewItem? GetTreeViewItem(DependencyObject? obj)
    {
        while (obj is not null && obj is not TreeViewItem)
            obj = VisualTreeHelper.GetParent(obj);
        return obj as TreeViewItem;
    }

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
}
