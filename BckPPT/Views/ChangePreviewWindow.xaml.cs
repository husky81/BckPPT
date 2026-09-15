using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using BckPPT.Models;

namespace BckPPT.Views;

/// <summary>
/// 붙여넣은 내용에서 뽑아낸 변경안을 반영 전에 확인하는 창.
/// 항목별로 켜고 끌 수 있고, 줄 넘침을 미리 경고한다.
/// </summary>
public partial class ChangePreviewWindow : Window
{
    private readonly List<ParagraphChange> _changes;

    /// <summary>사용자가 선택한 변경만. 취소하면 null.</summary>
    public List<ParagraphChange>? Accepted { get; private set; }

    /// <summary>치환한 문단을 빨간색으로 표시할지.</summary>
    public bool MarkRed => MarkRedBox.IsChecked == true;

    public ChangePreviewWindow(
        List<ParagraphChange> changes,
        IReadOnlyList<string> warnings,
        IReadOnlyList<string> unmatchedLines)
    {
        InitializeComponent();
        _changes = changes;

        SourceInitialized += (_, _) =>
        {
            var hwnd  = new WindowInteropHelper(this).Handle;
            int value = 1;
            DwmSetWindowAttribute(hwnd, 20, ref value, sizeof(int));
        };

        ChangeList.ItemsSource = _changes;

        int warnCount = _changes.Count(c => c.HasWarning);
        SummaryText.Text = warnCount > 0
            ? $"변경 {_changes.Count}건 — 이 중 {warnCount}건은 줄 넘침이 예상됩니다."
            : $"변경 {_changes.Count}건";

        ApplyButton.IsEnabled = _changes.Count > 0;

        BuildWarningPanel(warnings, unmatchedLines);
    }

    private void BuildWarningPanel(IReadOnlyList<string> warnings, IReadOnlyList<string> unmatchedLines)
    {
        var sb = new StringBuilder();

        foreach (var warning in warnings)
            sb.AppendLine($"• {warning}");

        if (unmatchedLines.Count > 0)
        {
            sb.AppendLine($"• ID는 있으나 새 텍스트를 찾지 못한 줄 {unmatchedLines.Count}개:");
            foreach (var line in unmatchedLines.Take(5))
                sb.AppendLine($"    {Trim(line, 90)}");
            if (unmatchedLines.Count > 5)
                sb.AppendLine($"    … 외 {unmatchedLines.Count - 5}줄");
        }

        if (sb.Length == 0) return;
        WarningText.Text     = sb.ToString().TrimEnd();
        WarningPanel.Visibility = Visibility.Visible;
    }

    private static string Trim(string s, int max)
        => s.Length <= max ? s : s[..max] + "…";

    private void SelectAll_Changed(object sender, RoutedEventArgs e)
    {
        bool value = SelectAllBox.IsChecked == true;
        foreach (var change in _changes) change.IsSelected = value;
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        var selected = _changes.Where(c => c.IsSelected).ToList();
        if (selected.Count == 0)
        {
            MessageBox.Show("선택된 항목이 없습니다.", "변경 미리보기",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        int warnCount = selected.Count(c => c.HasWarning);
        if (warnCount > 0)
        {
            var result = MessageBox.Show(
                $"선택한 {selected.Count}건 중 {warnCount}건은 글상자를 넘칠 것으로 보입니다.\n" +
                "그대로 반영할까요?\n\n" +
                "(반영 후 AutoFit 또는 도형 높이 조정으로 보정할 수 있습니다.)",
                "줄 넘침 예상", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes) return;
        }

        Accepted     = selected;
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
}
