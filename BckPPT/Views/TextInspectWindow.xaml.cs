using System.Windows;
using BckPPT.Models;

namespace BckPPT.Views;

public partial class TextInspectWindow : Window
{
    private readonly IReadOnlyList<TextLineInfo> _lines;

    public TextInspectWindow(string shapeName, IReadOnlyList<TextLineInfo> lines)
    {
        InitializeComponent();
        _lines = lines;
        HeaderText.Text = $"{shapeName}  ({lines.Count}줄)";
        LinesList.ItemsSource = lines;
    }

    private void CopyAll_Click(object sender, RoutedEventArgs e)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var line in _lines)
        {
            sb.Append($"{line.LineNumber}: ");
            if (line.HasBullet)
                sb.Append($"{line.BulletText} ");
            sb.Append(line.Text);
            if (!string.IsNullOrEmpty(line.BreakLabel))
                sb.Append($"  [{line.BreakLabel}]");
            sb.AppendLine();
        }
        Clipboard.SetText(sb.ToString());
    }
}
